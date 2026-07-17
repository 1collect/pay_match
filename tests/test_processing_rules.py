import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from processor.services.excel_processor import (
    _contains_excluded_keyword,
    _normalize_identifier,
)
from processor.services.name_matching import (
    match_person_name,
    match_person_name_nominative,
    nominative_person_name_tokens,
)
from processor.services.ner_service import NameExtraction
from processor.services.new_format_processor import (
    NewBankStatementProcessor,
    ProcessingCancelled,
    _detail_type_for_path,
)


class FakeNameExtractor:
    def __init__(self, names_by_text: dict[str, tuple[str, ...]]) -> None:
        self.names_by_text = names_by_text

    def extract_many(self, texts, should_cancel=None):
        return [
            NameExtraction(
                self.names_by_text.get(text, ()),
                "" if self.names_by_text.get(text) else "GLiNER 2 не обнаружил допустимое ФИО",
            )
            for text in texts
        ]


class ProcessingRuleTests(unittest.TestCase):
    def test_detail_type_supports_filename_prefixes(self) -> None:
        cases = {
            "CAFIRSTCOLLECTIONBURO_report.xlsx": "halyk",
            "astanaplat_weekend.xls": "astana",
            "20260715_anything.xlsx": "eurasian",
            "15631_report.xls": "kazpost",
            "5400-anything.xlsx": "kazpost",
            "15522.xlsx": "kazpost",
        }
        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                self.assertEqual(_detail_type_for_path(Path(filename)), expected)

        self.assertEqual(_detail_type_for_path(Path("20261340_invalid.xlsx")), "")

    def test_detail_type_keeps_legacy_filename_keywords(self) -> None:
        cases = {
            "details_astanaplat.xlsx": "astana",
            "details_eurasian.xls": "eurasian",
            "details_kazpost.xlsx": "kazpost",
            "details_halyk.xls": "halyk",
        }
        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                self.assertEqual(_detail_type_for_path(Path(filename)), expected)

    def test_numeric_exclusion_does_not_match_inside_iin(self) -> None:
        self.assertFalse(
            _contains_excluded_keyword("иин 010812600769", "108126")
        )

    def test_numeric_exclusion_matches_standalone_code(self) -> None:
        self.assertTrue(
            _contains_excluded_keyword("назначение платежа 108126", "108126")
        )

    def test_cyrillic_and_latin_dbz_letters_are_equivalent(self) -> None:
        self.assertEqual(
            _normalize_identifier("ЕМ23200-004753436"),
            _normalize_identifier("EM23200-004753436"),
        )

    def test_processor_stops_when_operation_is_cancelled(self) -> None:
        with TemporaryDirectory() as directory:
            with self.assertRaises(ProcessingCancelled):
                NewBankStatementProcessor().process_many(
                    statement_paths=[Path(directory) / "statement.xlsx"],
                    reference_path=Path(directory) / "reference.xlsx",
                    output_dir=Path(directory) / "output",
                    should_cancel=lambda: True,
                )

    def test_name_matching_is_case_insensitive_and_order_independent(self) -> None:
        evidence = match_person_name(
            "Иванов Иван Иванович",
            "иванович ИВАНОВ иВаН",
        )
        self.assertIsNotNone(evidence)
        self.assertEqual(evidence.criterion, "Полное ФИО GLiNER")

    def test_name_matching_accepts_initials_with_or_without_dots(self) -> None:
        for detected in ("Иванов И. И.", "И И ИВАНОВ", "Иванов ИИ"):
            with self.subTest(detected=detected):
                self.assertIsNotNone(
                    match_person_name("Иванов Иван Иванович", detected)
                )

    def test_name_matching_rejects_weak_abbreviations(self) -> None:
        for detected in ("Иванов", "Иванов И", "И И И", "И.И.И."):
            with self.subTest(detected=detected):
                self.assertIsNone(
                    match_person_name("Иванов Иван Иванович", detected)
                )

    def test_name_matching_allows_safe_grammatical_endings(self) -> None:
        self.assertIsNotNone(
            match_person_name(
                "Нуриева Сания Ирлановна",
                "Нуриевой Сании Ирлановны",
            )
        )

    def test_name_matching_retries_in_nominative_case(self) -> None:
        reference_name = "Петров Пётр Петрович"
        declined_name = "Петровым Петром Петровичем"

        self.assertIsNone(match_person_name(reference_name, declined_name))
        self.assertIsNotNone(
            match_person_name_nominative(reference_name, declined_name)
        )

    def test_name_matching_normalizes_declined_reference_column(self) -> None:
        self.assertIsNotNone(
            match_person_name_nominative(
                "Петровым Петром Петровичем",
                "Петров Пётр Петрович",
            )
        )

    def test_nominative_kazakh_and_feminine_names_are_not_rewritten(self) -> None:
        expected_names = {
            "Купесбаева Айнұр Бикенқызы": (
                "купесбаева",
                "айнұр",
                "бикенқызы",
            ),
            "Жанбетова Сауле Тыныштекбаевна": (
                "жанбетова",
                "сауле",
                "тыныштекбаевна",
            ),
            "СЕЙДІЛДА ЕЛДОС БӨКЕНҰЛЫ": (
                "сейділда",
                "елдос",
                "бөкенұлы",
            ),
        }
        for name, expected in expected_names.items():
            with self.subTest(name=name):
                self.assertEqual(nominative_person_name_tokens(name), expected)

    def test_real_statement_name_is_converted_from_genitive(self) -> None:
        declined_name = "Зленко Ивана Анатольевича"
        self.assertEqual(
            nominative_person_name_tokens(declined_name),
            ("зленко", "иван", "анатольевич"),
        )
        self.assertIsNone(
            match_person_name("Зленко Иван Анатольевич", declined_name)
        )
        self.assertIsNotNone(
            match_person_name_nominative(
                "Зленко Иван Анатольевич",
                declined_name,
            )
        )

    def test_every_statement_row_is_output_or_has_skip_reason(self) -> None:
        with TemporaryDirectory() as directory:
            root = Path(directory)
            statement = root / "KZ001_Компания.xlsx"
            reference = root / "reference.xlsx"
            output = root / "output"

            statement_book = Workbook()
            statement_sheet = statement_book.active
            statement_sheet.append(["Дата", "Отправитель", "Кредит", "Назначение платежа"])
            statement_sheet.append(["15.07.2026", "Плательщик", 1000, "Оплата Иван Иванович Иванов"])
            statement_sheet.append(["15.07.2026", "Плательщик", 2000, "Оплата Иванов"])
            statement_sheet.append(
                ["15.07.2026", "Плательщик", 3000, "Возврат ошибочных элементов"]
            )
            statement_sheet.append(["15.07.2026", "Плательщик", 0, "Нулевая сумма"])
            statement_book.save(statement)

            reference_book = Workbook()
            reference_sheet = reference_book.active
            reference_sheet.append(["Взыскатель", "ДБЗ", "ИИН", "ФИО"])
            reference_sheet.append(["Компания", "DBZ-1", "900101300001", "Иванов Иван Иванович"])
            reference_book.save(reference)

            extractor = FakeNameExtractor(
                {
                    "Оплата Иван Иванович Иванов": ("Иван Иванович Иванов",),
                    "Оплата Иванов": ("Иванов",),
                }
            )
            result = NewBankStatementProcessor(extractor).process_many(
                statement_paths=[statement],
                reference_path=reference,
                output_dir=output,
            )

            workbook = load_workbook(result.registry_path, read_only=True, data_only=True)
            self.assertEqual(workbook["Оплата"].max_row - 1, 1)
            self.assertEqual(workbook["Не найдено"].max_row - 1, 1)
            self.assertEqual(workbook["Журнал пропусков"].max_row - 1, 2)
            reasons = [
                row[6]
                for row in workbook["Журнал пропусков"].iter_rows(
                    min_row=2,
                    values_only=True,
                )
            ]
            self.assertTrue(all(reasons))
            workbook.close()


if __name__ == "__main__":
    unittest.main()
