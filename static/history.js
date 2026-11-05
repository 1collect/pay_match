(() => {
    const CLIENT_KEY = "paymentRegistryClientId";
    const ACTIVE_STATUSES = new Set(["queued", "processing", "cancelling"]);
    const STATUS_LABELS = {
        queued: "В очереди",
        processing: "В работе",
        cancelling: "Отмена",
        completed: "Завершена",
        cancelled: "Отменена",
        error: "Ошибка",
        interrupted: "Прервана",
    };
    const STATUS_CLASSES = {
        queued: "queued",
        processing: "processing",
        cancelling: "processing",
        completed: "completed",
        cancelled: "cancelled",
        error: "error",
        interrupted: "interrupted",
    };

    const refreshButton = document.getElementById("refreshHistory");
    const retryButton = document.getElementById("retryHistory");
    const loading = document.getElementById("historyLoading");
    const error = document.getElementById("historyError");
    const errorMessage = document.getElementById("historyErrorMessage");
    const empty = document.getElementById("historyEmpty");
    const emptyTitle = document.getElementById("historyEmptyTitle");
    const emptyMessage = document.getElementById("historyEmptyMessage");
    const overview = document.getElementById("historyOverview");
    const count = document.getElementById("historyCount");
    const updated = document.getElementById("historyUpdated");
    const toolbar = document.getElementById("historyToolbar");
    const search = document.getElementById("historySearch");
    const statusFilter = document.getElementById("historyStatus");
    const resultCount = document.getElementById("historyResultCount");
    const tableScroll = document.getElementById("historyTableScroll");
    const rows = document.getElementById("historyRows");

    let historyItems = [];
    let refreshTimer = null;
    let requestPending = false;
    let loadedAt = null;

    const pluralize = (value, forms) => {
        const mod100 = value % 100;
        const mod10 = value % 10;
        if (mod100 >= 11 && mod100 <= 14) return forms[2];
        if (mod10 === 1) return forms[0];
        if (mod10 >= 2 && mod10 <= 4) return forms[1];
        return forms[2];
    };

    const numberFormat = new Intl.NumberFormat("ru-RU");
    const dateFormat = new Intl.DateTimeFormat("ru-RU", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
    });
    const timeFormat = new Intl.DateTimeFormat("ru-RU", {
        hour: "2-digit",
        minute: "2-digit",
    });

    const safeTimestamp = (value) => {
        const timestamp = Number(value);
        return Number.isFinite(timestamp) && timestamp > 0 ? timestamp * 1000 : null;
    };

    const formatDuration = (item) => {
        const start = safeTimestamp(item.started_at || item.created_at);
        const end = safeTimestamp(item.finished_at) || Date.now();
        if (!start) return "";
        const seconds = Math.max(0, Math.floor((end - start) / 1000));
        if (seconds < 60) return `${seconds} с`;
        const minutes = Math.floor(seconds / 60);
        const remainder = seconds % 60;
        return remainder ? `${minutes} мин ${remainder} с` : `${minutes} мин`;
    };

    const clearRows = () => {
        rows.replaceChildren();
    };

    const appendCell = (row, className = "") => {
        const cell = document.createElement("td");
        if (className) cell.className = className;
        row.append(cell);
        return cell;
    };

    const allFileEntries = (item) => [
        ...(Array.isArray(item.statement_files)
            ? item.statement_files.map((filename) => ({ filename, kind: "Выписка" }))
            : []),
        ...(typeof item.reference_file === "string" && item.reference_file
            ? [{ filename: item.reference_file, kind: "Справочник" }]
            : []),
        ...(Array.isArray(item.detail_files)
            ? item.detail_files.map((filename) => ({ filename, kind: "Расшифровка" }))
            : []),
    ].filter((entry) => typeof entry.filename === "string" && entry.filename.length > 0);

    const allFileNames = (item) => allFileEntries(item).map((entry) => entry.filename);

    const appendFiles = (cell, item) => {
        const files = allFileEntries(item);
        const details = document.createElement("details");
        details.className = "history-files";
        const summary = document.createElement("summary");
        summary.textContent = files[0]?.filename || "Файлы не указаны";
        if (files.length > 1) {
            const extra = document.createElement("span");
            extra.className = "history-file-count";
            extra.textContent = `+ ещё ${files.length - 1}`;
            summary.append(extra);
        }
        details.append(summary);

        if (files.length) {
            const list = document.createElement("ul");
            list.className = "history-file-list";
            files.forEach(({ filename, kind: fileKind }) => {
                const listItem = document.createElement("li");
                const kind = document.createElement("span");
                kind.className = "history-file-kind";
                kind.textContent = `${fileKind} · `;
                listItem.append(kind, document.createTextNode(filename));
                list.append(listItem);
            });
            details.append(list);
        }

        if (item.error) {
            const message = document.createElement("p");
            message.className = "history-operation-error";
            message.textContent = item.error;
            details.append(message);
        }
        cell.append(details);
    };

    const appendRow = (item) => {
        const row = document.createElement("tr");

        const dateCell = appendCell(row);
        const timestamp = safeTimestamp(item.created_at);
        if (timestamp) {
            const time = document.createElement("time");
            time.dateTime = new Date(timestamp).toISOString();
            time.textContent = dateFormat.format(timestamp);
            dateCell.append(time);
        } else {
            dateCell.textContent = "—";
        }
        const durationText = formatDuration(item);
        if (durationText) {
            const duration = document.createElement("span");
            duration.className = "history-duration";
            duration.textContent = durationText;
            dateCell.append(duration);
        }

        const statusCell = appendCell(row);
        const status = document.createElement("span");
        status.className = `history-status history-status--${STATUS_CLASSES[item.status] || "cancelled"}`;
        status.textContent = STATUS_LABELS[item.status] || "Неизвестно";
        statusCell.append(status);

        appendFiles(appendCell(row), item);

        const stats = item.stats || {};
        [stats.total, stats.matched, stats.notFound].forEach((value) => {
            const cell = appendCell(row, "history-number");
            cell.textContent = Number.isFinite(Number(value))
                ? numberFormat.format(Number(value))
                : "—";
        });
        rows.append(row);
    };

    const matchesFilters = (item) => {
        const selectedStatus = statusFilter.value;
        if (selectedStatus !== "all" && item.status !== selectedStatus) return false;
        const query = search.value.trim().toLocaleLowerCase("ru-RU");
        if (!query) return true;
        return allFileNames(item).some((filename) => filename.toLocaleLowerCase("ru-RU").includes(query));
    };

    const render = () => {
        clearRows();
        const visibleItems = historyItems.filter(matchesFilters);
        overview.hidden = historyItems.length === 0;
        toolbar.hidden = historyItems.length === 0;
        tableScroll.hidden = visibleItems.length === 0;
        empty.hidden = historyItems.length > 0 && visibleItems.length > 0;

        if (historyItems.length) {
            count.textContent = `${numberFormat.format(historyItems.length)} ${pluralize(historyItems.length, ["операция", "операции", "операций"])}`;
            updated.textContent = loadedAt ? `Обновлено в ${timeFormat.format(loadedAt)}` : "";
            resultCount.textContent = visibleItems.length === historyItems.length
                ? `Показано: ${numberFormat.format(visibleItems.length)}`
                : `Показано: ${numberFormat.format(visibleItems.length)} из ${numberFormat.format(historyItems.length)}`;
        }

        if (!historyItems.length) {
            emptyTitle.textContent = "Истории пока нет";
            emptyMessage.textContent = "Записи обработок за последние 30 дней появятся здесь.";
        } else if (!visibleItems.length) {
            emptyTitle.textContent = "Нет подходящих операций";
            emptyMessage.textContent = "Измените запрос или выберите другой статус.";
        }

        visibleItems.forEach(appendRow);
    };

    const setBusy = (busy) => {
        refreshButton.disabled = busy;
        retryButton.disabled = busy;
        refreshButton.setAttribute("aria-busy", String(busy));
        refreshButton.querySelector("span").textContent = busy ? "Обновляем" : "Обновить";
    };

    const loadHistory = async () => {
        if (requestPending) return;
        requestPending = true;
        window.clearTimeout(refreshTimer);
        const isFirstLoad = loadedAt === null;
        if (isFirstLoad) loading.hidden = false;
        error.hidden = true;
        setBusy(true);

        try {
            let clientId = null;
            try {
                clientId = localStorage.getItem(CLIENT_KEY);
            } catch (_storageError) {
                clientId = null;
            }
            if (!clientId) {
                historyItems = [];
            } else {
                const query = new URLSearchParams({ client_id: clientId });
                const response = await fetch(`/api/history?${query}`, { cache: "no-store" });
                const payload = await response.json().catch(() => ({}));
                if (!response.ok) throw new Error(payload.error || "Проверьте соединение и попробуйте снова.");
                historyItems = Array.isArray(payload.items) ? payload.items : [];
            }
            loadedAt = new Date();
            render();
        } catch (requestError) {
            errorMessage.textContent = requestError.message || "Проверьте соединение и попробуйте снова.";
            error.hidden = false;
            if (isFirstLoad) empty.hidden = true;
        } finally {
            requestPending = false;
            loading.hidden = true;
            setBusy(false);
        }

        if (historyItems.some((item) => ACTIVE_STATUSES.has(item.status))) {
            refreshTimer = window.setTimeout(loadHistory, 12000);
        }
    };

    refreshButton.addEventListener("click", loadHistory);
    retryButton.addEventListener("click", loadHistory);
    search.addEventListener("input", render);
    statusFilter.addEventListener("change", render);
    window.addEventListener("storage", (event) => {
        if (event.key === CLIENT_KEY) loadHistory();
    });

    loadHistory();
})();
