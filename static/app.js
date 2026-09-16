(() => {
    const CLIENT_KEY = "paymentRegistryClientId";
    const OPERATION_KEY = "paymentRegistryOperationId";
    const ACTIVE_STATUSES = new Set(["queued", "processing", "cancelling"]);

    const form = document.getElementById("uploadForm");
    const status = document.getElementById("processingStatus");
    const timer = document.getElementById("processingTimer");
    const progress = document.getElementById("processingProgress");
    const percent = document.getElementById("processingPercent");
    const processingMessage = document.getElementById("processingMessage");
    const submitButton = document.getElementById("submitButton");
    const cancelButton = document.getElementById("cancelButton");
    const resultPanel = document.getElementById("resultPanel");
    const downloadButton = document.getElementById("downloadButton");
    const errorAlert = document.getElementById("errorAlert");

    const fileInputs = [
        {
            input: document.getElementById("statementFiles"),
            output: document.getElementById("statementSelection"),
            empty: "Файлы не выбраны",
        },
        {
            input: document.getElementById("referenceFile"),
            output: document.getElementById("referenceSelection"),
            empty: "Файл не выбран",
        },
        {
            input: document.getElementById("detailFiles"),
            output: document.getElementById("detailSelection"),
            empty: "Файлы не выбраны",
        },
    ];

    fileInputs.forEach(({ input, output, empty }) => {
        input.addEventListener("change", () => {
            const files = Array.from(input.files || []);
            input.closest(".file-field").classList.toggle("has-files", files.length > 0);
            if (!files.length) {
                output.textContent = empty;
            } else if (files.length === 1) {
                output.textContent = files[0].name;
            } else {
                output.textContent = `Выбрано файлов: ${files.length}`;
            }
        });
    });

    const createId = () => {
        if (crypto.randomUUID) return crypto.randomUUID();
        return "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (char) =>
            (char ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> char / 4).toString(16)
        );
    };

    let clientId = localStorage.getItem(CLIENT_KEY);
    if (!clientId) {
        clientId = createId();
        localStorage.setItem(CLIENT_KEY, clientId);
    }

    let operationId = sessionStorage.getItem(OPERATION_KEY);
    let currentStatus = "";
    let startedAt = null;
    let timerId = null;
    let socket = null;
    let reconnectTimer = null;
    let uploadController = null;
    let resultUrl = null;
    let resultLoading = false;
    let serverProgress = 0;
    let displayedProgress = 0;
    let progressFrame = null;
    let progressFrameTime = null;
    let processingStartedAt = null;
    let lastRenderedPercent = -1;
    let playfulMessageIndex = -1;

    const playfulMessages = [
        "Разминаем формулы",
        "Уговариваем Excel не чудить",
        "Сверяем дебет с суровой реальностью",
        "Ищем те самые две копейки",
        "Делаем вид, что это очень сложно",
        "Наводим финальный лоск",
    ];

    const renderTime = () => {
        if (!startedAt) return;
        const elapsedSeconds = Math.max(0, Math.floor((Date.now() - startedAt) / 1000));
        const minutes = String(Math.floor(elapsedSeconds / 60)).padStart(2, "0");
        const seconds = String(elapsedSeconds % 60).padStart(2, "0");
        timer.value = `${minutes}:${seconds}`;
        timer.textContent = `${minutes}:${seconds}`;
    };

    const startTimer = () => {
        if (timerId) window.clearInterval(timerId);
        renderTime();
        timerId = window.setInterval(renderTime, 1000);
    };

    const stopTimer = () => {
        if (timerId) window.clearInterval(timerId);
        timerId = null;
    };

    const fakeProgressFor = (elapsed) => {
        if (elapsed < 1600) return 4 + (elapsed / 1600) * 12;
        if (elapsed < 5200) return 16 + ((elapsed - 1600) / 3600) * 25;
        if (elapsed < 11000) return 41 + ((elapsed - 5200) / 5800) * 25;
        if (elapsed < 21000) return 66 + ((elapsed - 11000) / 10000) * 17;
        return 83 + 11 * (1 - Math.exp(-(elapsed - 21000) / 18000));
    };

    const renderDisplayedProgress = () => {
        progress.value = displayedProgress;
        progress.textContent = `${Math.round(displayedProgress)}%`;
        const rounded = Math.round(displayedProgress);
        if (rounded !== lastRenderedPercent) {
            percent.value = String(rounded);
            percent.textContent = String(rounded);
            lastRenderedPercent = rounded;
        }
    };

    const animateProgress = (now) => {
        if (!status.classList.contains("is-visible")) {
            progressFrame = null;
            progressFrameTime = null;
            return;
        }

        const elapsed = now - processingStartedAt;
        const fakeTarget = fakeProgressFor(elapsed);
        const target = Math.min(99, Math.max(serverProgress, fakeTarget));
        const delta = Math.min(50, now - (progressFrameTime || now));
        const ease = 1 - Math.exp(-delta / 520);
        displayedProgress += (target - displayedProgress) * ease;

        if (target - displayedProgress < 0.025) displayedProgress = target;
        renderDisplayedProgress();

        const messageIndex = Math.min(playfulMessages.length - 1, Math.floor(elapsed / 3300));
        if (currentStatus !== "cancelling" && messageIndex !== playfulMessageIndex) {
            playfulMessageIndex = messageIndex;
            processingMessage.textContent = playfulMessages[messageIndex];
            processingMessage.animate(
                [
                    { opacity: 0.25, transform: "translateY(3px)" },
                    { opacity: 1, transform: "translateY(0)" },
                ],
                { duration: 380, easing: "cubic-bezier(.2,.8,.2,1)" }
            );
        }
        progressFrameTime = now;
        progressFrame = window.requestAnimationFrame(animateProgress);
    };

    const startProgressAnimation = () => {
        if (!processingStartedAt) processingStartedAt = performance.now();
        if (!progressFrame) progressFrame = window.requestAnimationFrame(animateProgress);
    };

    const stopProgressAnimation = () => {
        if (progressFrame) window.cancelAnimationFrame(progressFrame);
        progressFrame = null;
        progressFrameTime = null;
    };

    const completeProgress = (message) => new Promise((resolve) => {
        stopProgressAnimation();
        serverProgress = 100;
        processingMessage.textContent = message || "Готово. Excel выжил";

        if (
            !status.classList.contains("is-visible") ||
            window.matchMedia("(prefers-reduced-motion: reduce)").matches
        ) {
            displayedProgress = 100;
            renderDisplayedProgress();
            resolve();
            return;
        }

        const from = displayedProgress;
        const started = performance.now();
        const finishFrame = (now) => {
            const t = Math.min(1, (now - started) / 620);
            const eased = 1 - Math.pow(1 - t, 3);
            displayedProgress = from + (100 - from) * eased;
            renderDisplayedProgress();
            if (t < 1) {
                window.requestAnimationFrame(finishFrame);
            } else {
                resolve();
            }
        };
        window.requestAnimationFrame(finishFrame);
    });

    const setProcessing = (processing) => {
        status.classList.toggle("is-visible", processing);
        submitButton.disabled = processing;
        cancelButton.hidden = !processing;
        submitButton.querySelector("span").textContent = processing
            ? "Обработка..."
            : "Запустить обработку";
        if (processing) {
            startProgressAnimation();
        } else {
            stopProgressAnimation();
            stopTimer();
        }
    };

    const setProgress = (value, message, immediate = false) => {
        const normalized = Math.max(0, Math.min(100, Number(value) || 0));
        serverProgress = normalized;
        if (immediate) {
            displayedProgress = normalized;
            renderDisplayedProgress();
        }
        if (message && (immediate || currentStatus === "cancelling")) {
            processingMessage.textContent = message;
        }
    };

    const showError = (message) => {
        errorAlert.textContent = message;
        errorAlert.hidden = false;
    };

    const clearError = () => {
        errorAlert.textContent = "";
        errorAlert.hidden = true;
    };

    const clearResult = () => {
        if (resultUrl) URL.revokeObjectURL(resultUrl);
        resultUrl = null;
        resultPanel.hidden = true;
    };

    const renderStats = (stats = {}) => {
        Object.entries(stats).forEach(([name, value]) => {
            const element = resultPanel.querySelector(`[data-stat="${name}"]`);
            if (element) element.textContent = String(value);
        });
    };

    const resultEndpoint = () =>
        `/operations/${encodeURIComponent(operationId)}/result?client_id=${encodeURIComponent(clientId)}`;

    const loadResult = async (snapshot) => {
        if (resultLoading || resultUrl || !operationId) return;
        resultLoading = true;
        try {
            const response = await fetch(resultEndpoint(), { cache: "no-store" });
            if (!response.ok) {
                const payload = await response.json().catch(() => ({}));
                throw new Error(payload.error || "Не удалось получить готовый файл.");
            }
            resultUrl = URL.createObjectURL(await response.blob());
            renderStats(snapshot.stats);
            resultPanel.hidden = false;
            resultPanel.scrollIntoView({ behavior: "smooth", block: "start" });
        } catch (error) {
            showError(error.message);
        } finally {
            resultLoading = false;
        }
    };

    const clearOperation = () => {
        sessionStorage.removeItem(OPERATION_KEY);
        operationId = null;
        currentStatus = "";
    };

    const applySnapshot = async (snapshot) => {
        if (!snapshot || snapshot.operation_id !== operationId) return;
        currentStatus = snapshot.status;
        startedAt = (snapshot.started_at || snapshot.created_at) * 1000;

        if (ACTIVE_STATUSES.has(snapshot.status)) {
            setProcessing(true);
            setProgress(snapshot.progress, snapshot.message);
            startTimer();
            return;
        }

        if (snapshot.status === "completed") {
            await completeProgress(snapshot.message);
            setProcessing(false);
            await loadResult(snapshot);
            return;
        }

        setProcessing(false);
        if (snapshot.status === "cancelled") {
            showError("Операция отменена.");
        } else if (snapshot.status === "error") {
            showError(snapshot.error || "Не удалось обработать файлы.");
        }
        clearOperation();
    };

    const websocketUrl = () => {
        const protocol = location.protocol === "https:" ? "wss:" : "ws:";
        const query = new URLSearchParams({ client_id: clientId, operation_id: operationId });
        return `${protocol}//${location.host}/ws?${query}`;
    };

    const connectSocket = () => {
        if (!operationId || !ACTIVE_STATUSES.has(currentStatus)) return;
        if (socket) socket.close();
        socket = new WebSocket(websocketUrl());

        socket.addEventListener("message", (event) => {
            const snapshot = JSON.parse(event.data);
            applySnapshot(snapshot);
        });

        socket.addEventListener("close", () => {
            socket = null;
            if (!operationId || !ACTIVE_STATUSES.has(currentStatus)) return;
            window.clearTimeout(reconnectTimer);
            reconnectTimer = window.setTimeout(refreshOperation, 1500);
        });
    };

    const refreshOperation = async () => {
        if (!operationId) return;
        try {
            const query = new URLSearchParams({ client_id: clientId });
            const response = await fetch(`/operations/${encodeURIComponent(operationId)}?${query}`, {
                cache: "no-store",
            });
            if (!response.ok) {
                clearOperation();
                setProcessing(false);
                return;
            }
            const snapshot = await response.json();
            await applySnapshot(snapshot);
            connectSocket();
        } catch (_error) {
            reconnectTimer = window.setTimeout(refreshOperation, 2000);
        }
    };

    form.addEventListener("submit", async (event) => {
        event.preventDefault();
        if (socket) socket.close();
        clearOperation();
        clearError();
        clearResult();
        startedAt = Date.now();
        processingStartedAt = performance.now();
        serverProgress = 0;
        displayedProgress = 0;
        lastRenderedPercent = -1;
        playfulMessageIndex = -1;
        currentStatus = "queued";
        setProgress(0, "Загружаем файлы", true);
        setProcessing(true);
        startTimer();

        uploadController = new AbortController();
        const formData = new FormData(form);
        formData.append("client_id", clientId);

        try {
            const response = await fetch("/operations", {
                method: "POST",
                body: formData,
                signal: uploadController.signal,
            });
            if (!response.ok) {
                const payload = await response.json().catch(() => ({}));
                throw new Error(payload.error || "Не удалось запустить обработку.");
            }
            const snapshot = await response.json();
            operationId = snapshot.operation_id;
            sessionStorage.setItem(OPERATION_KEY, operationId);
            await applySnapshot(snapshot);
            connectSocket();
        } catch (error) {
            setProcessing(false);
            if (error.name !== "AbortError") showError(error.message);
        } finally {
            uploadController = null;
        }
    });

    cancelButton.addEventListener("click", async () => {
        cancelButton.disabled = true;
        try {
            if (uploadController && !operationId) {
                uploadController.abort();
                showError("Загрузка отменена.");
                setProcessing(false);
                return;
            }
            if (!operationId) return;
            const response = await fetch(`/operations/${encodeURIComponent(operationId)}/cancel`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ client_id: clientId }),
            });
            const snapshot = await response.json();
            if (!response.ok) throw new Error(snapshot.error || "Не удалось отменить операцию.");
            await applySnapshot(snapshot);
        } catch (error) {
            showError(error.message);
        } finally {
            cancelButton.disabled = false;
        }
    });

    downloadButton.addEventListener("click", () => {
        if (!resultUrl) return;
        const link = document.createElement("a");
        link.href = resultUrl;
        link.download = "registry.xlsx";
        link.click();
    });

    window.addEventListener("beforeunload", () => {
        if (socket) socket.close();
        if (resultUrl) URL.revokeObjectURL(resultUrl);
    });

    if (operationId) refreshOperation();
})();
