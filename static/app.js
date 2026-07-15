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

    const setProcessing = (processing) => {
        status.classList.toggle("is-visible", processing);
        submitButton.disabled = processing;
        cancelButton.hidden = !processing;
        submitButton.textContent = processing ? "Обработка..." : "Запустить обработку";
        if (!processing) stopTimer();
    };

    const setProgress = (value, message) => {
        const normalized = Math.max(0, Math.min(100, Number(value) || 0));
        progress.value = normalized;
        progress.textContent = `${normalized}%`;
        percent.value = String(normalized);
        percent.textContent = String(normalized);
        if (message) processingMessage.textContent = message;
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

        setProcessing(false);
        if (snapshot.status === "completed") {
            setProgress(100, snapshot.message);
            await loadResult(snapshot);
            return;
        }

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
        currentStatus = "queued";
        setProgress(0, "Загружаем файлы");
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
