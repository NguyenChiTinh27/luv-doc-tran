/**
 * LUV-DOC-TRAN — Frontend Architecture Controller
 * ─────────────────────────────────────────────────────────────
 * Hệ thống Dịch thuật Song ngữ Kỹ thuật Chuyên Nghiệp (DOC-TRAN PRO)
 * Hỗ trợ Tải nhiều tệp (Multi-file Batch Queue) • Tải về tức thì từng tệp
 * Review Split-View Studio • Quản lý Bản Quyền & Dùng thử 48h
 */

window.LuvApp = {
    fileQueue: [],
    activeFileId: null,
    currentViewingResults: null,
    deviceId: null,
    trialRemainingSeconds: 0,
    trialTimerInterval: null,
    syncScrollEnabled: true,
    currentZoom: 100,
    isInlineEditing: false,
    isProcessingBatch: false,

    init: async function() {
        this.initDeviceId();
        this.setupDragAndDrop();
        await this.checkLicenseAndTrial();
        this.startTrialCountdown();
    },

    // ── 1. DEVICE IDENTIFIER (ANTI-TAMPER 48H FINGERPRINT) ────────────────
    initDeviceId: function() {
        let devId = localStorage.getItem("luv_device_id");
        if (!devId) {
            const raw = `${navigator.userAgent}_${screen.width}x${screen.height}_${new Date().getTimezoneOffset()}_${Math.random()}`;
            devId = "DEV-" + Math.abs(this.hashCode(raw)).toString(16).toUpperCase();
            localStorage.setItem("luv_device_id", devId);
        }
        this.deviceId = devId;
    },

    hashCode: function(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            hash = ((hash << 5) - hash) + str.charCodeAt(i);
            hash |= 0;
        }
        return hash;
    },

    // ── 2. LICENSE & 48H TRIAL STATUS ────────────────────────────────────
    checkLicenseAndTrial: async function() {
        const trialPill = document.getElementById("trial-pill");
        const trialText = document.getElementById("trial-timer-text");
        const licLabel = document.getElementById("license-header-label");

        try {
            // A. Check License status
            const licResp = await fetch("/api/license/status");
            const licData = await licResp.json();

            if (licData.is_active) {
                if (trialPill) trialPill.className = "badge-pill trial-badge trial-active";
                if (trialText) trialText.textContent = `Bản Quyền: ${licData.tier || 'Vĩnh Viễn'}`;
                if (licLabel) licLabel.textContent = "✓ Bản Quyền Đã Kích Hoạt";
                return;
            }

            // B. Check 48h Trial
            const trialResp = await fetch(`/api/trial/status?device_id=${encodeURIComponent(this.deviceId)}`);
            const trialData = await trialResp.json();

            this.trialRemainingSeconds = trialData.remaining_seconds || 0;

            if (trialData.is_trial_active && this.trialRemainingSeconds > 0) {
                if (trialPill) trialPill.className = "badge-pill trial-badge trial-active";
                this.updateTrialDisplay();
            } else {
                if (trialPill) trialPill.className = "badge-pill trial-badge trial-expired";
                if (trialText) trialText.textContent = "Hết Hạn Trải Nghiệm (Cần Key)";
            }
        } catch (err) {
            console.error("Lỗi kiểm tra bản quyền / dùng thử:", err);
        }
    },

    startTrialCountdown: function() {
        if (this.trialTimerInterval) clearInterval(this.trialTimerInterval);
        this.trialTimerInterval = setInterval(() => {
            if (this.trialRemainingSeconds > 0) {
                this.trialRemainingSeconds--;
                this.updateTrialDisplay();
            }
        }, 1000);
    },

    updateTrialDisplay: function() {
        const trialText = document.getElementById("trial-timer-text");
        if (!trialText) return;

        const sec = this.trialRemainingSeconds;
        const h = Math.floor(sec / 3600);
        const m = Math.floor((sec % 3600) / 60);
        const s = sec % 60;

        if (h > 0) {
            trialText.textContent = `Dùng thử: ${h}h ${m}m`;
        } else if (m > 0) {
            trialText.textContent = `Dùng thử: ${m}m ${s}s`;
        } else {
            trialText.textContent = `Hết hạn dùng thử`;
            const pill = document.getElementById("trial-pill");
            if (pill) pill.className = "badge-pill trial-badge trial-expired";
        }
    },

    // ── 3. DRAG AND DROP & MULTI-FILE QUEUE ──────────────────────────────
    setupDragAndDrop: function() {
        const dropzone = document.getElementById("dropzone");
        if (!dropzone) return;

        ['dragenter', 'dragover'].forEach(name => {
            dropzone.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add("dragover");
            });
        });

        ['dragleave', 'drop'].forEach(name => {
            dropzone.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.remove("dragover");
            });
        });

        dropzone.addEventListener('drop', (e) => {
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                this.handleFileSelect(e.dataTransfer.files);
            }
        });
    },

    handleFileSelect: function(files) {
        if (!files || files.length === 0) return;

        let addedCount = 0;
        const allowedExts = ['.docx', '.doc', '.pdf', '.xlsx', '.xls'];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

            if (!allowedExts.includes(ext)) {
                alert(`Tệp "${file.name}" không hợp lệ. Vui lòng chọn tệp Word (.docx, .doc), PDF (.pdf) hoặc Excel (.xlsx, .xls).`);
                continue;
            }

            // Generate unique item ID
            const itemId = "file_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
            this.fileQueue.push({
                id: itemId,
                file: file,
                name: file.name,
                size: file.size,
                ext: ext,
                status: 'waiting',      // 'waiting', 'translating', 'ready', 'error'
                statusText: 'Chờ xử lý',
                results: null,
                errorMsg: null
            });
            addedCount++;
        }

        // Reset input element
        const fileInput = document.getElementById("file-input");
        if (fileInput) fileInput.value = "";

        if (addedCount > 0) {
            this.renderQueueTable();
        }
    },

    renderQueueTable: function() {
        const container = document.getElementById("queue-container");
        const tbody = document.getElementById("queue-tbody");
        const countBadge = document.getElementById("queue-count-badge");
        const summaryText = document.getElementById("queue-summary-text");

        if (!container || !tbody) return;

        if (this.fileQueue.length === 0) {
            container.style.display = "none";
            return;
        }

        container.style.display = "block";
        if (countBadge) countBadge.textContent = `${this.fileQueue.length} tệp`;

        const completedCount = this.fileQueue.filter(f => f.status === 'ready').length;
        if (summaryText) {
            if (completedCount === 0) {
                summaryText.textContent = `Đã chọn ${this.fileQueue.length} tài liệu. Nhấn "Bắt đầu dịch tất cả" để tiến hành.`;
            } else {
                summaryText.textContent = `Đã hoàn thành ${completedCount}/${this.fileQueue.length} tài liệu. Bạn có thể tải ngay từng tệp bên trên!`;
            }
        }

        tbody.innerHTML = this.fileQueue.map((item, index) => {
            let fmtBadgeClass = 'fmt-word';
            let fmtLabel = 'WORD';
            if (item.ext === '.pdf') {
                fmtBadgeClass = 'fmt-pdf';
                fmtLabel = 'PDF';
            } else if (['.xlsx', '.xls'].includes(item.ext)) {
                fmtBadgeClass = 'fmt-excel';
                fmtLabel = 'EXCEL';
            }

            // Status badge rendering
            let statusHtml = '';
            if (item.status === 'waiting') {
                statusHtml = `<span class="status-badge status-waiting">⏳ Chờ xử lý</span>`;
            } else if (item.status === 'translating') {
                statusHtml = `<span class="status-badge status-processing"><span class="spinner-mini"></span> ${item.statusText || 'Đang dịch...'}</span>`;
            } else if (item.status === 'ready') {
                statusHtml = `<span class="status-badge status-complete">✓ Hoàn thành (${item.elapsed || 0}s)</span>`;
            } else if (item.status === 'error') {
                statusHtml = `<span class="status-badge status-error" title="${item.errorMsg || ''}">✕ Lỗi xử lý</span>`;
            }

            // Action buttons rendering:
            // "XONG FILE NÀO THÌ CÓ THỂ CHO PHÉP TẢI FILE ĐÓ VỀ NGAY LẬP TỨC"
            let actionsHtml = '';
            if (item.status === 'ready' && item.results) {
                const res = item.results;
                if (item.ext === '.xlsx' || item.ext === '.xls') {
                    actionsHtml = `
                        <div class="row-actions-group">
                            <a href="${res.download_url}" class="btn-row-action btn-row-word" download>
                                📊 Tải Excel (.xlsx)
                            </a>
                            <button type="button" class="btn-row-action btn-row-view" onclick="window.LuvApp.viewSplitStudio('${item.id}')">
                                🔍 Xem So Sánh
                            </button>
                        </div>
                    `;
                } else {
                    actionsHtml = `
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-row-word" onclick="window.LuvApp.downloadFileWord('${item.id}')" title="Tải Word xử lý tức thì">
                                ⚡ Tải Word (.docx)
                            </button>
                            <button type="button" class="btn-row-action btn-row-pdf" onclick="window.LuvApp.handleFilePdfDownload('${item.id}')" title="Tải PDF vector chuẩn in">
                                📄 Tải PDF
                            </button>
                            <button type="button" class="btn-row-action btn-row-view" onclick="window.LuvApp.viewSplitStudio('${item.id}')" title="Xem trước tài liệu">
                                🔍 So Sánh
                            </button>
                        </div>
                    `;
                }
            } else if (item.status === 'translating') {
                actionsHtml = `<span style="font-size:0.75rem; color:#94a3b8;">Đang xử lý tài liệu...</span>`;
            } else {
                actionsHtml = `
                    <div class="row-actions-group">
                        <button type="button" class="btn-row-delete" onclick="window.LuvApp.removeQueueItem('${item.id}')" title="Xóa khỏi hàng đợi">
                            ✕
                        </button>
                    </div>
                `;
            }

            return `
                <tr class="queue-row" id="row-${item.id}">
                    <td style="color:#64748b; font-family:var(--font-mono);">${index + 1}</td>
                    <td>
                        <div class="file-name-cell">
                            <span class="file-cell-icon">${this.getFileIconSvg(item.ext)}</span>
                            <span>${this.escapeHtml(item.name)}</span>
                        </div>
                    </td>
                    <td><span class="file-fmt-badge ${fmtBadgeClass}">${fmtLabel}</span></td>
                    <td style="color:#94a3b8; font-family:var(--font-mono); font-size:0.8rem;">${this.formatFileSize(item.size)}</td>
                    <td>${statusHtml}</td>
                    <td>${actionsHtml}</td>
                </tr>
            `;
        }).join('');
    },

    removeQueueItem: function(itemId) {
        if (this.isProcessingBatch) return;
        this.fileQueue = this.fileQueue.filter(f => f.id !== itemId);
        this.renderQueueTable();
    },

    clearAllQueue: function() {
        if (this.isProcessingBatch) return;
        this.fileQueue = [];
        this.renderQueueTable();
        const studio = document.getElementById("split-view-studio");
        if (studio) studio.style.display = "none";
    },

    getFileIconSvg: function(ext) {
        if (['.docx', '.doc'].includes(ext)) {
            return `<svg viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2" style="width:16px;height:16px;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`;
        } else if (ext === '.pdf') {
            return `<svg viewBox="0 0 24 24" fill="none" stroke="#f87171" stroke-width="2" style="width:16px;height:16px;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`;
        } else {
            return `<svg viewBox="0 0 24 24" fill="none" stroke="#34d399" stroke-width="2" style="width:16px;height:16px;"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/></svg>`;
        }
    },

    formatFileSize: function(bytes) {
        if (!bytes) return "0 B";
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / 1048576).toFixed(2) + " MB";
    },

    escapeHtml: function(str) {
        return (str || '').replace(/[&<>"']/g, function(m) {
            return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m];
        });
    },

    // ── 4. BATCH PROCESSING PIPELINE ─────────────────────────────────────
    startBatchProcessing: async function() {
        if (this.fileQueue.length === 0) {
            alert("Vui lòng tải lên ít nhất một tệp tài liệu trước khi bắt đầu.");
            return;
        }

        const pendingItems = this.fileQueue.filter(f => f.status === 'waiting' || f.status === 'error');
        if (pendingItems.length === 0) {
            alert("Tất cả tài liệu trong hàng đợi đã được xử lý xong!");
            return;
        }

        this.isProcessingBatch = true;
        const btnBatch = document.getElementById("btn-start-batch");
        const btnText = document.getElementById("btn-batch-text");
        if (btnBatch) btnBatch.disabled = true;

        const sourceLang = document.getElementById("sel-source-lang")?.value || "auto";
        const engine = document.getElementById("sel-engine")?.value || "deepseek";

        for (let i = 0; i < pendingItems.length; i++) {
            const item = pendingItems[i];
            if (btnText) btnText.textContent = `ĐANG DỊCH TỆP ${i + 1}/${pendingItems.length}...`;

            item.status = 'translating';
            item.statusText = 'Đang phân tích bố cục...';
            this.renderQueueTable();

            try {
                const formData = new FormData();
                formData.append("file", item.file);
                formData.append("source_lang", sourceLang);
                formData.append("target_lang", "vi");
                formData.append("provider", engine === "deepseek" ? "deepseek" : "google");
                formData.append("engine_mode", "auto");
                formData.append("output_mode", "bilingual");
                formData.append("device_id", this.deviceId);

                // Simulation updater for file item
                let simPercent = 20;
                const simInterval = setInterval(() => {
                    if (item.status === 'translating' && simPercent < 85) {
                        simPercent += 15;
                        item.statusText = `Đang dịch liên dòng... (${simPercent}%)`;
                        this.renderQueueTable();
                    }
                }, 900);

                const resp = await fetch("/api/translate", {
                    method: "POST",
                    body: formData
                });

                clearInterval(simInterval);

                if (resp.status === 403) {
                    const errData = await resp.json();
                    item.status = 'error';
                    item.errorMsg = errData.message || "Hết hạn dùng thử!";
                    this.renderQueueTable();
                    alert(errData.message || "Thời hạn dùng thử 48h đã kết thúc. Vui lòng kích hoạt bản quyền để tiếp tục.");
                    this.openLicenseModal();
                    break;
                }

                const res = await resp.json();
                if (!resp.ok || !res.success) {
                    throw new Error(res.detail || "Không thể dịch tài liệu.");
                }

                // Item completed successfully!
                item.status = 'ready';
                item.results = res;
                item.elapsed = res.elapsed_seconds || 3;
                this.renderQueueTable();

                // If this is the first completed file or current view is empty, open Split-View
                if (!this.currentViewingResults) {
                    this.viewSplitStudio(item.id);
                }

            } catch (err) {
                console.error(`Lỗi dịch tệp ${item.name}:`, err);
                item.status = 'error';
                item.errorMsg = err.message;
                this.renderQueueTable();
            }
        }

        this.isProcessingBatch = false;
        if (btnBatch) btnBatch.disabled = false;
        if (btnText) btnText.textContent = "BẮT ĐẦU DỊCH TẤT CẢ TỆP";
    },

    // ── 5. SPLIT-VIEW STUDIO CONTROLS ────────────────────────────────────
    viewSplitStudio: async function(itemId) {
        const item = this.fileQueue.find(f => f.id === itemId);
        if (!item || !item.results) return;

        this.activeFileId = itemId;
        this.currentViewingResults = item.results;

        const studio = document.getElementById("split-view-studio");
        const hudName = document.getElementById("hud-doc-name");
        const paneOrig = document.getElementById("pane-original");
        const paneTran = document.getElementById("pane-translated");

        if (hudName) hudName.textContent = item.name;
        if (studio) {
            studio.style.display = "block";
            studio.scrollIntoView({ behavior: "smooth", block: "start" });
        }

        // 1. Load original preview
        if (paneOrig) {
            paneOrig.innerHTML = `<div class="pane-loading">Đang tải bản xem trước tài liệu nguồn...</div>`;
            try {
                const fd = new FormData();
                fd.append("file", item.file);
                const resp = await fetch("/api/preview/upload", { method: "POST", body: fd });
                const res = await resp.json();
                if (res.success && res.html) {
                    paneOrig.innerHTML = res.html;
                    this.setupSyncScroll();
                } else {
                    paneOrig.innerHTML = `<div class="pane-loading">Không thể hiển thị xem trước tệp nguồn.</div>`;
                }
            } catch (err) {
                paneOrig.innerHTML = `<div class="pane-loading" style="color:#f87171;">Lỗi tải xem trước: ${err.message}</div>`;
            }
        }

        // 2. Load translated preview
        const outName = item.results.preview_output_filename || (item.results.docx_path ? item.results.docx_path.split(/[\\/]/).pop() : null);
        if (outName && paneTran) {
            paneTran.innerHTML = `<div class="pane-loading">Đang tải bản dịch song ngữ...</div>`;
            try {
                const resp = await fetch(`/api/preview/output/${encodeURIComponent(outName)}`);
                const res = await resp.json();
                if (res.success && res.html) {
                    paneTran.innerHTML = res.html;
                    this.setupSyncScroll();
                } else {
                    paneTran.innerHTML = `<div class="pane-loading">Bản dịch đã hoàn tất và sẵn sàng tải về!</div>`;
                }
            } catch (err) {
                paneTran.innerHTML = `<div class="pane-loading" style="color:#10b981;">✓ Bản dịch đã hoàn tất và sẵn sàng tải về!</div>`;
            }
        }
    },

    setupSyncScroll: function() {
        const p1 = document.getElementById("pane-original");
        const p2 = document.getElementById("pane-translated");
        if (!p1 || !p2) return;

        let isSyncingLeft = false;
        let isSyncingRight = false;

        p1.onscroll = () => {
            if (!this.syncScrollEnabled || isSyncingLeft) return;
            isSyncingRight = true;
            const percentage = p1.scrollTop / (p1.scrollHeight - p1.clientHeight || 1);
            p2.scrollTop = percentage * (p2.scrollHeight - p2.clientHeight);
            setTimeout(() => { isSyncingRight = false; }, 40);
        };

        p2.onscroll = () => {
            if (!this.syncScrollEnabled || isSyncingRight) return;
            isSyncingLeft = true;
            const percentage = p2.scrollTop / (p2.scrollHeight - p2.clientHeight || 1);
            p1.scrollTop = percentage * (p1.scrollHeight - p1.clientHeight);
            setTimeout(() => { isSyncingLeft = false; }, 40);
        };
    },

    toggleSyncScroll: function() {
        this.syncScrollEnabled = !this.syncScrollEnabled;
        const btn = document.getElementById("btn-sync-scroll");
        if (btn) {
            btn.className = this.syncScrollEnabled ? "hud-btn active" : "hud-btn";
            btn.querySelector("span").textContent = `Cuộn đồng bộ: ${this.syncScrollEnabled ? "BẬT" : "TẮT"}`;
        }
    },

    zoomChange: function(delta) {
        this.currentZoom = Math.max(60, Math.min(180, this.currentZoom + delta));
        const lbl = document.getElementById("zoom-level-label");
        if (lbl) lbl.textContent = this.currentZoom + "%";

        const p1 = document.getElementById("pane-original");
        const p2 = document.getElementById("pane-translated");
        const scale = this.currentZoom / 100;

        if (p1) p1.style.zoom = scale;
        if (p2) p2.style.zoom = scale;
    },

    toggleInlineEdit: function() {
        this.isInlineEditing = !this.isInlineEditing;
        const p2 = document.getElementById("pane-translated");
        const btn = document.getElementById("btn-inline-edit");
        if (p2) {
            p2.contentEditable = this.isInlineEditing ? "true" : "false";
            p2.style.outline = this.isInlineEditing ? "2px dashed #0284c7" : "none";
        }
        if (btn) {
            btn.className = this.isInlineEditing ? "hud-btn active" : "hud-btn";
            btn.querySelector("span").textContent = this.isInlineEditing ? "✓ Đang Sửa Nhanh" : "Sửa Nhanh";
        }
    },

    // ── 6. INSTANT PER-FILE DOWNLOAD HANDLERS ─────────────────────────────
    downloadFileWord: function(itemId) {
        const item = this.fileQueue.find(f => f.id === itemId);
        if (!item || !item.results) return;
        const url = item.results.download_docx_url || item.results.download_url;
        if (url) {
            window.location.href = url;
        } else {
            alert("Không tìm thấy đường dẫn tải tệp Word.");
        }
    },

    handleFilePdfDownload: async function(itemId) {
        const item = this.fileQueue.find(f => f.id === itemId);
        if (!item || !item.results) return;

        this.activeFileId = itemId;
        this.currentViewingResults = item.results;
        this.openPdfModal();

        // If PDF already rendered, trigger download
        if (item.results.download_pdf_url) {
            const stepReady = document.getElementById("pdf-step-ready");
            const stepRender = document.getElementById("pdf-step-rendering");
            if (stepRender) {
                stepRender.className = "pipeline-step step-complete";
                stepRender.innerHTML = `<span class="step-bullet">✓</span><span class="step-label">Đã kết xuất thành công tệp PDF vector!</span>`;
            }
            if (stepReady) {
                stepReady.className = "pipeline-step step-complete";
                stepReady.innerHTML = `<span class="step-bullet">✓</span><span class="step-label">Đang tải xuống tệp PDF...</span>`;
            }

            setTimeout(() => {
                this.closePdfModal();
                window.location.href = item.results.download_pdf_url;
            }, 1400);
            return;
        }

        // On-demand PDF conversion via backend
        const docxFile = item.results.preview_output_filename || (item.results.docx_path ? item.results.docx_path.split(/[\\/]/).pop() : null);
        if (!docxFile) {
            alert("Chưa có tệp Word dịch thuật để xuất PDF.");
            this.closePdfModal();
            return;
        }

        try {
            const formData = new FormData();
            formData.append("docx_filename", docxFile);

            const resp = await fetch("/api/export-pdf", {
                method: "POST",
                body: formData
            });
            const res = await resp.json();

            if (res.success && res.download_pdf_url) {
                const stepReady = document.getElementById("pdf-step-ready");
                const stepRender = document.getElementById("pdf-step-rendering");
                if (stepRender) {
                    stepRender.className = "pipeline-step step-complete";
                    stepRender.innerHTML = `<span class="step-bullet">✓</span><span class="step-label">Đã kết xuất thành công tệp PDF vector!</span>`;
                }
                if (stepReady) {
                    stepReady.className = "pipeline-step step-complete";
                    stepReady.innerHTML = `<span class="step-bullet">✓</span><span class="step-label">Bắt đầu tải tệp PDF...</span>`;
                }

                item.results.download_pdf_url = res.download_pdf_url;

                setTimeout(() => {
                    this.closePdfModal();
                    window.location.href = res.download_pdf_url;
                }, 1000);
            } else {
                throw new Error(res.detail || "Không thể kết xuất PDF.");
            }
        } catch (err) {
            console.error("Lỗi xuất PDF:", err);
            alert("Không thể kết xuất PDF tự động: " + err.message + "\nKhuyên dùng: Bạn vui lòng tải file Word (.docx) để sử dụng tức thì.");
            this.closePdfModal();
        }
    },

    downloadCurrentWord: function() {
        if (!this.currentViewingResults) return;
        const url = this.currentViewingResults.download_docx_url || this.currentViewingResults.download_url;
        if (url) window.location.href = url;
    },

    downloadCurrentPdf: function() {
        if (!this.activeFileId) return;
        this.handleFilePdfDownload(this.activeFileId);
    },

    downloadCurrentZip: function() {
        if (!this.currentViewingResults) return;
        const url = this.currentViewingResults.download_zip_url || this.currentViewingResults.download_docx_url;
        if (url) window.location.href = url;
    },

    openPdfModal: function() {
        const modal = document.getElementById("modal-pdf-progress");
        if (modal) modal.style.display = "flex";
    },

    closePdfModal: function() {
        const modal = document.getElementById("modal-pdf-progress");
        if (modal) modal.style.display = "none";
    },

    // ── 7. LICENSE ACTIVATION MODAL (PROFESSIONAL) ───────────────────────
    openLicenseModal: function() {
        const modal = document.getElementById("modal-license");
        if (modal) modal.style.display = "flex";
    },

    closeLicenseModal: function() {
        const modal = document.getElementById("modal-license");
        if (modal) modal.style.display = "none";
        const msg = document.getElementById("license-msg-box");
        if (msg) msg.style.display = "none";
    },

    handleActivateLicense: async function() {
        const regInput = document.getElementById("act-reg-name");
        const keyInput = document.getElementById("act-key");
        const msgBox = document.getElementById("license-msg-box");
        const btnSubmit = document.getElementById("btn-submit-activate");

        const reg = regInput?.value.trim() || "Doanh nghiệp LUV";
        const key = keyInput?.value.trim().toUpperCase() || "";

        if (!key) {
            alert("Vui lòng nhập Mã Bản Quyền chính thức!");
            return;
        }

        if (btnSubmit) btnSubmit.disabled = true;

        const formData = new FormData();
        formData.append("key", key);
        formData.append("registered_to", reg);

        try {
            const resp = await fetch("/api/license/activate", {
                method: "POST",
                body: formData
            });
            const res = await resp.json();

            if (resp.ok && res.success) {
                if (msgBox) {
                    msgBox.className = "license-alert-box success";
                    msgBox.style.display = "block";
                    msgBox.innerHTML = `✓ <strong>Kích hoạt thành công!</strong><br>Gói: ${res.data?.tier || 'Vĩnh Viễn'}<br>Đơn vị: ${res.data?.registered_to || reg}`;
                }

                // Update Header Badges
                const trialPill = document.getElementById("trial-pill");
                const licLabel = document.getElementById("license-header-label");
                if (trialPill) {
                    trialPill.className = "badge-pill trial-badge trial-active";
                    trialPill.innerHTML = `<span>Bản Quyền: ${res.data?.tier || 'Vĩnh Viễn'}</span>`;
                }
                if (licLabel) licLabel.textContent = "✓ Đã Kích Hoạt";

                setTimeout(() => {
                    this.closeLicenseModal();
                    alert("Kích hoạt bản quyền thành công. Toàn bộ tính năng dịch thuật đã được mở khóa.");
                }, 1400);
            } else {
                throw new Error(res.detail || "Mã bản quyền không hợp lệ.");
            }
        } catch (err) {
            if (msgBox) {
                msgBox.className = "license-alert-box error";
                msgBox.style.display = "block";
                msgBox.textContent = "✕ " + err.message;
            }
        } finally {
            if (btnSubmit) btnSubmit.disabled = false;
        }
    }
};

// Auto boot on DOM load
window.addEventListener("DOMContentLoaded", () => {
    window.LuvApp.init();
});
