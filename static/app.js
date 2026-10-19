/**
 * LUV-DOC-TRAN STUDIO PRO — Frontend Architecture Controller
 * ─────────────────────────────────────────────────────────────
 * Nền tảng Dịch thuật Song ngữ Đa ngữ & Đa định dạng Cao Cấp
 * Hỗ trợ Đa tệp Word, PDF & Excel • Tải về tức thì từng tệp (Word & Excel Only)
 * Review Split-View Studio • Quản lý Bản Quyền Ed25519 & Thu hồi Real-time
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

    // ── 1. DEVICE IDENTIFIER (ANTI-TAMPER FINGERPRINT) ───────────────────
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

    // ── 2. REAL-TIME LICENSE & ACTIVATION STATUS ─────────────────────────
    checkLicenseAndTrial: async function() {
        const trialPill = document.getElementById("trial-pill");
        const trialText = document.getElementById("trial-timer-text");
        const licLabel = document.getElementById("license-header-label");
        const dropWarning = document.getElementById("dropzone-key-warning");

        try {
            const savedKey = localStorage.getItem("luv_license_key") || "";
            // A. Check License status with local cache & auto-sync
            const url = `/api/license/status?device_id=${encodeURIComponent(this.deviceId)}&key=${encodeURIComponent(savedKey)}`;
            const licResp = await fetch(url);
            const licData = await licResp.json();

            if (licData.is_active) {
                this.isLicensed = true;
                if (trialPill) trialPill.className = "badge-pill key-badge key-licensed";
                const cleanTier = (licData.tier || 'Key Kích Hoạt (PRO)').replace(/doanh\s*nghiệp/gi, 'Key Kích Hoạt');
                if (trialText) trialText.textContent = cleanTier.includes('Key') ? cleanTier : `Key Kích Hoạt: ${cleanTier}`;
                if (licLabel) licLabel.textContent = "✓ Đã Kích Hoạt Key";
                if (dropWarning) dropWarning.style.display = "none";
                return;
            }

            // B. If Key was revoked on Cloudflare D1
            if (licData.revoked) {
                this.isLicensed = false;
                localStorage.removeItem("luv_license_key");
                if (trialPill) trialPill.className = "badge-pill key-badge key-revoked";
                if (trialText) trialText.textContent = "Key Đã Bị Thu Hồi";
                if (licLabel) licLabel.textContent = "✕ Key Bị Khóa";
                if (dropWarning) {
                    dropWarning.style.display = "inline-flex";
                    dropWarning.innerHTML = `<span>✕ Key kích hoạt đã bị thu hồi. Vui lòng nhập Key mới.</span>`;
                }
                return;
            }

            // C. Mandatory Key required - No Free Trial
            this.isLicensed = false;
            if (trialPill) trialPill.className = "badge-pill key-badge key-unactivated";
            if (trialText) trialText.textContent = "Chưa Nhập Key Kích Hoạt";
            if (licLabel) licLabel.textContent = "Nhập Key Kích Hoạt";
            if (dropWarning) dropWarning.style.display = "inline-flex";

        } catch (err) {
            console.error("Lỗi kiểm tra Key kích hoạt:", err);
        }
    },

    startTrialCountdown: function() {
        // No-op: Free trial mechanism eliminated, mandatory key enforced
    },

    updateTrialDisplay: function() {
        // No-op: Free trial mechanism eliminated, mandatory key enforced
    },

    // ── 3. DRAG & DROP MULTI-FILE QUEUE ──────────────────────────────────
    setupDragAndDrop: function() {
        const dropzone = document.getElementById("dropzone");
        if (!dropzone) return;

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, () => {
                dropzone.classList.add("dragover");
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, () => {
                dropzone.classList.remove("dragover");
            }, false);
        });

        dropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            this.handleFileSelect(files);
        }, false);
    },

    handleFileSelect: function(files) {
        if (!this.isLicensed) {
            this.showToast("Vui lòng nhập Key kích hoạt trước khi tải tệp lên hệ thống.", "warning");
            this.openLicenseModal();
            return;
        }
        if (!files || files.length === 0) return;

        const allowedExtensions = ['.docx', '.doc', '.pdf', '.xlsx', '.xls'];
        let addedCount = 0;

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const ext = '.' + file.name.split('.').pop().toLowerCase();

            if (!allowedExtensions.includes(ext)) {
                this.showToast(`Định dạng "${file.name}" không được hỗ trợ. Vui lòng chọn Word (.docx, .doc), PDF (.pdf) hoặc Excel (.xlsx, .xls).`, 'error');
                continue;
            }

            const queueItem = {
                id: 'file-' + Date.now() + '-' + Math.random().toString(36).substr(2, 6),
                file: file,
                name: file.name,
                size: file.size,
                ext: ext,
                status: 'waiting', // waiting | translating | ready | error
                statusText: 'Chờ xử lý',
                progress: 0,
                results: null,
                errorMsg: null,
                elapsed: 0
            };

            this.fileQueue.push(queueItem);
            addedCount++;
        }

        const fileInput = document.getElementById("file-input");
        if (fileInput) fileInput.value = "";

        if (addedCount > 0) {
            this.renderQueueTable();
            this.showToast(`Đã thêm ${addedCount} tệp vào hàng đợi xử lý.`, 'success');
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
        const btnDownloadAll = document.getElementById("btn-download-all");
        if (btnDownloadAll) {
            btnDownloadAll.style.display = completedCount > 0 ? "inline-flex" : "none";
        }
        if (summaryText) {
            if (completedCount === 0) {
                summaryText.textContent = `Đã chọn ${this.fileQueue.length} tài liệu. Nhấn "Bắt đầu dịch tất cả" để tiến hành.`;
            } else {
                summaryText.textContent = `Đã hoàn thành ${completedCount}/${this.fileQueue.length} tài liệu. Bạn có thể tải toàn bộ tệp (.ZIP) hoặc từng tệp bên trên!`;
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
            // EXCLUSIVE OUTPUT FORMATS: WORD (.docx) & EXCEL (.xlsx) ONLY. NO PDF OUTPUT.
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
                                🔍 So Sánh
                            </button>
                        </div>
                    `;
                } else {
                    // Both Word and PDF inputs exclusively output Word (.docx) for instant delivery
                    actionsHtml = `
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-row-word" onclick="window.LuvApp.downloadFileWord('${item.id}')" title="Tải Word xử lý tức thì">
                                ⚡ Tải Word (.docx)
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
        if (!this.isLicensed) {
            this.showToast("Vui lòng nhập Key kích hoạt để bắt đầu dịch thuật.", "warning");
            this.openLicenseModal();
            return;
        }

        if (this.fileQueue.length === 0) {
            this.showToast("Vui lòng tải lên ít nhất một tệp tài liệu trước khi bắt đầu.", "error");
            return;
        }

        const pendingItems = this.fileQueue.filter(f => f.status === 'waiting' || f.status === 'error');
        if (pendingItems.length === 0) {
            this.showToast("Tất cả tài liệu trong hàng đợi đã được xử lý xong!", "success");
            return;
        }

        this.isProcessingBatch = true;
        const btnBatch = document.getElementById("btn-start-batch");
        const btnText = document.getElementById("btn-batch-text");
        if (btnBatch) btnBatch.disabled = true;

        const sourceLang = document.getElementById("sel-source-lang")?.value || "auto";
        const targetLang = document.getElementById("sel-target-lang")?.value || "vi";
        const engine = document.getElementById("sel-engine")?.value || "neural";
        const fontFamily = document.getElementById("sel-font-family")?.value || "auto";

        for (let i = 0; i < pendingItems.length; i++) {
            const item = pendingItems[i];
            if (btnText) btnText.textContent = `ĐANG DỊCH TỆP ${i + 1}/${pendingItems.length}...`;

            item.status = 'translating';
            item.statusText = 'Đang phân tích bố cục...';
            this.renderQueueTable();

            try {
                const savedKey = localStorage.getItem("luv_license_key") || "";
                const formData = new FormData();
                formData.append("file", item.file);
                formData.append("source_lang", sourceLang);
                formData.append("target_lang", targetLang);
                formData.append("provider", engine === "fast" ? "google" : "neural");
                formData.append("engine_mode", "auto");
                formData.append("output_mode", "bilingual");
                formData.append("font_family", fontFamily);
                formData.append("device_id", this.deviceId);
                if (savedKey) {
                    formData.append("license_key", savedKey);
                }

                let simPercent = 20;
                item.statusText = `Đang phân tích bố cục... (${simPercent}%)`;
                this.renderQueueTable();

                const simInterval = setInterval(() => {
                    if (item.status === 'translating') {
                        if (simPercent < 45) {
                            simPercent += 10;
                            item.statusText = `Đang phân tích cấu trúc & bảng biểu... (${simPercent}%)`;
                        } else if (simPercent < 75) {
                            simPercent += 8;
                            item.statusText = `Đang dịch song ngữ liên dòng... (${simPercent}%)`;
                        } else if (simPercent < 92) {
                            simPercent += 3;
                            item.statusText = `Đang hoàn tất và tái tạo văn bản... (${simPercent}%)`;
                        }
                        this.renderQueueTable();
                    }
                }, 750);

                const resp = await fetch("/api/translate", {
                    method: "POST",
                    body: formData
                });

                clearInterval(simInterval);

                if (resp.status === 403) {
                    const errData = await resp.json();
                    item.status = 'error';
                    item.errorMsg = errData.message || "Yêu cầu Key kích hoạt!";
                    this.renderQueueTable();
                    this.showToast(errData.message || "Yêu cầu nhập Key kích hoạt.", "error");
                    if (errData.revoked) {
                        localStorage.removeItem("luv_license_key");
                    }
                    this.openLicenseModal();
                    break;
                }

                const res = await resp.json();
                if (!resp.ok || !res.success) {
                    throw new Error(res.detail || "Không thể dịch tài liệu.");
                }

                item.status = 'ready';
                item.statusText = 'Hoàn thành';
                item.results = res;
                item.elapsed = res.elapsed_seconds || 1;
                this.renderQueueTable();

                this.showToast(`Đã dịch xong "${item.name}". Bạn có thể tải ngay!`, 'success');

                // If first completed item, open Split-View
                if (i === 0) {
                    this.viewSplitStudio(item.id);
                }

            } catch (err) {
                console.error("Lỗi dịch tệp:", item.name, err);
                item.status = 'error';
                item.statusText = 'Lỗi xử lý';
                item.errorMsg = err.message || "Lỗi xử lý tệp tin.";
                this.renderQueueTable();
                this.showToast(`Lỗi khi xử lý tệp ${item.name}: ${err.message}`, 'error');
            }
        }

        this.isProcessingBatch = false;
        if (btnBatch) btnBatch.disabled = false;
        if (btnText) btnText.textContent = "BẮT ĐẦU DỊCH TẤT CẢ TỆP";
    },

    // ── 5. REVIEW SPLIT-VIEW STUDIO (SIDE-BY-SIDE) ───────────────────────
    viewSplitStudio: async function(itemId) {
        const item = this.fileQueue.find(f => f.id === itemId);
        if (!item || !item.results) return;

        this.activeFileId = itemId;
        this.currentViewingResults = item.results;

        const studio = document.getElementById("split-view-studio");
        const docName = document.getElementById("hud-doc-name");
        const paneOrig = document.getElementById("pane-original");
        const paneTran = document.getElementById("pane-translated");
        const dlTitle = document.getElementById("studio-dl-title");

        if (studio) {
            studio.style.display = "block";
            studio.scrollIntoView({ behavior: "smooth" });
        }

        if (docName) docName.textContent = item.name;

        // Update studio download button label
        if (dlTitle) {
            if (item.ext === '.xlsx' || item.ext === '.xls') {
                dlTitle.textContent = "TẢI FILE EXCEL (.XLSX)";
            } else {
                dlTitle.textContent = "TẢI FILE WORD (.DOCX)";
            }
        }

        if (paneOrig) paneOrig.innerHTML = `<div class="pane-loading"><span class="spinner-mini"></span> Đang nạp bản gốc...</div>`;
        if (paneTran) paneTran.innerHTML = `<div class="pane-loading"><span class="spinner-mini"></span> Đang nạp bản dịch song ngữ...</div>`;

        // Load preview of original
        try {
            const formData = new FormData();
            formData.append("file", item.file);
            const respOrig = await fetch("/api/preview/upload", {
                method: "POST",
                body: formData
            });
            const resOrig = await respOrig.json();
            if (resOrig.success && resOrig.html) {
                if (paneOrig) paneOrig.innerHTML = resOrig.html;
            } else {
                if (paneOrig) paneOrig.innerHTML = `<div class="pane-loading">Không thể hiển thị bản xem trước gốc.</div>`;
            }
        } catch (e) {
            if (paneOrig) paneOrig.innerHTML = `<div class="pane-loading">Lỗi xem trước tệp gốc.</div>`;
        }

        // Load preview of translated document
        try {
            const outFilename = item.results.preview_output_filename || (item.results.docx_path ? item.results.docx_path.split(/[\\/]/).pop() : null);
            if (outFilename) {
                const respTran = await fetch(`/api/preview/output/${encodeURIComponent(outFilename)}`);
                const resTran = await respTran.json();
                if (resTran.success && resTran.html) {
                    if (paneTran) paneTran.innerHTML = resTran.html;
                } else {
                    if (paneTran) paneTran.innerHTML = `<div class="pane-loading">Bản dịch đã hoàn tất. Bạn có thể tải ngay tệp bên dưới.</div>`;
                }
            }
        } catch (e) {
            if (paneTran) paneTran.innerHTML = `<div class="pane-loading">Bản dịch đã hoàn tất. Bạn có thể tải ngay tệp bên dưới.</div>`;
        }

        this.bindSynchronizedScroll();
    },

    bindSynchronizedScroll: function() {
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
            p2.style.outline = this.isInlineEditing ? "2px dashed #06b6d4" : "none";
        }
        if (btn) {
            btn.className = this.isInlineEditing ? "hud-btn active" : "hud-btn";
            btn.querySelector("span").textContent = this.isInlineEditing ? "✓ Đang Sửa Nhanh" : "Sửa Nhanh";
        }
    },

    // ── 6. INSTANT PER-FILE DOWNLOAD HANDLERS (WORD & EXCEL ONLY) ─────────
    downloadFileWord: function(itemId) {
        const item = this.fileQueue.find(f => f.id === itemId);
        if (!item || !item.results) return;
        const url = item.results.download_docx_url || item.results.download_url;
        if (url) {
            window.location.href = url;
        } else {
            this.showToast("Không tìm thấy đường dẫn tải tệp Word.", "error");
        }
    },

    downloadCurrentActiveFile: function() {
        if (!this.currentViewingResults) return;
        const url = this.currentViewingResults.download_docx_url || this.currentViewingResults.download_url;
        if (url) window.location.href = url;
    },

    downloadAllCompleted: async function() {
        const completedItems = this.fileQueue.filter(f => f.status === 'ready' && f.results);
        if (completedItems.length === 0) {
            this.showToast("Chưa có tệp nào được dịch hoàn tất để tải về.", "warning");
            return;
        }

        const filenames = [];
        completedItems.forEach(item => {
            const fname = item.results.preview_output_filename ||
                          item.results.filename ||
                          (item.results.docx_path ? item.results.docx_path.split(/[\\/]/).pop() : null);
            if (fname && !filenames.includes(fname)) {
                filenames.push(fname);
            }
        });

        if (filenames.length === 0) {
            this.showToast("Không tìm thấy tệp kết quả khả dụng để đóng gói.", "error");
            return;
        }

        const btn = document.getElementById("btn-download-all");
        const originalHtml = btn ? btn.innerHTML : "";
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = `<span class="spinner-mini"></span> Đang nén toàn bộ tệp (.ZIP)...`;
        }

        try {
            const resp = await fetch("/api/download-all-zip", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ filenames: filenames })
            });

            const res = await resp.json();
            if (!resp.ok || !res.success) {
                throw new Error(res.detail || "Không thể tạo gói nén ZIP.");
            }

            this.showToast(`Đã đóng gói thành công ${res.file_count} tệp! Bắt đầu tải về...`, "success");
            window.location.href = res.download_url;
        } catch (err) {
            console.error("Lỗi tải toàn bộ:", err);
            this.showToast("Lỗi đóng gói tải về: " + err.message, "error");
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = originalHtml;
            }
        }
    },

    // ── 7. LICENSE ACTIVATION MODAL (ED25519) ────────────────────────────
    openLicenseModal: function() {
        const modal = document.getElementById("modal-license");
        if (modal) modal.style.display = "flex";
        const keyInput = document.getElementById("act-key");
        const savedKey = localStorage.getItem("luv_license_key");
        if (keyInput && savedKey && !keyInput.value) {
            keyInput.value = savedKey;
        }
    },

    closeLicenseModal: function() {
        const modal = document.getElementById("modal-license");
        if (modal) modal.style.display = "none";
        const msgBox = document.getElementById("license-msg-box");
        if (msgBox) msgBox.style.display = "none";
    },

    handleActivateLicense: async function() {
        const keyInput = document.getElementById("act-key");
        const msgBox = document.getElementById("license-msg-box");
        const submitBtn = document.getElementById("btn-submit-activate");

        const key = (keyInput ? keyInput.value : "").trim();

        if (!key) {
            if (msgBox) {
                msgBox.className = "license-alert-box alert-error";
                msgBox.textContent = "Vui lòng nhập Mã Key Kích Hoạt chính thức!";
                msgBox.style.display = "block";
            }
            return;
        }

        if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.innerHTML = `<span class="spinner-mini"></span> Đang xác thực Key kích hoạt...`;
        }

        try {
            const formData = new FormData();
            formData.append("key", key);
            formData.append("registered_to", "Người Dùng Bản Quyền");
            formData.append("device_id", this.deviceId);

            const resp = await fetch("/api/license/activate", {
                method: "POST",
                body: formData
            });

            const res = await resp.json();

            if (resp.ok && res.success) {
                // Permanently persist Key in localStorage
                localStorage.setItem("luv_license_key", key);

                if (msgBox) {
                    msgBox.className = "license-alert-box alert-success";
                    msgBox.textContent = "✓ Key kích hoạt đã được xác thực thành công!";
                    msgBox.style.display = "block";
                }

                await this.checkLicenseAndTrial();
                this.showToast("Key kích hoạt đã được xác thực thành công!", "success");

                setTimeout(() => {
                    this.closeLicenseModal();
                }, 1400);

            } else {
                let errDetail = res.detail || res.message || "Mã Key kích hoạt không tồn tại hoặc đã đạt giới hạn thiết bị.";
                if (msgBox) {
                    msgBox.className = "license-alert-box alert-error";
                    msgBox.textContent = "✕ " + errDetail;
                    msgBox.style.display = "block";
                }
            }
        } catch (err) {
            if (msgBox) {
                msgBox.className = "license-alert-box alert-error";
                msgBox.textContent = "✕ Lỗi kết nối máy chủ xác thực: " + err.message;
                msgBox.style.display = "block";
            }
        } finally {
            if (submitBtn) {
                submitBtn.disabled = false;
                submitBtn.innerHTML = `<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> XÁC NHẬN KEY KÍCH HOẠT`;
            }
        }
    },

    // ── 8. MODERN TOAST NOTIFICATIONS ────────────────────────────────────
    showToast: function(message, type = 'info') {
        const container = document.getElementById("toast-container");
        if (!container) return;

        const toast = document.createElement("div");
        toast.className = `toast toast-${type}`;
        toast.textContent = message;

        container.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(10px)';
            toast.style.transition = 'all 0.3s ease';
            setTimeout(() => { toast.remove(); }, 300);
        }, 3500);
    }
};

// Bootstrap app on DOM ready
document.addEventListener("DOMContentLoaded", () => {
    window.LuvApp.init();
});
