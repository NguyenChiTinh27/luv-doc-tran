/**
 * LUV-DOC-TRAN — Frontend Controller
 * ──────────────────────────────────
 * Trình điều khiển giao diện Pastel Kawaii, Review Split-View Studio,
 * Đếm ngược dùng thử 48h và Cơ chế ưu tiên file Word + Loading PDF 2-3 phút.
 */

window.LuvApp = {
    selectedFile: null,
    currentResults: null,
    deviceId: null,
    trialRemainingSeconds: 0,
    trialTimerInterval: null,
    syncScrollEnabled: true,
    currentZoom: 100,
    isInlineEditing: false,

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
            // A. Check License first
            const licResp = await fetch("/api/license/status");
            const licData = await licResp.json();

            const trialIcon = document.getElementById("trial-icon");

            if (licData.is_active) {
                if (trialPill) trialPill.className = "trial-badge trial-active";
                if (trialIcon) trialIcon.textContent = "👑";
                if (trialText) trialText.textContent = `Bản Quyền: ${licData.tier || 'Vĩnh Viễn'}`;
                if (licLabel) licLabel.textContent = "✓ Đã Kích Hoạt";
                return;
            }

            // B. Check 48h Trial
            const trialResp = await fetch(`/api/trial/status?device_id=${encodeURIComponent(this.deviceId)}`);
            const trialData = await trialResp.json();

            this.trialRemainingSeconds = trialData.remaining_seconds || 0;

            if (trialData.is_trial_active && this.trialRemainingSeconds > 0) {
                if (trialPill) trialPill.className = "trial-badge trial-active";
                if (trialIcon) trialIcon.textContent = "🌸";
                this.updateTrialDisplay();
            } else {
                if (trialPill) trialPill.className = "trial-badge trial-expired";
                if (trialIcon) trialIcon.textContent = "🔒";
                if (trialText) trialText.textContent = "Hết Hạn Dùng Thử (Cần Key)";
            }
        } catch (err) {
            console.error("Lỗi kiểm tra bản quyền/dùng thử:", err);
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
            trialText.textContent = `Dùng thử: ${h}h ${m}m còn lại`;
        } else if (m > 0) {
            trialText.textContent = `Dùng thử: ${m}m ${s}s còn lại`;
        } else {
            trialText.textContent = `Hết hạn dùng thử`;
            const pill = document.getElementById("trial-pill");
            if (pill) pill.className = "trial-badge trial-expired";
        }
    },

    // ── 3. DRAG AND DROP & FILE SELECTION ────────────────────────────────
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
        const file = files[0];
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (!['.docx', '.doc', '.pdf', '.xlsx', '.xls'].includes(ext)) {
            alert('Vui lòng chọn định dạng được hỗ trợ: .docx, .doc, .pdf, .xlsx hoặc .xls nhen!');
            return;
        }

        this.selectedFile = file;

        // Display selected file card
        const dropzoneContent = document.querySelector(".dropzone-content");
        const fileCard = document.getElementById("selected-file-card");
        const fileName = document.getElementById("selected-file-name");
        const fileSize = document.getElementById("selected-file-size");
        const fileIcon = document.getElementById("file-type-icon");

        if (dropzoneContent) dropzoneContent.style.display = "none";
        if (fileCard) fileCard.style.display = "flex";
        if (fileName) fileName.textContent = file.name;
        if (fileSize) fileSize.textContent = this.formatFileSize(file.size);

        if (fileIcon) {
            if (['.docx', '.doc'].includes(ext)) fileIcon.textContent = "📝";
            else if (ext === '.pdf') fileIcon.textContent = "📄";
            else if (['.xlsx', '.xls'].includes(ext)) fileIcon.textContent = "📊";
        }

        // Change hero mascot expression
        const mascot = document.getElementById("hero-mascot-avatar");
        if (mascot) mascot.textContent = "😻";
    },

    clearSelectedFile: function() {
        this.selectedFile = null;
        const input = document.getElementById("file-input");
        if (input) input.value = "";

        const dropzoneContent = document.querySelector(".dropzone-content");
        const fileCard = document.getElementById("selected-file-card");
        if (dropzoneContent) dropzoneContent.style.display = "block";
        if (fileCard) fileCard.style.display = "none";

        const mascot = document.getElementById("hero-mascot-avatar");
        if (mascot) mascot.textContent = "🐱";
    },

    formatFileSize: function(bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / 1048576).toFixed(2) + " MB";
    },

    toggleApiKeySection: function() {
        const panel = document.getElementById("api-key-panel");
        const arrow = document.getElementById("accordion-arrow");
        if (panel) {
            const isHidden = panel.style.display === "none";
            panel.style.display = isHidden ? "block" : "none";
            if (arrow) arrow.textContent = isHidden ? "▲" : "▼";
        }
    },

    // ── 4. TRANSLATION PIPELINE ──────────────────────────────────────────
    startTranslation: async function() {
        if (!this.selectedFile) {
            alert("Bạn ơi, hãy chọn một tệp tài liệu trước khi bấm dịch nhen! 🌸");
            return;
        }

        const btnTranslate = document.getElementById("btn-start-translate");
        const btnText = document.getElementById("btn-translate-text");
        const progressBox = document.getElementById("progress-box");
        const progressFill = document.getElementById("progress-bar-fill");
        const progressStatus = document.getElementById("progress-status-text");
        const progressPercent = document.getElementById("progress-percent-text");

        const sourceLang = document.getElementById("sel-source-lang")?.value || "auto";
        const engine = document.getElementById("sel-engine")?.value || "deepseek";
        const apiKey = document.getElementById("input-api-key")?.value || "";

        const formData = new FormData();
        formData.append("file", this.selectedFile);
        formData.append("source_lang", sourceLang);
        formData.append("target_lang", "vi");
        formData.append("provider", engine === "deepseek" ? "deepseek" : "google");
        formData.append("engine_mode", "auto");
        formData.append("output_mode", "bilingual");
        formData.append("api_key", apiKey);
        formData.append("device_id", this.deviceId);

        // UI state: translating
        if (btnTranslate) btnTranslate.disabled = true;
        if (btnText) btnText.textContent = "ĐANG DỊCH SONG NGỮ...";
        if (progressBox) progressBox.style.display = "block";

        let percent = 15;
        this.updateProgress(percent, "Đang tải tệp lên và phân tích bố cục tài liệu...");

        const progressTimer = setInterval(() => {
            if (percent < 85) {
                percent += Math.floor(Math.random() * 8) + 4;
                if (percent > 85) percent = 85;
                let status = "Đang dịch liên dòng chuẩn nhận diện (#002060 nghiêng)...";
                if (percent > 50) status = "Đang bảo toàn bảng biểu, cấu trúc ô và mã CAS...";
                if (percent > 70) status = "Đang chuẩn bị bản xem trước Split-View...";
                this.updateProgress(percent, status);
            }
        }, 800);

        try {
            // First load original preview in background
            this.loadOriginalPreview(this.selectedFile);

            const resp = await fetch("/api/translate", {
                method: "POST",
                body: formData
            });

            clearInterval(progressTimer);

            if (resp.status === 403) {
                const errData = await resp.json();
                this.updateProgress(100, "Đã hết hạn dùng thử!");
                alert(errData.message || "Thời hạn dùng thử 48h đã kết thúc! Vui lòng kích hoạt bản quyền để tiếp tục.");
                this.openLicenseModal();
                return;
            }

            const res = await resp.json();

            if (!resp.ok || !res.success) {
                throw new Error(res.detail || "Không thể dịch tài liệu.");
            }

            this.updateProgress(100, "Dịch song ngữ hoàn tất thành công! ✨");
            this.currentResults = res;

            // Show Split-View Studio
            setTimeout(() => {
                if (progressBox) progressBox.style.display = "none";
                this.displaySplitViewStudio(res);
            }, 600);

        } catch (err) {
            clearInterval(progressTimer);
            console.error("Lỗi khi dịch:", err);
            alert("Có lỗi xảy ra trong quá trình dịch: " + err.message);
            this.updateProgress(0, "Lỗi: " + err.message);
        } finally {
            if (btnTranslate) btnTranslate.disabled = false;
            if (btnText) btnText.textContent = "BẮT ĐẦU DỊCH SONG NGỮ NGAY";
        }
    },

    updateProgress: function(percent, status) {
        const fill = document.getElementById("progress-bar-fill");
        const stat = document.getElementById("progress-status-text");
        const pct = document.getElementById("progress-percent-text");
        if (fill) fill.style.width = percent + "%";
        if (stat) stat.textContent = status;
        if (pct) pct.textContent = percent + "%";
    },

    // ── 5. REVIEW SPLIT-VIEW STUDIO & PREVIEWS ───────────────────────────
    loadOriginalPreview: async function(file) {
        const pane = document.getElementById("pane-original");
        const meta = document.getElementById("orig-page-indicator");
        if (!pane) return;

        try {
            pane.innerHTML = `<div class="pane-loading">Đang phân giải văn bản gốc...</div>`;
            const fd = new FormData();
            fd.append("file", file);

            const resp = await fetch("/api/preview/upload", {
                method: "POST",
                body: fd
            });
            const res = await resp.json();

            if (res.success && res.html) {
                pane.innerHTML = res.html;
                if (meta) meta.textContent = `${res.type?.toUpperCase() || 'VĂN BẢN'} (${file.name})`;
                this.setupSyncScroll();
            } else {
                pane.innerHTML = `<div class="pane-loading">Không thể hiển thị xem trước tệp gốc.</div>`;
            }
        } catch (err) {
            pane.innerHTML = `<div class="pane-loading" style="color:#ef4444;">Lỗi xem trước gốc: ${err.message}</div>`;
        }
    },

    displaySplitViewStudio: async function(results) {
        const studio = document.getElementById("split-view-studio");
        const hudName = document.getElementById("hud-doc-name");
        const paneTran = document.getElementById("pane-translated");

        if (hudName) hudName.textContent = this.selectedFile?.name || "document";
        if (studio) {
            studio.style.display = "block";
            studio.scrollIntoView({ behavior: "smooth", block: "start" });
        }

        // Change hero mascot to celebratory
        const mascot = document.getElementById("hero-mascot-avatar");
        if (mascot) mascot.textContent = "🎉";

        // Load translated preview
        const outName = results.preview_output_filename || (results.docx_path ? results.docx_path.split(/[\\/]/).pop() : null);
        if (outName && paneTran) {
            paneTran.innerHTML = `<div class="pane-loading">Đang tải bản dịch song ngữ...</div>`;
            try {
                const resp = await fetch(`/api/preview/output/${encodeURIComponent(outName)}`);
                const res = await resp.json();
                if (res.success && res.html) {
                    paneTran.innerHTML = res.html;
                    this.setupSyncScroll();
                } else {
                    paneTran.innerHTML = `<div class="pane-loading">Bản dịch đã hoàn tất và sẵn sàng tải về qua bộ nút bên dưới!</div>`;
                }
            } catch (err) {
                paneTran.innerHTML = `<div class="pane-loading" style="color:#059669;">✓ Bản dịch đã sẵn sàng tải về qua nút bên dưới!</div>`;
            }
        }
    },

    // ── 6. SYNCHRONIZED SCROLLING & ZOOM CONTROLS ───────────────────────
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
            btn.textContent = `🔄 Cuộn đồng bộ: ${this.syncScrollEnabled ? "BẬT" : "TẮT"}`;
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
            btn.textContent = this.isInlineEditing ? "✓ Đang Sửa Nhanh" : "✏️ Sửa Nhanh";
        }
    },

    // ── 7. DOWNLOAD BUTTONS (WORD PRIORITY & CUTE PDF MODAL) ─────────────
    downloadWord: function() {
        if (!this.currentResults) return;
        const url = this.currentResults.download_docx_url || this.currentResults.download_url;
        if (url) {
            window.location.href = url;
        } else {
            alert("Không tìm thấy đường dẫn tải tệp Word.");
        }
    },

    handlePdfDownload: async function() {
        if (!this.currentResults) return;

        // Always open Cute Progress Modal showing the 2-3 minute reminder & animation
        this.openPdfModal();

        // If PDF is already ready in results, show the cute progress and download after brief notice
        if (this.currentResults.download_pdf_url) {
            const stepReady = document.getElementById("pdf-step-ready");
            const stepRender = document.getElementById("pdf-step-rendering");
            if (stepRender) {
                stepRender.className = "pdf-step step-done";
                stepRender.innerHTML = `<span class="step-icon">✓</span><span class="step-text">Đã kết xuất thành công tệp PDF vector!</span>`;
            }
            if (stepReady) {
                stepReady.className = "pdf-step step-done";
                stepReady.innerHTML = `<span class="step-icon">✓</span><span class="step-text">Tự động tải về tệp PDF...</span>`;
            }

            setTimeout(() => {
                this.closePdfModal();
                window.location.href = this.currentResults.download_pdf_url;
            }, 1800);
            return;
        }

        const docxFile = this.currentResults.preview_output_filename || (this.currentResults.docx_path ? this.currentResults.docx_path.split(/[\\/]/).pop() : null);
        if (!docxFile) {
            alert("Chưa có bản Word dịch thuật để xuất PDF.");
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
                // Mark steps complete
                const stepReady = document.getElementById("pdf-step-ready");
                const stepRender = document.getElementById("pdf-step-rendering");
                if (stepRender) {
                    stepRender.className = "pdf-step step-done";
                    stepRender.innerHTML = `<span class="step-icon">✓</span><span class="step-text">Đã kết xuất thành công tệp PDF vector!</span>`;
                }
                if (stepReady) {
                    stepReady.className = "pdf-step step-done";
                    stepReady.innerHTML = `<span class="step-icon">✓</span><span class="step-text">Tự động bắt đầu tải tệp PDF...</span>`;
                }

                this.currentResults.download_pdf_url = res.download_pdf_url;

                setTimeout(() => {
                    this.closePdfModal();
                    window.location.href = res.download_pdf_url;
                }, 1200);
            } else {
                throw new Error(res.detail || "Lỗi không xác định khi xuất PDF.");
            }
        } catch (err) {
            console.error("Lỗi xuất PDF:", err);
            alert("Không thể kết xuất PDF tự động: " + err.message + "\nBạn vui lòng tải file Word (.docx) trước nhen!");
            this.closePdfModal();
        }
    },

    downloadZip: function() {
        if (!this.currentResults) return;
        const url = this.currentResults.download_zip_url;
        if (url) {
            window.location.href = url;
        } else {
            // Fallback: download Word
            this.downloadWord();
        }
    },

    openPdfModal: function() {
        const modal = document.getElementById("modal-pdf-progress");
        if (modal) modal.style.display = "flex";
    },

    closePdfModal: function() {
        const modal = document.getElementById("modal-pdf-progress");
        if (modal) modal.style.display = "none";
    },

    // ── 8. LICENSE ACTIVATION MODAL ──────────────────────────────────────
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
            alert("Vui lòng nhập Mã Bản Quyền!");
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
                    msgBox.className = "license-msg-box success";
                    msgBox.style.display = "block";
                    msgBox.innerHTML = `🎉 <strong>Kích hoạt thành công!</strong><br>Gói: ${res.data?.tier || 'Vĩnh Viễn'}<br>Đăng ký cho: ${res.data?.registered_to || reg}`;
                }

                // Update Header
                const trialPill = document.getElementById("trial-pill");
                const licLabel = document.getElementById("license-header-label");
                if (trialPill) {
                    trialPill.className = "trial-badge trial-active";
                    trialPill.innerHTML = `<span>👑</span><span>Bản Quyền: ${res.data?.tier || 'Vĩnh Viễn'}</span>`;
                }
                if (licLabel) licLabel.textContent = "✓ Đã Kích Hoạt";

                setTimeout(() => {
                    this.closeLicenseModal();
                    alert("Chúc mừng bạn! Bản quyền LUV-DOC-TRAN đã được kích hoạt thành công!");
                }, 1600);
            } else {
                throw new Error(res.detail || "Mã bản quyền không hợp lệ.");
            }
        } catch (err) {
            if (msgBox) {
                msgBox.className = "license-msg-box error";
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
