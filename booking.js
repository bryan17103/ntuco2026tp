document.addEventListener("DOMContentLoaded", () => {
    const csvFileInput = document.getElementById("csv-file");
    const batchCountInput = document.getElementById("batch-count");
    const fileNameDisplay = document.getElementById("file-name-display");
    const bookingTbody = document.getElementById("booking-tbody");
    const loadingOverlay = document.getElementById("orders-loading-overlay");
    
    // 篩選控制項
    const filterNameInput = document.getElementById("filter-name");
    const filterBatchSelect = document.getElementById("filter-batch");
    const filterStatusSelect = document.getElementById("filter-status");
    const concertTabs = document.querySelectorAll("[data-filter-concert]");
    const sortOrderNoBtn = document.getElementById("sort-order-no-btn");
    const sortIcon = document.getElementById("sort-icon");  
    
    const saveChangesBtn = document.getElementById("save-all-changes-btn");
    const unsavedBadge = document.getElementById("unsaved-badge");
    const filterOrderNoInput = document.getElementById("filter-order-no"); 

    // 全域變數
    let allLoadedEntries = [];
    let currentConcertFilter = "all";
    let currentNameFilter = "";
    let currentBatchFilter = "all";
    let currentStatusFilter = "all";
    let currentOrderNoFilter = ""; 
    let currentSortMode = "default";

    let unsavedChanges = {};

    // 網頁初始化載入
    loadHistoryRecords();

    function loadHistoryRecords() {
        loadingOverlay.classList.remove("hidden");
        document.getElementById("orders-loading-text").textContent = "正在載入調票紀錄...";
        unsavedChanges = {};
        updateUnsavedBadge();

        fetch("/api/admin/booking/records")
        .then(response => {
            // 💡 新增：如果後端判定沒權限（401）
            if (response.status === 401) {
                alert("你不是票務！滾吧 😠");
                window.location.href = "/";
                throw new Error("Unauthorized");
            }
            return response.json();
        })
        .then(res => {
            loadingOverlay.classList.add("hidden");
            document.getElementById("orders-loading-text").textContent = "正在讀取 CSV 並跨工作表比對、回填中，請勿關閉視窗...";
            
            if (res.success) {
                allLoadedEntries = res.data;
                populateBatchOptions(allLoadedEntries);
                applyFiltersAndRender();
            } else {
                // 💡 如果權限不足或其他後端錯誤
                alert(res.message || "處理失敗");
                window.location.href = "/";
            }
        })
        .catch(err => {
            loadingOverlay.classList.add("hidden");
            console.error(err);
        });
    }

    function populateBatchOptions(entries) {
        const previousSelection = filterBatchSelect.value;
        const batchSet = new Set();
        entries.forEach(row => {
            if (row[7]) batchSet.add(row[7].trim ? row[7].trim() : row[7]);
        });

        const sortedBatches = Array.from(batchSet).sort();
        filterBatchSelect.innerHTML = '<option value="all">全部批次</option>';
        sortedBatches.forEach(batchName => {
            const option = document.createElement("option");
            option.value = batchName;
            option.textContent = batchName;
            filterBatchSelect.appendChild(option);
        });

        if (batchSet.has(previousSelection)) {
            filterBatchSelect.value = previousSelection;
            currentBatchFilter = previousSelection;
        } else {
            filterBatchSelect.value = "all";
            currentBatchFilter = "all";
        }
    }

    csvFileInput.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) return;

        if (Object.keys(unsavedChanges).length > 0) {
            alert("您目前有尚未儲存的情況狀態變更，請先點擊「儲存所有變更」按鈕，再上傳新檔案！");
            csvFileInput.value = "";
            return;
        }

        fileNameDisplay.textContent = `已選取檔案：${file.name}`;
        const batchCount = batchCountInput.value.trim();
        if (!batchCount) {
            alert("請先輸入「這是第幾次調票」的次數，再進行上傳匯入！");
            csvFileInput.value = "";
            fileNameDisplay.textContent = "";
            batchCountInput.focus();
            return;
        }

        if (!confirm(`確認要上傳此調票檔案嗎？\n\n檔案名稱：${file.name}\n調票批次：第 ${batchCount} 次調票`)) {
            csvFileInput.value = "";
            fileNameDisplay.textContent = "";
            return;
        }

        loadingOverlay.classList.remove("hidden");
        const formData = new FormData();
        formData.append("file", file);
        formData.append("batch_count", batchCount);

        fetch("/api/admin/booking/import", {
            method: "POST",
            body: formData
        })
        .then(response => response.json())
        .then(res => {
            loadingOverlay.classList.add("hidden");
            csvFileInput.value = "";
            if (res.success) {
                alert(res.message);
                loadHistoryRecords();
            } else {
                alert(`處理失敗：${res.message}`);
            }
        })
        .catch(err => {
            loadingOverlay.classList.add("hidden");
            csvFileInput.value = "";
            alert(`連線伺服器發生錯誤：${err}`);
        });
    });

    concertTabs.forEach(tab => {
        tab.addEventListener("click", () => {
            concertTabs.forEach(t => t.classList.remove("active"));
            tab.classList.add("active");
            currentConcertFilter = tab.getAttribute("data-filter-concert");
            applyFiltersAndRender();
        });
    });

    filterNameInput.addEventListener("input", (e) => {
        currentNameFilter = e.target.value.trim().toLowerCase();
        applyFiltersAndRender();
    });

    filterOrderNoInput.addEventListener("input", (e) => {
        currentOrderNoFilter = e.target.value.trim().toLowerCase();
        applyFiltersAndRender();
    });

    sortOrderNoBtn.addEventListener("click", () => {
        if (currentSortMode === "default") {
            currentSortMode = "asc";
            sortIcon.textContent = "🔼 A-Z";
            sortOrderNoBtn.style.color = "var(--orders-main)";
        } else if (currentSortMode === "asc") {
            currentSortMode = "desc";
            sortIcon.textContent = "🔽 Z-A";
            sortOrderNoBtn.style.color = "var(--orders-main)";
        } else {
            currentSortMode = "default";
            sortIcon.textContent = "↕️";
            sortOrderNoBtn.style.color = "";
        }
        applyFiltersAndRender();
    });

    filterBatchSelect.addEventListener("change", (e) => {
        currentBatchFilter = e.target.value;
        applyFiltersAndRender();
    });

    filterStatusSelect.addEventListener("change", (e) => {
        currentStatusFilter = e.target.value;
        applyFiltersAndRender();
    });

    function getNaturalSortKey(orderNoStr) {
        const str = orderNoStr || "";
        const parts = str.split("-");
        if (parts.length === 2 && !isNaN(parts[1])) {
            return {
                prefix: parts[0],
                suffixNum: parseInt(parts[1], 10)
            };
        }
        return { prefix: str, suffixNum: 0 };
    }

    function applyFiltersAndRender() {
        if (allLoadedEntries.length === 0) {
            renderBookingTable([]);
            return;
        }

        let filtered = allLoadedEntries.filter(row => {
            const orderNo = (row[0] || "").toLowerCase();
            const seatRaw = row[1];
            const location = row[3] || "";
            const buyerName = (row[5] || "").toLowerCase();
            const batchName = row[7] || "";
            
            const changeKey = `${row[0]}_${seatRaw}`;
            const currentStatus = unsavedChanges[changeKey]?.status !== undefined ? unsavedChanges[changeKey].status : row[6];

            let matchConcert = true;
            if (currentConcertFilter === "tp") matchConcert = location.includes("中山堂");
            if (currentConcertFilter === "kh") matchConcert = location.includes("衛武營");

            let matchName = true;
            if (currentNameFilter) matchName = buyerName.includes(currentNameFilter);

            let matchOrderNo = true;
            if (currentOrderNoFilter) matchOrderNo = orderNo.includes(currentOrderNoFilter);

            let matchBatch = true;
            if (currentBatchFilter !== "all") matchBatch = (batchName === currentBatchFilter);

            let matchStatus = true;
            if (currentStatusFilter !== "all") matchStatus = (currentStatus === currentStatusFilter);

            return matchConcert && matchName && matchOrderNo && matchBatch && matchStatus;
        });

        if (currentSortMode === "asc") {
            filtered.sort((a, b) => {
                const keyA = getNaturalSortKey(a[0]);
                const keyB = getNaturalSortKey(b[0]);
                const prefixCompare = keyA.prefix.localeCompare(keyB.prefix);
                if (prefixCompare !== 0) return prefixCompare;
                return keyA.suffixNum - keyB.suffixNum;
            });
        } else if (currentSortMode === "desc") {
            filtered.sort((a, b) => {
                const keyA = getNaturalSortKey(a[0]);
                const keyB = getNaturalSortKey(b[0]);
                const prefixCompare = keyB.prefix.localeCompare(keyA.prefix);
                if (prefixCompare !== 0) return prefixCompare;
                return keyB.suffixNum - keyA.suffixNum;
            });
        }

        renderBookingTable(filtered);
    }

    // 表格渲染與「情況/姓名編輯」連動處理
    function renderBookingTable(entries) {
        const countBadge = document.getElementById("filtered-count-badge");
        if (countBadge) {
            countBadge.textContent = entries ? entries.length : 0;
        }

        bookingTbody.innerHTML = "";
        if (entries.length === 0) {
            bookingTbody.innerHTML = `<tr class="empty-row"><td colspan="8"><div class="empty-state">沒有符合目前複合篩選條件的調票紀錄。</div></td></tr>`;
            return;
        }

        entries.forEach(row => {
            const tr = document.createElement("tr");
            const orderNo = row[0];
            const seatRaw = row[1];
            const changeKey = `${orderNo}_${seatRaw}`;

            const displayedStatus = unsavedChanges[changeKey]?.status !== undefined ? unsavedChanges[changeKey].status : row[6];
            const displayedName = unsavedChanges[changeKey]?.buyer_name !== undefined ? unsavedChanges[changeKey].buyer_name : (row[5] || '（未對應）');
            
            const statusClassMap = { "已調": "closed", "已傳": "open", "已取": "done", "已分好": "open" };
            const currentClass = statusClassMap[displayedStatus] || "closed";

            tr.innerHTML = `
                <td class="order-strong">${orderNo}</td>
                <td>${seatRaw}</td>
                <td style="font-size: 13px;">${row[2]}</td>
                <td style="text-align: left; font-size: 13px;">${row[3]}</td>
                <td class="order-strong" style="color: var(--orders-main);">$${row[4]}</td>
                <td>
                    <div class="buyer-name-input" contenteditable="true" style="font-weight: 900; color: #333; border-bottom: 1.5px dashed #ccc; padding: 2px 4px; display: inline-block; min-width: 60px; outline: none; cursor: text;" title="點擊可直接修改姓名">
                        ${displayedName}
                    </div>
                </td>
                <td>
                    <select class="status-select pickup-pill ${currentClass}" style="height:34px; min-width:95px; padding:0 6px; font-size:13px; border:1px solid #ddd; cursor:pointer; text-align-last:center;">
                        <option value="已調" ${displayedStatus === '已調' ? 'selected' : ''}>已調</option>
                        <option value="已傳" ${displayedStatus === '已傳' ? 'selected' : ''}>已傳</option>
                        <option value="已取" ${displayedStatus === '已取' ? 'selected' : ''}>已取</option>
                        <option value="已分好" ${displayedStatus === '已分好' ? 'selected' : ''}>已分好</option>
                    </select>
                </td>
                <td>${row[7]}</td>
            `;

            const selectEl = tr.querySelector(".status-select");
            const nameEl = tr.querySelector(".buyer-name-input");
            
            let isAltPressed = false;
            selectEl.addEventListener("click", (e) => { isAltPressed = e.altKey; });

            selectEl.addEventListener("change", (e) => {
                const updatedStatus = e.target.value;
                const currentName = nameEl.textContent.trim();

                if (isAltPressed) {
                    selectEl.className = `status-select pickup-pill ${statusClassMap[updatedStatus] || 'closed'}`;
                    saveToUnsavedChanges(changeKey, orderNo, seatRaw, updatedStatus, currentName, row);
                } else {
                    const mainOrderId = orderNo.split("-")[0];
                    const sameOrderSelects = document.querySelectorAll("#booking-tbody tr");

                    sameOrderSelects.forEach(rowEl => {
                        const cellOrderNo = rowEl.cells[0].textContent.trim();
                        const cellSeatRaw = rowEl.cells[1].textContent.trim();
                        const cellMainId = cellOrderNo.split("-")[0];

                        if (cellMainId === mainOrderId) {
                            const targetSelect = rowEl.querySelector(".status-select");
                            const targetNameEl = rowEl.querySelector(".buyer-name-input");
                            
                            if (targetSelect && targetNameEl) {
                                targetSelect.value = updatedStatus;
                                targetSelect.className = `status-select pickup-pill ${statusClassMap[updatedStatus] || 'closed'}`;
                                
                                const loopKey = `${cellOrderNo}_${cellSeatRaw}`;
                                const loopOriginalRow = allLoadedEntries.find(r => r[0] === cellOrderNo && r[1] === cellSeatRaw);
                                saveToUnsavedChanges(loopKey, cellOrderNo, cellSeatRaw, updatedStatus, targetNameEl.textContent.trim(), loopOriginalRow);
                            }
                        }
                    });
                }
                updateUnsavedBadge();
                isAltPressed = false;
                
                if (currentStatusFilter !== "all") {
                    setTimeout(applyFiltersAndRender, 300);
                }
            });

            nameEl.addEventListener("blur", () => {
                const updatedName = nameEl.textContent.trim();
                const currentStatus = selectEl.value;
                saveToUnsavedChanges(changeKey, orderNo, seatRaw, currentStatus, updatedName, row);
                updateUnsavedBadge();
            });

            bookingTbody.appendChild(tr);
        });
    }

    function saveToUnsavedChanges(key, orderNo, seatRaw, status, buyer_name, originalRow) {
        const origStatus = originalRow ? originalRow[6] : "";
        const origName = originalRow ? (originalRow[5] || '（未對應）') : "";

        if (status === origStatus && buyer_name === origName) {
            delete unsavedChanges[key];
        } else {
            unsavedChanges[key] = {
                order_no: orderNo,
                seat_raw: seatRaw,
                status: status,
                buyer_name: buyer_name
            };
        }
    }

    function updateUnsavedBadge() {
        const count = Object.keys(unsavedChanges).length;
        if (count > 0) {
            unsavedBadge.style.display = "inline-flex";
            unsavedBadge.textContent = count;
            saveChangesBtn.style.boxShadow = "0 0 15px #f0a59a";
            saveChangesBtn.style.borderColor = "#f0a59a";
        } else {
            unsavedBadge.style.display = "none";
            saveChangesBtn.style.boxShadow = "0 4px 12px rgba(127,159,152,0.3)";
            saveChangesBtn.style.borderColor = "transparent";
        }
    }

    saveChangesBtn.addEventListener("click", () => {
        const changeItems = Object.values(unsavedChanges);
        if (changeItems.length === 0) {
            alert("目前沒有任何欄位被修改，不需要儲存唷！");
            return;
        }

        loadingOverlay.classList.remove("hidden");
        document.getElementById("orders-loading-text").textContent = "儲存中⋯⋯";

        fetch("/api/admin/booking/batch-update-status", {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ updates: changeItems })
        })
        .then(response => {
            if (response.status === 401) {
                alert("儲存失敗：您不是票務，請重新登入！");
                window.location.href = "/";
                throw new Error("Unauthorized");
            }
            return response.json();
        })
        .then(res => {
            loadingOverlay.classList.add("hidden");
            if (res.success) {
                alert(res.message);
                loadHistoryRecords();
            } else {
                alert(`儲存失敗：${res.message}`);
            }
        })
        .catch(err => {
            loadingOverlay.classList.add("hidden");
            alert(`儲存通訊錯誤：${err}`);
        });
    });

    window.addEventListener("beforeunload", (e) => {
        if (Object.keys(unsavedChanges).length > 0) {
            e.preventDefault();
            e.returnValue = "您尚有修改後的調票狀態未儲存，確定要離開嗎？";
        }
    });
});