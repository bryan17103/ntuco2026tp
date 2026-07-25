let mySeatMapScale = 1;
let currentSearchName = "";
let currentOrders = [];
let currentConcertMode = "all"; // all | tp | kh
let cachedSeatMapData = null;
let currentRewardStats = {};

const SECOND_FLOOR_START_ROW = 33;

// 💡 獎勵說明對照表
const REWARD_GUIDELINES = {
    "慶功宴": "關於慶功宴之獎勵待慶功宴場地確定後會正式公告。",
    "公演USB": "領取此獎勵者請待樂隊部長於大薰風發布 USB 表單後填寫表單，未在期限內填寫表單者將無法取得此獎勵 😭",
    "文宣品": "領取此獎勵者請在加練期間主動找到票務領取，高雄場總彩結束後將無法取得此獎勵！（或開學後的迎新茶會）"
};

function cleanEmoji(text) {
    return String(text ?? "")
        .replace(/[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]/gu, '')
        .trim();
}

function getRewardGuidelineText(rewardName) {
    for (const key in REWARD_GUIDELINES) {
        if (rewardName.includes(key)) return REWARD_GUIDELINES[key];
    }
    return "詳細領取方式請洽票務或關注團內公告。";
}

function getRewardIcon(rewardName) {
    if (rewardName.includes("慶功宴")) return "🍷";
    if (rewardName.includes("公演USB")) return "💾";
    if (rewardName.includes("文宣品")) return "🎁";
    return "✨";
}

function escapeHtml(text) {
    return String(text ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatNumber(value) {
    const num = Number(value || 0);
    return Number.isInteger(num) ? String(num) : num.toFixed(1);
}

function formatMoney(value) {
    return `$${Number(value || 0).toLocaleString()}`;
}

function getModeLabel(mode = currentConcertMode) {
    if (mode === "tp") return "台北場";
    if (mode === "kh") return "高雄場";
    return "全部";
}

function showOrdersLoading(text = "載入中...") {
    const overlay = document.getElementById("orders-loading-overlay");
    const textEl = document.getElementById("orders-loading-text");

    if (textEl) textEl.textContent = text;
    overlay?.classList.remove("hidden");
}

function hideOrdersLoading() {
    const overlay = document.getElementById("orders-loading-overlay");
    overlay?.classList.add("hidden");
}

function pickupClass(order) {
    if (order.picked_up) return "done";
    if (order.pickup_open) return "open";
    return "closed";
}

function pickupStatusHtml(order) {
    const paymentText = order.payment_done ? "已付款" : "未付款";

    let pickupText = "尚未開放取票";
    if (order.picked_up) pickupText = "已取票";
    else if (order.pickup_open) pickupText = "已開放取票";

    return `
        <button class="pickup-pill ${pickupClass(order)}" type="button">
            ${paymentText}｜${pickupText}
        </button>
    `;
}

function orderNoteInputId(order) {
    return `note-${order.order_id}-${order.floor}-${order.row_label}`;
}

function makeSeatKey(floor, rowLabel, seatNumber) {
    return `${String(floor || "").trim()}|${String(rowLabel || "").trim()}|${Number(seatNumber)}`;
}

function floorFromSeat(seat) {
    return seat.floor || (Number(seat.excel_row) >= SECOND_FLOOR_START_ROW ? "2樓" : "1樓");
}

function normalizePriceZone(price) {
    const value = Number(price || 0);

    if (value === 800 || value === 640) return "800";
    if (value === 500 || value === 400) return "500";
    if (value === 300 || value === 240) return "300";
    if (value === 200 || value === 160) return "200";

    return "";
}

function getOrderConcertCode(order) {
    return (
        order.concert_code ||
        order.concert ||
        order.mode ||
        order.event ||
        currentConcertMode
    );
}

function updateStatsDisplayByMode() {
    const wrapper800 = document.getElementById("stats-800-wrapper");
    const legend800 = document.getElementById("legend-800-wrapper");

    const show800 = currentConcertMode !== "tp";

    if (wrapper800) {
        wrapper800.style.display = show800 ? "" : "none";
    }

    if (legend800) {
        legend800.style.display = show800 ? "" : "none";
    }
}

function updateSeatMapButtonByMode() {
    const btn = document.getElementById("my-seat-map-btn");
    if (!btn) return;

    const main = btn.querySelector(".seat-map-card-main");
    const sub = btn.querySelector(".seat-map-card-sub");

    if (currentConcertMode === "all") {
        btn.disabled = true;
        btn.classList.add("disabled");

        if (main) main.textContent = "不支援查看座位圖";
        if (sub) sub.textContent = "請切換至台北場或高雄場查看座位";
    } else {
        btn.disabled = false;
        btn.classList.remove("disabled");

        if (main) main.textContent = "打開座位圖";
        if (sub) sub.textContent = `查看${getModeLabel()}已購買座位`;
    }
}

function updateOrdersSummary(orders) {
    orders = Array.isArray(orders) ? orders : [];

    let paidAmount = 0;
    let unpaidAmount = 0;
    let pickedCount = 0;
    let unpickedCount = 0;

    let count800 = 0;
    let count500 = 0;
    let count300 = 0;
    let count200 = 0;
    let basePoints = 0;

    for (const order of orders) {
        const price = Number(order.price || 0);
        const seats = Array.isArray(order.seats) ? order.seats : [];
        const seatCount = seats.length;

        if (order.payment_done) paidAmount += price;
        else unpaidAmount += price;

        if (order.picked_up) pickedCount += seatCount;
        else unpickedCount += seatCount;

        const priceCounts = order.price_counts || {};

        let c800 = Number(priceCounts["800"] || 0);
        let c500 = Number(priceCounts["500"] || 0);
        let c300 = Number(priceCounts["300"] || 0);
        let c200 = Number(priceCounts["200"] || 0);

        if (c800 + c500 + c300 + c200 === 0 && seatCount > 0) {
            const unitPrice = price / seatCount;

            if (unitPrice === 800 || unitPrice === 640) c800 = seatCount;
            else if (unitPrice === 500 || unitPrice === 400) c500 = seatCount;
            else if (unitPrice === 300 || unitPrice === 240) c300 = seatCount;
            else if (unitPrice === 200 || unitPrice === 160) c200 = seatCount;
        }

        count800 += c800;
        count500 += c500;
        count300 += c300;
        count200 += c200;

        basePoints += c800 * 4.5 + c500 * 2.5 + c300 * 1.5 + c200 * 1;
    }

    document.getElementById("paid-amount").textContent = formatMoney(paidAmount);
    document.getElementById("unpaid-amount").textContent = formatMoney(unpaidAmount);

    document.getElementById("picked-count").textContent = pickedCount;
    document.getElementById("unpicked-count").textContent = unpickedCount;

    document.getElementById("count-800").textContent = count800;
    document.getElementById("count-500").textContent = count500;
    document.getElementById("count-300").textContent = count300;
    document.getElementById("count-200").textContent = count200;

    const totalPointsEl = document.getElementById("total-points");
    totalPointsEl.textContent = formatNumber(basePoints);
    totalPointsEl.dataset.base = basePoints;

    document.getElementById("manual-points-note").textContent = "";

    updateStatsDisplayByMode();
}

function renderOrdersTable(orders) {
    orders = Array.isArray(orders) ? orders : [];

    const tbody = document.getElementById("orders-tbody");
    if (!tbody) return;

    if (!orders.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="empty-state">查無資料</td>
            </tr>
        `;
        updateOrdersSummary([]);
        return;
    }

    tbody.innerHTML = orders.map(order => {
        const seats = Array.isArray(order.seats) ? order.seats : [];
        const noteInputId = orderNoteInputId(order);
        const orderMode = getOrderConcertCode(order);
        const modeLabel = currentConcertMode === "all" ? `（${getModeLabel(orderMode)}）` : "";

        return `
            <tr>
                <td class="order-strong">${escapeHtml(order.datetime)}${escapeHtml(modeLabel)}</td>
                <td class="order-strong">${escapeHtml(order.floor)}${escapeHtml(order.row_label)}排</td>
                <td class="order-strong">${seats.map(seat => escapeHtml(seat)).join("、")}</td>
                <td class="order-strong">$${Number(order.price || 0).toLocaleString()}</td>

                <td>
                    <div class="note-wrap">
                        <input
                            id="${escapeHtml(noteInputId)}"
                            class="note-input"
                            type="text"
                            value="${escapeHtml(order.note || "")}"
                            placeholder="自行輸入備註"
                        />
                        <button
                            class="save-btn"
                            type="button"
                            onclick="saveNote('${escapeHtml(order.order_id)}', '${escapeHtml(order.floor)}', '${escapeHtml(order.row_label)}', '${escapeHtml(orderMode)}')"
                        >
                            儲存
                        </button>
                    </div>
                </td>

                <td>
                    <div class="cell-center">
                        ${
                            order.order_status === "locked"
                                ? `<button class="delete-btn locked" type="button" disabled title="已鎖定！">已鎖定！</button>`
                                : `<button class="delete-btn" type="button" onclick="deleteOrder('${escapeHtml(order.order_id)}', '${escapeHtml(order.floor)}', '${escapeHtml(order.row_label)}', '${escapeHtml(orderMode)}')">刪除</button>`
                        }
                    </div>
                </td>

                <td>
                    <div class="cell-center">${pickupStatusHtml(order)}</div>
                </td>
            </tr>
        `;
    }).join("");
}

async function searchOrders() {
    const input = document.getElementById("order-name-input");
    const name = input.value.trim();

    if (!name) {
        alert("請輸入姓名");
        return;
    }

    currentSearchName = name;

    try {
        showOrdersLoading("正在查詢訂單...");

        const res = await fetch(
            `/api/orders?name=${encodeURIComponent(name)}&mode=${encodeURIComponent(currentConcertMode)}`
        );

        const raw = await res.text();

        let data;

        try {
            data = JSON.parse(raw);
            console.log("discount_amount =", data.discount_amount);     
        } catch {
            console.log(raw);
            alert(`查詢回傳不是 JSON：${res.status}`);
            return;
        }

        if (!res.ok || !data.success) {
            alert(data.message || "查詢失敗");
            return;
        }

        const orders = Array.isArray(data.orders) ? data.orders : [];
        currentOrders = orders;

        currentRewardStats = {
            manual_points: Number(data.manual_points || 0),
            total_points: Number(data.total_points || 0),
            all_total_points: Number(data.all_total_points || 0),
            discount_amount: Number(data.discount_amount || 0),
            identity_code: String(data.identity_code || "5"),
            identity: data.identity || "請先查詢姓名",
            unlocked_rewards: data.unlocked_rewards || []
        };

        renderOrdersTable(orders);
        updateOrdersSummary(orders);

        const manualPoints = Number(data.manual_points || 0);

        const totalPointsEl = document.getElementById("total-points");
        const basePoints = Number(totalPointsEl.dataset.base || 0);

        const finalTotalPoints = basePoints + manualPoints;

        const discountAmount = Number(data.discount_amount || 0);
        const discountText =
            currentConcertMode === "all" && discountAmount > 0
                ? `（可折 ${formatNumber(discountAmount)} 元）`
                : "";

        totalPointsEl.textContent =
            `${formatNumber(finalTotalPoints)}${discountText}`;

        document.getElementById("manual-points-note").textContent =
            manualPoints > 0
                ? `（含手動加分 ${formatNumber(manualPoints)}）`
                : "";

        updateStatsHelpModalText();

    } finally {
        hideOrdersLoading();
        showRewardCongratModalFast();
    }
}

async function saveNote(orderId, floor, rowLabel, mode = currentConcertMode) {
    const input = document.getElementById(`note-${orderId}-${floor}-${rowLabel}`);
    const note = input ? input.value.trim() : "";

    const url =
        `/api/orders/${encodeURIComponent(orderId)}/note` +
        `?floor=${encodeURIComponent(floor)}` +
        `&row_label=${encodeURIComponent(rowLabel)}` +
        `&mode=${encodeURIComponent(mode)}`;

    const res = await fetch(url, {
        method: "PATCH",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ note })
    });

    const raw = await res.text();

    let data = {};
    try {
        data = JSON.parse(raw);
    } catch {
        alert(`回傳不是 JSON：${raw}`);
        return;
    }

    if (!res.ok || !data.success) {
        alert(data.message || "更新失敗");
        return;
    }

    alert("備註已更新");

    if (currentSearchName) searchOrders();
}

async function deleteOrder(orderId, floor, rowLabel, mode = currentConcertMode) {
    const ok = confirm("確定要刪除這一排的座位嗎？刪除後座位會重新釋出。");
    if (!ok) return;

    const url =
        `/api/orders/${encodeURIComponent(orderId)}` +
        `?floor=${encodeURIComponent(floor)}` +
        `&row_label=${encodeURIComponent(rowLabel)}` +
        `&mode=${encodeURIComponent(mode)}`;

    const res = await fetch(url, { method: "DELETE" });
    const raw = await res.text();

    let data = {};
    try {
        data = JSON.parse(raw);
    } catch {
        alert(`回傳不是 JSON：${raw}`);
        return;
    }

    if (!res.ok || !data.success) {
        alert(data.message || "刪除失敗");
        return;
    }

    alert(data.message || "刪除成功");

    if (currentSearchName) searchOrders();
}

function getSeatMapApiByMode() {
    if (currentConcertMode === "kh") {
        return "/api/kh/seats?show_third=false";
    }

    if (currentConcertMode === "tp") {
        return "/api/tp/seats";
    }

    return null;
}

async function fetchSeatMapData() {
    if (currentConcertMode === "all") {
        throw new Error("全部模式不支援查看座位圖");
    }

    if (cachedSeatMapData) {
        return cachedSeatMapData;
    }

    const api = getSeatMapApiByMode();

    const res = await fetch(api);
    const raw = await res.text();

    let data;
    try {
        data = JSON.parse(raw);
    } catch {
        console.log(raw);
        alert(`座位圖回傳不是 JSON：${res.status}`);
        throw new Error("Seat map API returned non-JSON");
    }

    cachedSeatMapData = data;

    return data;
}

function buildPurchasedSeatSet(orders) {
    const purchased = new Set();

    for (const order of orders) {
        const seats = Array.isArray(order.seats) ? order.seats : [];

        for (const seatNumber of seats) {
            purchased.add(makeSeatKey(order.floor, order.row_label, seatNumber));
        }
    }

    return purchased;
}

function renderTpMySeatMap(allSeats, rowLabels, orders) {
    const mapEl = document.getElementById("my-seat-map");
    const subtitle = document.getElementById("my-seat-map-subtitle");

    if (!mapEl) return;

    allSeats = Array.isArray(allSeats) ? allSeats : [];
    rowLabels = rowLabels || {};

    mapEl.innerHTML = "";

    if (!allSeats.length) {
        mapEl.innerHTML = `<div class="my-seat-map-empty">目前無法載入座位圖</div>`;
        return;
    }

    const purchased = buildPurchasedSeatSet(orders);

    subtitle.textContent = currentSearchName
        ? `右上角打勾代表已取票`
        : "右上角打勾代表已取票";

    const minCol = Math.min(...allSeats.map(s => Number(s.excel_col)));
    const maxCol = Math.max(...allSeats.map(s => Number(s.excel_col)));
    const minRow = Math.min(...allSeats.map(s => Number(s.excel_row)));
    const maxRow = Math.max(...allSeats.map(s => Number(s.excel_row)));

    const totalCols = (maxCol - minCol + 1) + 2;
    const topTitleOffset = 1;
    const totalRows = (maxRow - minRow + 1) + topTitleOffset;

    const gridColSize = currentConcertMode === "kh" ? 18 : 24;
    const gridRowSize = currentConcertMode === "kh" ? 18 : 22;

    mapEl.style.gridTemplateColumns =
        `${gridColSize}px repeat(${maxCol - minCol + 1}, ${gridColSize}px) ${gridColSize}px`;

    mapEl.style.gridTemplateRows =
        `repeat(${totalRows}, ${gridRowSize}px)`;

    Object.entries(rowLabels).forEach(([excelRow, label]) => {
        const displayRow = (Number(excelRow) - minRow + 1) + topTitleOffset;

        const left = document.createElement("div");
        left.className = "my-map-row-label";
        left.textContent = label;
        left.style.gridColumn = 1;
        left.style.gridRow = displayRow;
        mapEl.appendChild(left);

        const right = document.createElement("div");
        right.className = "my-map-row-label";
        right.textContent = label;
        right.style.gridColumn = totalCols;
        right.style.gridRow = displayRow;
        mapEl.appendChild(right);
    });

    for (const seat of allSeats) {
        const btn = document.createElement("button");

        const floor = floorFromSeat(seat);
        const seatKey = makeSeatKey(floor, seat.row_label, seat.seat_number);
        const priceZone = normalizePriceZone(seat.price);

        btn.type = "button";
        btn.className = "my-map-seat";
        btn.textContent = seat.seat_number || "";
        btn.disabled = true;

        btn.style.gridColumn = Number(seat.excel_col) - minCol + 2;
        btn.style.gridRow = (Number(seat.excel_row) - minRow + 1) + topTitleOffset;

        if (purchased.has(seatKey)) {
            btn.classList.add("my-seat-owned", ownedSeatClass(priceZone));

            const matchedOrder = orders.find(order => {
                const seats = Array.isArray(order.seats) ? order.seats : [];

                return (
                    order.floor === floor &&
                    order.row_label === seat.row_label &&
                    seats.includes(seat.seat_number)
                );
            });

            if (matchedOrder?.picked_up) {
                btn.innerHTML = `
                    ${seat.seat_number}
                    <span class="seat-check">✓</span>
                `;
            }

            btn.title = `${floor}${seat.row_label}排${seat.seat_number}號｜${priceZone} 元區`;
        } else if (seat.sold) {
            btn.classList.add("my-seat-other-sold");
            btn.title = "其他已售座位";
        } else {
            btn.classList.add("my-seat-not-owned");
            btn.title = "未購買座位";
        }

        mapEl.appendChild(btn);
    }
}

function renderKhMySeatMap(allSeats, rowLabels, orders) {
    const mapEl = document.getElementById("my-seat-map");
    const subtitle = document.getElementById("my-seat-map-subtitle");

    if (!mapEl) return;

    mapEl.innerHTML = "";

    if (!allSeats.length) {
        mapEl.innerHTML = `<div class="my-seat-map-empty">目前無法載入座位圖</div>`;
        return;
    }

    const purchased = buildPurchasedSeatSet(orders);

    subtitle.textContent = currentSearchName
        ? `右上角打勾代表已取票`
        : "右上角打勾代表已取票";

    const minCol = Math.min(...allSeats.map(s => Number(s.excel_col)));
    const maxCol = Math.max(...allSeats.map(s => Number(s.excel_col)));
    const minRow = Math.min(...allSeats.map(s => Number(s.excel_row)));
    const maxRow = Math.max(...allSeats.map(s => Number(s.excel_row)));

    const gridColSize = 18;
    const gridRowSize = 18;

    const TOP_TITLE_OFFSET = 0;
    
    mapEl.style.gridTemplateColumns =
        `repeat(${maxCol - minCol + 3}, ${gridColSize}px)`;
    
    mapEl.style.gridTemplateRows =
        `repeat(${maxRow - minRow + 2}, ${gridRowSize}px)`;
    
    addKhFloorFramesToMap(mapEl, minCol, minRow, TOP_TITLE_OFFSET, false);
    addKhStageToMap(mapEl, minCol, minRow, TOP_TITLE_OFFSET, {
        className: "my-map-stage kh-stage",
        text: "舞台"
    });
    addKhRowMarkersToMap(mapEl, minCol, minRow, TOP_TITLE_OFFSET, false);
    
    allSeats.forEach(seat => {
        const btn = document.createElement("button");

        const floor = seat.floor || "";
        const seatKey = makeSeatKey(floor, seat.row_label, seat.seat_number);
        const priceZone = normalizePriceZone(seat.price);

        btn.type = "button";
        btn.className = "my-map-seat";
        btn.textContent = seat.seat_number || "";
        btn.disabled = true;

        btn.style.gridColumn = Number(seat.excel_col) - minCol + 2;
        btn.style.gridRow = Number(seat.excel_row) - minRow + 1 + TOP_TITLE_OFFSET;

        if (purchased.has(seatKey)) {
            btn.classList.add("my-seat-owned", ownedSeatClass(priceZone));

            const matchedOrder = orders.find(order => {
                const seats = Array.isArray(order.seats) ? order.seats : [];

                return (
                    order.floor === floor &&
                    order.row_label === seat.row_label &&
                    seats.includes(seat.seat_number)
                );
            });

            if (matchedOrder?.picked_up) {
                btn.innerHTML = `
                    ${seat.seat_number}
                    <span class="seat-check">✓</span>
                `;
            }

            btn.title = `${floor}${seat.row_label}排${seat.seat_number}號｜${priceZone} 元區`;
        } else if (seat.sold) {
            btn.classList.add("my-seat-other-sold");
        } else {
            btn.classList.add("my-seat-not-owned");
        }

        mapEl.appendChild(btn);
    });

    enableMySeatMapZoom();
}

async function openMySeatMapModal() {
    if (currentConcertMode === "all") {
        alert("全部模式不支援查看座位圖，請切換至台北場或高雄場。");
        return;
    }

    if (!currentOrders.length) {
        alert("請先查詢姓名，並確認有訂單資料。");
        return;
    }

    const modal = document.getElementById("my-seat-map-modal");
    if (!modal) return;

    try {
        showOrdersLoading("正在載入座位圖...");

        const data = await fetchSeatMapData();

        if (currentConcertMode === "kh") {
            renderKhMySeatMap(data.seats || [], data.row_labels || {}, currentOrders);
        } else {
            renderTpMySeatMap(data.seats || [], data.row_labels || {}, currentOrders);
        }

        modal.classList.remove("hidden");

    } catch (error) {
        console.error(error);

    } finally {
        hideOrdersLoading();
    }
}

function closeMySeatMapModal() {
    const modal = document.getElementById("my-seat-map-modal");
    if (!modal) return;

    modal.classList.add("hidden");
}

function updateStatsHelpModalText() {
    const identityEl = document.getElementById("stats-identity-text");
    const discountEl = document.getElementById("stats-discount-text");

    if (identityEl) {
        identityEl.textContent =
            currentRewardStats.identity || "請先查詢姓名";
    }

    if (discountEl) {
        const discountAmount = Number(currentRewardStats.discount_amount || 0);
        discountEl.textContent = `${formatNumber(discountAmount)} 元`;
    }

    highlightCurrentIdentity(currentRewardStats.identity_code);
}

function highlightCurrentIdentity(identityCode) {
    document.querySelectorAll(".discount-th-item").forEach(item => {
        item.classList.remove("is-current-identity");

        if (String(item.dataset.identityCode) === String(identityCode)) {
            item.classList.add("is-current-identity");
        }
    });
}

function setupModeTabs() {
    document.querySelectorAll(".orders-mode-tab").forEach(btn => {
        btn.addEventListener("click", () => {
            currentConcertMode = btn.dataset.mode || "all";

            document.querySelectorAll(".orders-mode-tab").forEach(item => {
                item.classList.toggle("active", item === btn);
            });

            cachedSeatMapData = null;

            updateSeatMapButtonByMode();
            updateStatsDisplayByMode();
            updateSeatMapLegendByMode();

            if (currentSearchName) {
                searchOrders();
            }
        });
    });
}

function setupBasicEvents() {
    document.getElementById("order-search-btn")
        ?.addEventListener("click", searchOrders);

    document.getElementById("order-name-input")
        ?.addEventListener("keydown", e => {
            if (e.key === "Enter") searchOrders();
        });

    const statsHelpBtn = document.getElementById("stats-help-btn");
    const statsHelpModal = document.getElementById("stats-help-modal");
    const statsHelpClose = document.getElementById("stats-help-close");

    function openStatsHelpModal() {
        updateStatsHelpModalText();

        document.querySelectorAll(".stats-help-tab").forEach(btn => {
            btn.classList.toggle("active", btn.dataset.tab === "points");
        });

        document.querySelectorAll(".stats-help-content").forEach(content => {
            content.classList.toggle(
                "active",
                content.dataset.content === "points"
            );
        });

        statsHelpModal?.classList.remove("hidden");
    }

    function closeStatsHelpModal() {
        statsHelpModal?.classList.add("hidden");
    }

    statsHelpBtn?.addEventListener("click", openStatsHelpModal);
    statsHelpClose?.addEventListener("click", closeStatsHelpModal);

    statsHelpModal?.addEventListener("click", e => {
        if (e.target === statsHelpModal) closeStatsHelpModal();
    });

    document.querySelectorAll(".stats-help-tab").forEach(btn => {
        btn.addEventListener("click", () => {
            const tab = btn.dataset.tab;

            document.querySelectorAll(".stats-help-tab").forEach(item => {
                item.classList.toggle("active", item === btn);
            });

            document.querySelectorAll(".stats-help-content").forEach(content => {
                content.classList.toggle(
                    "active",
                    content.dataset.content === tab
                );
            });
        });
    });

    const mySeatMapBtn = document.getElementById("my-seat-map-btn");
    const mySeatMapModal = document.getElementById("my-seat-map-modal");
    const mySeatMapClose = document.getElementById("my-seat-map-close");

    mySeatMapBtn?.addEventListener("click", openMySeatMapModal);
    mySeatMapClose?.addEventListener("click", closeMySeatMapModal);

    mySeatMapModal?.addEventListener("click", e => {
        if (e.target === mySeatMapModal) closeMySeatMapModal();
    });

    const ordersHelpModal = document.getElementById("orders-help-modal");
    const ordersHelpBtn = document.getElementById("orders-help-btn");
    const ordersHelpCloseBtn = document.getElementById("orders-help-close-btn");
    const ordersHelpConfirmBtn = document.getElementById("orders-help-confirm-btn");

    function openOrdersHelpModal() {
        ordersHelpModal?.classList.remove("hidden");
    }

    function closeOrdersHelpModal() {
        ordersHelpModal?.classList.add("hidden");
    }

    ordersHelpBtn?.addEventListener("click", openOrdersHelpModal);
    ordersHelpCloseBtn?.addEventListener("click", closeOrdersHelpModal);
    ordersHelpConfirmBtn?.addEventListener("click", closeOrdersHelpModal);

    ordersHelpModal?.addEventListener("click", e => {
        if (e.target === ordersHelpModal) closeOrdersHelpModal();
    });

    document.addEventListener("keydown", e => {
        if (e.key === "Escape") {
            closeStatsHelpModal();
            closeMySeatMapModal();
            closeOrdersHelpModal();
            closeRewardCongratModal();
        }
    });
}

function updateSeatMapLegendByMode() {
    const dot500 = document.getElementById("legend-500-dot");
    if (!dot500) return;

    dot500.classList.remove("legend-500-tp", "legend-500-kh");

    if (currentConcertMode === "kh") {
        dot500.classList.add("legend-500-kh");
    } else {
        dot500.classList.add("legend-500-tp");
    }
}

function ownedSeatClass(priceZone) {
    if (priceZone === "500") {
        return currentConcertMode === "kh"
            ? "my-seat-500-kh"
            : "my-seat-500-tp";
    }

    return `my-seat-${priceZone}`;
}

function enableMySeatMapZoom() {
    const viewport = document.getElementById("my-seat-map-viewport");
    const map = document.getElementById("my-seat-map");

    if (!viewport || !map) return;

    mySeatMapScale = 1;

    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let startScrollLeft = 0;
    let startScrollTop = 0;

    viewport.onwheel = e => {
        if (!e.ctrlKey) return;

        e.preventDefault();

        const delta = e.deltaY < 0 ? 0.1 : -0.1;

        mySeatMapScale = Math.min(
            3,
            Math.max(0.45, mySeatMapScale + delta)
        );

        map.style.transform =
            `scale(${mySeatMapScale})`;
    };

    viewport.onmousedown = e => {
        isDragging = true;

        viewport.classList.add("dragging");

        startX = e.clientX;
        startY = e.clientY;

        startScrollLeft = viewport.scrollLeft;
        startScrollTop = viewport.scrollTop;
    };

    window.onmouseup = () => {
        isDragging = false;
        viewport.classList.remove("dragging");
    };

    viewport.onmousemove = e => {
        if (!isDragging) return;

        viewport.scrollLeft =
            startScrollLeft - (e.clientX - startX);

        viewport.scrollTop =
            startScrollTop - (e.clientY - startY);
    };

    /* mobile pinch */

    let initialDistance = null;
    let initialScale = mySeatMapScale;

    viewport.ontouchmove = e => {
        if (e.touches.length === 2) {
            e.preventDefault();

            const dx =
                e.touches[0].clientX - e.touches[1].clientX;

            const dy =
                e.touches[0].clientY - e.touches[1].clientY;

            const distance =
                Math.sqrt(dx * dx + dy * dy);

            if (!initialDistance) {
                initialDistance = distance;
                initialScale = mySeatMapScale;
                return;
            }

            mySeatMapScale =
                Math.min(
                    3,
                    Math.max(
                        0.45,
                        initialScale * (distance / initialDistance)
                    )
                );

            map.style.transform =
                `scale(${mySeatMapScale})`;
        }
    };

    viewport.ontouchend = () => {
        initialDistance = null;
    };
}

// =========================================================
// 💡 ⚡ 推票獎勵祝賀 Modal（0 延遲完整補齊版）
// =========================================================

function showRewardCongratModalFast() {
    if (!currentSearchName) return;

    const unlocked = currentRewardStats.unlocked_rewards || [];
    const allTotalPts = currentRewardStats.all_total_points || 0;

    if (unlocked.length === 0) return;

    const myUnlockedRewards = unlocked.map(item => ({
        name: item.name,
        requirement: item.requirement,
        desc: getRewardGuidelineText(item.name),
        icon: getRewardIcon(item.name)
    }));

    renderRewardCongratModal(allTotalPts, myUnlockedRewards);
}
// 3. 渲染 Modal HTML 內容
function renderRewardCongratModal(allTotalPts, rewards) {
    const modal = document.getElementById("reward-congrat-modal");
    const ptsEl = document.getElementById("reward-modal-points");
    const itemsContainer = document.getElementById("reward-modal-items");

    if (!modal || !itemsContainer) return;

    if (ptsEl) ptsEl.textContent = formatNumber(allTotalPts);

    itemsContainer.innerHTML = rewards.map(item => `
        <div class="reward-item-card">
            <div class="reward-item-icon">${item.icon}</div>
            <div>
                <div class="reward-item-title">${escapeHtml(cleanEmoji(item.name))} <span style="font-size:12px; color:#888; font-weight:normal;">(${escapeHtml(item.requirement)})</span></div>
                <div class="reward-item-desc">${escapeHtml(item.desc)}</div>
            </div>
        </div>
    `).join("");

    modal.classList.remove("hidden");
}

function closeRewardCongratModal() {
    const modal = document.getElementById("reward-congrat-modal");
    modal?.classList.add("hidden");
}

function setupRewardCongratEvents() {
    const modal = document.getElementById("reward-congrat-modal");
    const closeBtn = document.getElementById("reward-congrat-close");
    const confirmBtn = document.getElementById("reward-congrat-confirm-btn");

    closeBtn?.addEventListener("click", closeRewardCongratModal);
    confirmBtn?.addEventListener("click", closeRewardCongratModal);

    modal?.addEventListener("click", (e) => {
        if (e.target === modal) closeRewardCongratModal();
    });
}

// 初始化綁定與設定
setupModeTabs();
setupBasicEvents();
setupRewardCongratEvents();

updateSeatMapButtonByMode();
updateStatsDisplayByMode();
updateSeatMapLegendByMode();