// ==UserScript==
// @name         MQC 資料暫存與統計儀表板
// @namespace    http://tampermonkey.net/
// @version      22.0
// @description  自動讀取 PVIM5242 Frame 的下拉選單，將卡片標題替換為完整商品名稱。
// @author       Ken
// @match        https://appsvr12.panasonic.com.tw/VIMS/PVIM5241.asp*
// @match        https://appsvr12.panasonic.com.tw/VIMS/PVIM5243.asp*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @license      All rights reserved
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // --- 設定 ---
    const DB_BASE_NAME = 'mqc_record';
    const STORE_NAME = 'reasons';
    const DB_VERSION = 1;
    const MAPPING_KEY = 'mqc_product_map_cache'; // 用來存商品名稱對照表

    // --- 注入 CSS ---
    function injectStyles() {
        const style = document.createElement('style');
        style.innerHTML = `
            .mqc-switch { position: relative; display: inline-block; width: 34px; height: 18px; vertical-align: middle; }
            .mqc-switch input { opacity: 0; width: 0; height: 0; }
            .mqc-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 18px; }
            .mqc-slider:before { position: absolute; content: ""; height: 14px; width: 14px; left: 2px; bottom: 2px; background-color: white; transition: .4s; border-radius: 50%; }
            input:checked + .mqc-slider { background-color: #2196F3; }
            input:checked + .mqc-slider:before { transform: translateX(16px); }

            .mqc-card-input {
                border: 1px solid #ddd; border-radius: 4px; padding: 4px;
                font-size: 13px; color: #333; width: 100%; box-sizing: border-box;
                font-family: Arial, sans-serif;
            }
            .mqc-card-input:disabled { background: #eee; color: #aaa; cursor: not-allowed; }

            .mqc-btn { border: none; border-radius: 4px; cursor: pointer; color: white; font-size: 13px; padding: 5px 0; transition: opacity 0.2s; }
            .mqc-btn:hover { opacity: 0.9; }
            .mqc-btn-blue { background: #007bff; }
            .mqc-btn-info { background: #17a2b8; }
            .mqc-btn-gray { background: #6c757d; font-size: 11px; padding: 3px 0; }
            .mqc-btn-green { background: #28a745; font-size: 11px; padding: 3px 0; }
        `;
        document.head.appendChild(style);
    }
    injectStyles();

    // --- 核心工具 ---
    function getDynamicDBName(doc) {
        try {
            const loc = doc.location || window.location;
            const urlParams = new URLSearchParams(loc.search);
            const ord1 = urlParams.get('ORD1');
            if (ord1 && ord1.trim() !== '') {
                return `${DB_BASE_NAME}_${ord1.trim()}`;
            }
        } catch (e) { }
        return null;
    }

    function getTodayString() {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function getFirstDayOfMonthString() {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        return `${yyyy}-${mm}-01`;
    }

    // --- 商品名稱同步工具 (Frame Interop) ---
    function syncProductMapping() {
        try {
            // 嘗試取得 PVIM5242 Frame
            const siblingFrame = window.top.frames['PVIM5242'];
            if (!siblingFrame) return;

            const doc = siblingFrame.document;
            const select = doc.getElementById('ORD1');

            if (select && select.options.length > 0) {
                const mapping = {};
                // 讀取既有的 Cache，避免覆蓋掉沒抓到的部分(雖然通常是一次全抓)
                const existing = localStorage.getItem(MAPPING_KEY);
                if (existing) Object.assign(mapping, JSON.parse(existing));

                // 解析 Select Options
                for (let i = 0; i < select.options.length; i++) {
                    const opt = select.options[i];
                    const val = opt.value.trim();
                    const txt = opt.text.trim(); // 例如 "V1VRF冷氣進口"
                    if (val) {
                        mapping[val] = txt;
                    }
                }

                localStorage.setItem(MAPPING_KEY, JSON.stringify(mapping));
            }
        } catch (e) {
            // 跨域或 Frame 尚未載入完成，忽略錯誤
        }
    }

    function getProductName(code) {
        // 從 Cache 讀取名稱，若無則回傳原本代碼
        try {
            const json = localStorage.getItem(MAPPING_KEY);
            if (json) {
                const map = JSON.parse(json);
                if (map[code]) return map[code];
            }
        } catch (e) { }
        return code;
    }

    // --- IndexedDB 工具 ---
    const dbUtils = {
        open: (dbName) => new Promise((resolve, reject) => {
            const request = indexedDB.open(dbName, DB_VERSION);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'id' });
                }
            };
            request.onsuccess = (e) => resolve(e.target.result);
            request.onerror = (e) => reject(e);
        }),
        putBatch: async (dbName, dataArray) => {
            if (dataArray.length === 0) return;
            const db = await dbUtils.open(dbName);
            return new Promise((resolve, reject) => {
                const tx = db.transaction([STORE_NAME], 'readwrite');
                const store = tx.objectStore(STORE_NAME);
                tx.oncomplete = () => resolve(true);
                tx.onerror = () => reject('Transaction error');
                dataArray.forEach(item => store.put(item));
            });
        },
        get: async (dbName, key) => {
            const db = await dbUtils.open(dbName);
            return new Promise((resolve, reject) => {
                const tx = db.transaction([STORE_NAME], 'readonly');
                const store = tx.objectStore(STORE_NAME);
                const req = store.get(key);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject('Get error');
            });
        },
        getAllAsMap: async (dbName) => {
            const db = await dbUtils.open(dbName);
            return new Promise((resolve, reject) => {
                const tx = db.transaction([STORE_NAME], 'readonly');
                const store = tx.objectStore(STORE_NAME);
                const req = store.getAll();
                req.onsuccess = () => {
                    const map = new Map();
                    if (req.result) {
                        req.result.forEach(item => map.set(item.id, item));
                    }
                    resolve(map);
                };
                req.onerror = () => reject('GetAll error');
            });
        },
        getAll: async (dbName) => {
            const db = await dbUtils.open(dbName);
            return new Promise((resolve, reject) => {
                const tx = db.transaction([STORE_NAME], 'readonly');
                const store = tx.objectStore(STORE_NAME);
                const req = store.getAll();
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject('GetAll error');
            });
        },
        getStats: async (dbName, startDate, endDate) => {
            const db = await dbUtils.open(dbName);
            return new Promise((resolve, reject) => {
                const tx = db.transaction([STORE_NAME], 'readonly');
                const store = tx.objectStore(STORE_NAME);
                const request = store.openCursor();
                let total = 0; let filled = 0;
                const start = new Date(startDate); start.setHours(0, 0, 0, 0);
                const end = new Date(endDate); end.setHours(23, 59, 59, 999);
                request.onsuccess = (event) => {
                    const cursor = event.target.result;
                    if (cursor) {
                        const record = cursor.value;
                        if (record.timestamp) {
                            const rDate = new Date(record.timestamp);
                            if (rDate >= start && rDate <= end) {
                                total++;
                                if (record.history && record.history.trim() !== '') {
                                    filled++;
                                }
                            }
                        }
                        cursor.continue();
                    } else { resolve({ total, filled }); }
                };
                request.onerror = () => reject('Cursor error');
            });
        }
    };

    // --- 建立單一卡片的 DOM ---
    function createCard(dbName) {
        const suffix = dbName.replace('mqc_record_', '');
        const cardId = `mqc-card-${suffix}`;

        // 取得完整商品名稱
        const displayName = getProductName(suffix);

        // 如果已存在，檢查是否需要更新標題名稱 (應對非同步載入)
        const existingCard = document.getElementById(cardId);
        if (existingCard) {
            const titleSpan = existingCard.querySelector('.mqc-card-title');
            if (titleSpan && titleSpan.innerText !== `📂 ${displayName}`) {
                titleSpan.innerText = `📂 ${displayName}`;
            }
            return null;
        }

        const card = document.createElement('div');
        card.id = cardId;
        card.style.cssText = `
            background: #fff;
            border: 1px solid #ccc;
            border-radius: 6px;
            padding: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex; flex-direction: column; gap: 8px;
            font-size: 13px; color: #333;
            transition: transform 0.2s;
        `;
        card.onmouseover = () => card.style.transform = "translateY(-3px)";
        card.onmouseout = () => card.style.transform = "translateY(0)";

        const html = `
            <div style="font-weight:bold; color:#0056b3; border-bottom:1px solid #eee; padding-bottom:5px; margin-bottom:0px; display:flex; justify-content:space-between; align-items:center;">
                <span class="mqc-card-title" style="font-size:14px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:110px;" title="${displayName}">📂 ${displayName}</span>
                <div style="display:flex; align-items:center; gap:5px;">
                    <span style="font-size:12px; color:#666;">今日</span>
                    <label class="mqc-switch">
                        <input type="checkbox" class="mqc-toggle-today">
                        <span class="mqc-slider"></span>
                    </label>
                </div>
            </div>

            <div style="display:flex; flex-direction:column; gap:6px;">
                <div style="display:flex; align-items:center; gap:5px;">
                    <span style="width:20px; text-align:right; font-weight:bold; color:#555;">起</span>
                    <input type="date" class="mqc-start mqc-card-input">
                </div>
                <div style="display:flex; align-items:center; gap:5px;">
                    <span style="width:20px; text-align:right; font-weight:bold; color:#555;">迄</span>
                    <input type="date" class="mqc-end mqc-card-input">
                </div>
            </div>

            <div style="background:#f8f9fa; padding:8px; border-radius:4px; margin-top:2px;">
                <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                    <span>總數:</span><strong class="mqc-total" style="font-size:14px;">0</strong>
                </div>
                <div style="display:flex; justify-content:space-between; color:#28a745; margin-bottom:4px;">
                    <span>撤回:</span><strong class="mqc-filled" style="font-size:14px;">0</strong>
                </div>
                <div style="background:#e9ecef; height:8px; border-radius:4px; overflow:hidden;">
                    <div class="mqc-bar" style="width:0%; height:100%; background:#28a745; transition:width 0.5s;"></div>
                </div>
                <div class="mqc-percent" style="text-align:right; font-size:12px; margin-top:4px; color:#666; font-weight:bold;">0%</div>
            </div>

            <div style="display:flex; gap:5px;">
                <button class="mqc-refresh-btn mqc-btn mqc-btn-blue" style="flex:1;">更新</button>
                <button class="mqc-export-xlsx mqc-btn mqc-btn-info" style="flex:2;">Excel</button>
            </div>
             <div style="display:flex; gap:5px; margin-top:2px;">
                <button class="mqc-json-export mqc-btn mqc-btn-gray" style="flex:1;">備份</button>
                <button class="mqc-json-import mqc-btn mqc-btn-green" style="flex:1;">還原</button>
            </div>
        `;
        card.innerHTML = html;

        const startInput = card.querySelector('.mqc-start');
        const endInput = card.querySelector('.mqc-end');
        const toggleToday = card.querySelector('.mqc-toggle-today');
        const totalEl = card.querySelector('.mqc-total');
        const filledEl = card.querySelector('.mqc-filled');
        const barEl = card.querySelector('.mqc-bar');
        const percentEl = card.querySelector('.mqc-percent');

        startInput.value = getFirstDayOfMonthString();
        endInput.value = getTodayString();

        // 取得當前要使用的統計區間
        const getDateRange = () => {
            if (toggleToday.checked) {
                return { start: getTodayString(), end: getTodayString() };
            } else {
                return { start: startInput.value, end: endInput.value };
            }
        };

        const refresh = async () => {
            try {
                if (toggleToday.checked) {
                    startInput.disabled = true;
                    endInput.disabled = true;
                } else {
                    startInput.disabled = false;
                    endInput.disabled = false;
                }

                const range = getDateRange();
                const stats = await dbUtils.getStats(dbName, range.start, range.end);

                totalEl.innerText = stats.total;
                filledEl.innerText = stats.filled;
                let p = 0;
                if (stats.total > 0) p = Math.round((stats.filled / stats.total) * 100);
                barEl.style.width = p + '%';
                percentEl.innerText = p + '%';
            } catch(e) { console.error(e); }
        };

        toggleToday.onchange = refresh;
        card.querySelector('.mqc-refresh-btn').onclick = refresh;

        card.querySelector('.mqc-export-xlsx').onclick = async () => {
             try {
                const range = getDateRange();
                const sDate = new Date(range.start); sDate.setHours(0,0,0,0);
                const eDate = new Date(range.end); eDate.setHours(23,59,59,999);

                const allData = await dbUtils.getAll(dbName);
                if (allData.length === 0) { alert('資料庫無資料'); return; }

                const filteredData = allData.filter(item => {
                    if (!item.timestamp) return false;
                    const itemDate = new Date(item.timestamp);
                    return itemDate >= sDate && itemDate <= eDate;
                });

                if (filteredData.length === 0) { alert('此區間無資料可匯出'); return; }

                const excelData = filteredData.map(item => ({
                    '管號': item.id,
                    '撤回紀錄': item.history || '',
                    '最後確認時間': item.timestamp ? new Date(item.timestamp).toLocaleString() : ''
                }));

                const worksheet = XLSX.utils.json_to_sheet(excelData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "MQC_Data");

                const fileNameDate = toggleToday.checked ? `Today_${getTodayString()}` : `${range.start}_to_${range.end}`;
                XLSX.writeFile(workbook, `MQC_${suffix}_${fileNameDate}.xlsx`);

            } catch (e) { console.error(e); alert('匯出失敗'); }
        };

        card.querySelector('.mqc-json-export').onclick = async () => {
            try {
                const d = await dbUtils.getAll(dbName);
                const b = new Blob([JSON.stringify(d, null, 2)], { type: "application/json" });
                const a = document.createElement('a'); a.href = URL.createObjectURL(b);
                a.download = `MQC_${suffix}_Backup_${getTodayString()}.json`; a.click();
            } catch(e) {}
        };

        card.querySelector('.mqc-json-import').onclick = () => {
             const i = document.createElement('input'); i.type = 'file'; i.accept = '.json';
            i.onchange = e => {
                const r = new FileReader();
                r.onload = async ev => {
                    try {
                        const j = JSON.parse(ev.target.result);
                        if (confirm(`還原 ${j.length}筆至 ${suffix}?`)) {
                            await dbUtils.putBatch(dbName, j); refresh(); alert('OK');
                        }
                    } catch (x) { alert('失敗'); }
                };
                r.readAsText(e.target.files[0]);
            };
            i.click();
        };

        setTimeout(refresh, 500);
        return card;
    }

    async function renderGlobalGrid() {
        let container = document.getElementById('mqc-grid-container');
        if (!container) {
            container = document.createElement('div');
            container.id = 'mqc-grid-container';
            container.style.cssText = `
                position: absolute;
                top: 60px; left: 0; right: 0;
                padding: 15px;
                display: grid;
                grid-template-columns: repeat(auto-fit, 200px);
                justify-content: center;
                gap: 15px;
                pointer-events: none;
                z-index: 9999;
            `;
            document.body.appendChild(container);
        }

        let dbs = [];
        try {
            if (window.indexedDB && window.indexedDB.databases) {
                const allDbs = await window.indexedDB.databases();
                dbs = allDbs.filter(db => db.name.startsWith(DB_BASE_NAME)).map(db => db.name);
            }
        } catch (e) { console.warn("瀏覽器不支援列舉資料庫"); return; }

        if (dbs.length === 0) return;

        dbs.forEach(dbName => {
            const card = createCard(dbName);
            if (card) {
                card.style.pointerEvents = "auto";
                container.appendChild(card);
            }
        });
    }

    async function initSpecificDashboard(currentDBName) {
        if (document.getElementById('mqc-dashboard')) return;
        const div = document.createElement('div');
        div.id = 'mqc-dashboard';
        div.style.cssText = `
            position: fixed; top: 10px; right: 20px;
            width: 40px; height: 40px; background: #007bff; color: white;
            border-radius: 50%; box-shadow: 0 4px 8px rgba(0,0,0,0.3);
            z-index: 10000; cursor: pointer; display: flex; justify-content: center; align-items: center;
            font-size: 20px;
        `;
        div.innerText = '📊';
        div.onclick = () => {
             const existing = document.getElementById('mqc-single-panel');
             if(existing) { existing.remove(); return; }

             const card = createCard(currentDBName);
             card.id = 'mqc-single-panel';
             card.style.position = 'fixed';
             card.style.top = '60px';
             card.style.right = '20px';
             card.style.zIndex = '10001';
             card.style.width = '200px';
             document.body.appendChild(card);
        };
        document.body.appendChild(div);
    }

    // --- 頁面邏輯 ---
    async function initPageLogic(currentDBName) {
        const countInput = document.getElementById('reCnt');
        if (!countInput) return;
        const count = parseInt(countInput.value, 10);

        for (let i = 1; i <= count; i++) {
            const elCtl = document.getElementById('SSHCTLNO' + i);
            const elWhy = document.getElementById('SSHWHY' + i);
            const elCheck = document.getElementById('check' + i);

            if (elCtl && elWhy) {
                const key = elCtl.value.trim();
                const tdCtl = elCtl.closest('td');
                try {
                    const record = await dbUtils.get(currentDBName, key);
                    if (record) {
                        if (record.reason && elWhy.value === '' && (!elCheck || !elCheck.checked)) {
                            elWhy.value = record.reason;
                            elWhy.style.backgroundColor = "#e6ffe6";
                        }
                        if (record.history && record.history.trim() !== '') {
                            if (tdCtl) tdCtl.style.backgroundColor = '#fce6de';
                            elCtl.title = `撤回紀錄: ${record.history}`;
                        }
                    }
                } catch (e) { }
            }
        }

        for (let i = 1; i <= count; i++) {
            const elCheck = document.getElementById('check' + i);
            const elWhy = document.getElementById('SSHWHY' + i);
            const elCtl = document.getElementById('SSHCTLNO' + i);
            if (elCheck && elWhy && elCtl && !elCheck.dataset.mqcListened) {
                elCheck.dataset.mqcListened = "true";
                elCheck.addEventListener('change', async function () {
                    if (this.checked) {
                        const currentText = elWhy.value;
                        const key = elCtl.value.trim();
                        if (key && currentText && currentText.trim() !== '') {
                            const oldRecord = await dbUtils.get(currentDBName, key);
                            let finalHistory = currentText;
                            if (oldRecord && oldRecord.history && oldRecord.history.trim() !== '') {
                                finalHistory = oldRecord.history;
                            }
                            await dbUtils.putBatch(currentDBName, [{
                                id: key,
                                reason: currentText,
                                history: finalHistory,
                                timestamp: new Date()
                            }]);
                        }
                        elWhy.value = '';
                        elWhy.style.backgroundColor = '';
                    }
                });
            }
        }
    }

    // =========================================================
    // 主流程
    // =========================================================
    let hasInitGrid = false;

    setInterval(() => {
        // 每秒嘗試同步一次商品名稱 (因為 Frame 載入時間不確定)
        syncProductMapping();

        const currentDBName = getDynamicDBName(document);
        const is5243 = window.location.href.indexOf('PVIM5243.asp') > -1;

        if (currentDBName) {
            // [模式 A] 特定工單模式
            const grid = document.getElementById('mqc-grid-container');
            if (grid) grid.style.display = 'none';

            initSpecificDashboard(currentDBName);
            initPageLogic(currentDBName);
        } else {
            // [模式 B] 全域 Grid 模式 (僅 5243)
            if (is5243) {
                if (!hasInitGrid) {
                    renderGlobalGrid();
                    hasInitGrid = true;
                }
                const grid = document.getElementById('mqc-grid-container');
                if (grid) grid.style.display = 'grid';
            }
        }
    }, 1000);

    // --- 存檔攔截 ---
    function findDataFrame(currentWindow) {
        try { const doc = currentWindow.document; if (doc && doc.getElementById('reCnt')) return currentWindow; } catch (e) { }
        for (let i = 0; i < currentWindow.frames.length; i++) { const found = findDataFrame(currentWindow.frames[i]); if (found) return found; }
        return null;
    }

    const btnSend = document.getElementById('button2');
    if (btnSend && (btnSend.value.indexOf("確認傳送") > -1 || btnSend.value.indexOf("確認") > -1)) {
        if (!btnSend.dataset.mqcAttached) {
            btnSend.dataset.mqcAttached = "true";
            btnSend.addEventListener('click', async function (e) {
                const targetWin = findDataFrame(window.top);
                if (!targetWin) return;
                const doc = targetWin.document;
                const targetDBName = getDynamicDBName(doc);
                if (!targetDBName) return;

                console.log('[MQC] 存檔中...');
                const originalBg = btnSend.style.backgroundColor;
                btnSend.style.backgroundColor = '#00aaff';
                setTimeout(() => { btnSend.style.backgroundColor = originalBg; }, 200);

                const count = parseInt(doc.getElementById('reCnt').value, 10);
                const backupData = [];
                let dbMap = new Map();
                try { dbMap = await dbUtils.getAllAsMap(targetDBName); } catch (err) { }

                for (let i = 1; i <= count; i++) {
                    const elCtl = doc.getElementById('SSHCTLNO' + i);
                    const elWhy = doc.getElementById('SSHWHY' + i);
                    if (elCtl && elWhy) {
                        const key = elCtl.value.trim();
                        const val = elWhy.value;
                        if (key) {
                            const oldRecord = dbMap.get(key);
                            let finalHistory = '';
                            if (oldRecord && oldRecord.history !== undefined) finalHistory = oldRecord.history;
                            else finalHistory = val;
                            if (finalHistory.trim() === '' && val.trim() !== '') finalHistory = val;

                            backupData.push({
                                id: key,
                                reason: val,
                                history: finalHistory,
                                timestamp: new Date()
                            });
                        }
                    }
                }
                if (backupData.length > 0) {
                    dbUtils.putBatch(targetDBName, backupData);
                }
            }, true);
        }
    }
})();
