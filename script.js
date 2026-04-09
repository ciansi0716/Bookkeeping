// 【重要】請替換為你最新部署的 GAS URL
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxAlklJVn9ibCGzpOtSF0JnN3mCGVD5Iaw3aYsDcnNUE_Kn6i27qEgOuBWg5JpOK4xTqA/exec';

let appData = { '帳本': [], '使用者': [], '大分類': [], '小分類': [], '專案': [] };
let validExpenseData = []; 
let exceptionExpenseData = []; 

let currentChartInOut = '支出'; 
let currentChartTime = '本月';
let currentCalDate = new Date();
let selectedDateStr = formatDate(new Date());
let currentChartDetailLabel = '';

let compareMode = 'year'; 
let compareSelectedPeriods = [];
const compareColors = ['#F44336', '#2196F3', '#4CAF50', '#FF9800', '#9C27B0'];

document.addEventListener('DOMContentLoaded', function() {
    initUI();
    initDateSelects(); 
    fetchData(); 
});

function formatDate(d) {
    let month = '' + (d.getMonth() + 1), day = '' + d.getDate(), year = d.getFullYear();
    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;
    return [year, month, day].join('-');
}

function getRandomColor() {
    const letters = '0123456789ABCDEF';
    let color = '#';
    for (let i = 0; i < 6; i++) color += letters[Math.floor(Math.random() * 16)];
    return color;
}

function getUserColorStyle(username) {
    let userObj = appData['使用者'].find(u => u.name === username);
    let hex = (userObj && userObj.color) ? userObj.color : '#1976d2'; 
    if(/^#[0-9A-F]{6}$/i.test(hex)) {
        let r = parseInt(hex.slice(1, 3), 16);
        let g = parseInt(hex.slice(3, 5), 16);
        let b = parseInt(hex.slice(5, 7), 16);
        return `background-color: rgba(${r}, ${g}, ${b}, 0.15); color: ${hex}; border: 1px solid rgba(${r}, ${g}, ${b}, 0.3);`;
    }
    return `background-color: #e3f2fd; color: #1976d2; border: 1px solid #bbdefb;`;
}

function getIcon(categoryName, type) {
    if (type === '大分類') {
        let cat = appData['大分類'].find(c => c.name === categoryName);
        return cat && cat.icon ? cat.icon : '📁';
    } else if (type === '小分類') {
        let sub = appData['小分類'].find(c => c.name === categoryName);
        if (sub) {
            let parentCat = appData['大分類'].find(c => c.name === sub.parentName);
            return parentCat && parentCat.icon ? parentCat.icon : '📁';
        }
        return '🏷️';
    }
    return '📁';
}

function renderCalendar() {
    const grid = document.getElementById('calendarGrid');
    if(!grid) return;
    grid.innerHTML = '';
    
    const calY = document.getElementById('calYearSelect');
    const calM = document.getElementById('calMonthSelect');
    if(calY) calY.value = currentCalDate.getFullYear();
    if(calM) calM.value = currentCalDate.getMonth() + 1;

    ['日','一','二','三','四','五','六'].forEach(day => {
        let el = document.createElement('div'); 
        el.innerText = day; el.style.fontWeight = 'bold'; el.style.padding = '5px 0'; el.style.color = '#666';
        grid.appendChild(el);
    });
    
    let firstDay = new Date(currentCalDate.getFullYear(), currentCalDate.getMonth(), 1).getDay();
    let daysInMonth = new Date(currentCalDate.getFullYear(), currentCalDate.getMonth() + 1, 0).getDate();
    
    for(let i=0; i<firstDay; i++) grid.appendChild(document.createElement('div'));
    
    const realTodayStr = formatDate(new Date());
    let recordDates = new Set(validExpenseData.map(e => e.date));

    for(let i=1; i<=daysInMonth; i++) {
        let cellDate = new Date(currentCalDate.getFullYear(), currentCalDate.getMonth(), i);
        let cellDateStr = formatDate(cellDate);
        let el = document.createElement('div'); 
        el.className = 'cal-day'; 
        
        el.innerHTML = `<span>${i}</span>`;
        if (recordDates.has(cellDateStr)) {
            el.innerHTML += `<div class="cal-dot"></div>`;
        }
        
        if(cellDateStr === realTodayStr) el.classList.add('today');
        if(cellDateStr === selectedDateStr) el.classList.add('selected');
        
        el.onclick = () => {
            selectedDateStr = cellDateStr;
            renderCalendar(); 
            renderDailyList(); 
        };
        grid.appendChild(el);
    }
}

function initUI() {
    document.querySelectorAll('.bottom-nav .nav-item').forEach(item => {
        item.onclick = function() {
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            this.classList.add('active');
            
            const target = this.getAttribute('data-target');
            document.getElementById(target).classList.add('active');

            const titleMap = { 'page-home': '日概況', 'page-charts': '統計分析', 'page-manage': '項目管理', 'page-exceptions': '異常資料', 'page-compare': '數據比較' };
            document.getElementById('headerTitle').innerText = titleMap[target] || '';
            
            if(target === 'page-charts') renderChart();
            if(target === 'page-manage') renderManageLists();
            if(target === 'page-exceptions') renderExceptionList();
            if(target === 'page-compare') {
                initCompareDropdowns();
                renderComparePage();
            }
        };
    });

    document.getElementById('backToChartBtn').onclick = () => {
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.getElementById('page-charts').classList.add('active');
        document.getElementById('headerTitle').innerText = '統計分析';
    };

    document.querySelectorAll('.manage-tabs .tab-item').forEach(tab => {
        tab.onclick = function() {
            document.querySelectorAll('.manage-tabs .tab-item').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.manage-content').forEach(c => c.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(this.getAttribute('data-target')).classList.add('active');
        };
    });

    document.getElementById('prevMonth').onclick = () => { currentCalDate.setMonth(currentCalDate.getMonth() - 1); renderCalendar(); };
    document.getElementById('nextMonth').onclick = () => { currentCalDate.setMonth(currentCalDate.getMonth() + 1); renderCalendar(); };

    const inOutSel = document.getElementById('chartInOutSelect');
    const timeSel = document.getElementById('chartTimeSelect');
    const yearSel = document.getElementById('chartYearSelect');
    const monthSel = document.getElementById('chartMonthSelect');
    const userSel = document.getElementById('chartUserSelect');
    const targetUserSel = document.getElementById('chartTargetUserSelect');
    const projectSel = document.getElementById('chartProjectSelect');

    [inOutSel, userSel, targetUserSel, monthSel, projectSel].forEach(el => {
        if(el) el.addEventListener('change', renderChart);
    });

    if(timeSel) {
        timeSel.addEventListener('change', () => {
            currentChartTime = timeSel.value;
            if (currentChartTime === '歷年') {
                yearSel.style.display = 'none'; monthSel.style.display = 'none';
            } else if (currentChartTime === '當年') {
                yearSel.style.display = 'inline-block'; monthSel.style.display = 'none';
            } else { 
                yearSel.style.display = 'inline-block'; monthSel.style.display = 'inline-block';
            }
            renderChart();
        });
    }
    
    if(yearSel) {
        yearSel.addEventListener('change', () => {
            updateChartMonthDropdown();
            renderChart();
        });
    }

    const compModeSel = document.getElementById('compareModeSelect');
    const compStartYear = document.getElementById('compareStartYear');
    const compEndYear = document.getElementById('compareEndYear');
    const compMonth = document.getElementById('compareMonth');

    if(compModeSel) {
        compModeSel.addEventListener('change', () => {
            compareMode = compModeSel.value;
            compMonth.style.display = compareMode === 'month' ? 'inline-block' : 'none';
            renderComparePage();
        });
    }
    [compStartYear, compEndYear, compMonth].forEach(el => {
        if(el) el.addEventListener('change', renderComparePage);
    });

    const addModalUserSel = document.getElementById('user');
    const addModalTargetSel = document.getElementById('targetUser');
    if (addModalUserSel && addModalTargetSel) {
        addModalUserSel.addEventListener('change', function() { addModalTargetSel.value = this.value; });
    }

    document.getElementById('fabAddBtn').onclick = () => {
        document.getElementById('expenseForm').reset();
        document.getElementById('expenseId').value = Date.now();
        document.getElementById('expenseOperation').value = 'add';
        document.getElementById('submitBtn').innerText = '儲存新增';
        document.getElementById('date').value = selectedDateStr; 
        updateFormDropdowns();
        document.getElementById('addModal').style.display = 'flex';
    };
    
    document.getElementById('closeModalBtn').onclick = () => document.getElementById('addModal').style.display = 'none';
    window.onclick = function(e) { if (e.target.className === 'modal') e.target.style.display = 'none'; };

    const globalSel = document.getElementById('globalLedgerSelect');
    if(globalSel) globalSel.addEventListener('change', fetchData);

    const mainCatSel = document.getElementById('mainCategory');
    if(mainCatSel) mainCatSel.addEventListener('change', () => updateFormSubCategory());
}

function initDateSelects() {
    const now = new Date();
    const currentY = now.getFullYear();
    
    const calY = document.getElementById('calYearSelect');
    const calM = document.getElementById('calMonthSelect');
    if (calY && calM) {
        let yHtml = '';
        for(let y = currentY - 5; y <= currentY + 5; y++) yHtml += `<option value="${y}">${y}年</option>`;
        calY.innerHTML = yHtml;
        calY.value = currentY;

        let mHtml = '';
        for(let m = 1; m <= 12; m++) mHtml += `<option value="${m}">${m}月</option>`;
        calM.innerHTML = mHtml;
        calM.value = now.getMonth() + 1;

        calY.addEventListener('change', () => { currentCalDate.setFullYear(parseInt(calY.value, 10)); renderCalendar(); });
        calM.addEventListener('change', () => { currentCalDate.setMonth(parseInt(calM.value, 10) - 1); renderCalendar(); });
    }
}

function initCompareDropdowns() {
    const sY = document.getElementById('compareStartYear');
    const eY = document.getElementById('compareEndYear');
    const cM = document.getElementById('compareMonth');
    if(!sY || !eY || !cM) return;

    let uniqueYears = new Set();
    validExpenseData.forEach(e => {
        if(e.date) { let parts = e.date.split('-'); if(parts.length > 0) uniqueYears.add(parts[0]); }
    });
    let yearsArray = Array.from(uniqueYears).sort((a,b) => a - b); 

    if (yearsArray.length === 0) {
        sY.innerHTML = `<option value="">無</option>`;
        eY.innerHTML = `<option value="">無</option>`;
        return;
    }

    let yHtml = yearsArray.map(y => `<option value="${y}">${y}年</option>`).join('');
    
    let curStart = sY.value;
    let curEnd = eY.value;

    sY.innerHTML = yHtml;
    eY.innerHTML = yHtml;

    if (yearsArray.includes(curStart)) sY.value = curStart;
    else sY.value = yearsArray.length > 1 ? yearsArray[yearsArray.length - 2] : yearsArray[0];

    if (yearsArray.includes(curEnd)) eY.value = curEnd;
    else eY.value = yearsArray[yearsArray.length - 1];
    
    if (!cM.value) cM.value = new Date().getMonth() + 1;
}

function updateChartDropdowns() {
    const ys = document.getElementById('chartYearSelect');
    const ms = document.getElementById('chartMonthSelect');
    const userSel = document.getElementById('chartUserSelect');
    const targetUserSel = document.getElementById('chartTargetUserSelect');
    const projectSel = document.getElementById('chartProjectSelect');
    
    if(!ys || !ms || !userSel || !targetUserSel || !projectSel) return;

    const currentSelectedUser = userSel.value || '全付款人';
    let uniqueUsers = new Set();
    validExpenseData.forEach(e => { if(e.user) uniqueUsers.add(e.user); });
    userSel.innerHTML = `<option value="全付款人">全付款人</option>`;
    Array.from(uniqueUsers).sort().forEach(u => userSel.innerHTML += `<option value="${u}">${u}</option>`);
    if (uniqueUsers.has(currentSelectedUser)) userSel.value = currentSelectedUser; else userSel.value = '全付款人';

    const currentSelectedTarget = targetUserSel.value || '全被使用者';
    let uniqueTargets = new Set();
    validExpenseData.forEach(e => { let t = e.targetUser || e.user; if(t) uniqueTargets.add(t); });
    targetUserSel.innerHTML = `<option value="全被使用者">全被使用者</option>`;
    Array.from(uniqueTargets).sort().forEach(u => targetUserSel.innerHTML += `<option value="${u}">${u}</option>`);
    if (uniqueTargets.has(currentSelectedTarget)) targetUserSel.value = currentSelectedTarget; else targetUserSel.value = '全被使用者';

    const currentSelectedProject = projectSel.value || '全專案';
    let uniqueProjects = new Set();
    validExpenseData.forEach(e => { if(e.project) uniqueProjects.add(e.project); });
    projectSel.innerHTML = `<option value="全專案">全專案</option>`;
    Array.from(uniqueProjects).sort().forEach(p => projectSel.innerHTML += `<option value="${p}">${p}</option>`);
    if (uniqueProjects.has(currentSelectedProject)) projectSel.value = currentSelectedProject; else projectSel.value = '全專案';

    let uniqueYears = new Set();
    validExpenseData.forEach(e => {
        if(e.date) { let parts = e.date.split('-'); if(parts.length > 0) uniqueYears.add(parts[0]); }
    });
    let yearsArray = Array.from(uniqueYears).sort((a,b) => b - a); 
    const currentSelectedYear = ys.value;

    if(yearsArray.length === 0) ys.innerHTML = `<option value="">無資料</option>`;
    else {
        ys.innerHTML = yearsArray.map(y => `<option value="${y}">${y}年</option>`).join('');
        if (yearsArray.includes(currentSelectedYear)) ys.value = currentSelectedYear; else ys.value = yearsArray[0];
    }
    updateChartMonthDropdown();
}

function updateChartMonthDropdown() {
    const ys = document.getElementById('chartYearSelect');
    const ms = document.getElementById('chartMonthSelect');
    if(!ys || !ms) return;

    const selectedYear = ys.value;
    let uniqueMonths = new Set();
    validExpenseData.forEach(e => {
        if(e.date && e.date.startsWith(selectedYear + '-')) {
            let parts = e.date.split('-');
            if(parts.length > 1) uniqueMonths.add(parseInt(parts[1], 10));
        }
    });

    let monthsArray = Array.from(uniqueMonths).sort((a,b) => b - a); 
    const currentSelectedMonth = parseInt(ms.value, 10);

    if(monthsArray.length === 0) ms.innerHTML = `<option value="">無資料</option>`;
    else {
        ms.innerHTML = monthsArray.map(m => `<option value="${m}">${m}月</option>`).join('');
        if (monthsArray.includes(currentSelectedMonth)) ms.value = currentSelectedMonth; else ms.value = monthsArray[0]; 
    }
}

function getExceptionReason(exp) {
    let reasons = [];
    if (!appData['使用者'].find(u => u.name === exp.user)) reasons.push("付款人");
    let tUser = exp.targetUser || exp.user;
    if (!appData['使用者'].find(u => u.name === tUser)) reasons.push("被使用者");
    if (!appData['大分類'].find(c => c.name === exp.mainCategory)) reasons.push("大分類");
    if (!appData['小分類'].find(c => c.name === exp.subCategory)) reasons.push("小分類");
    if (reasons.length > 0) return `缺少 ${reasons.join('、')} 資料`;
    return null;
}

function fetchData() {
    const globalSel = document.getElementById('globalLedgerSelect');
    const ledger = globalSel ? globalSel.value : '日常帳本';
    
    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: "getData", ledgerName: ledger }) })
    .then(r => r.json())
    .then(res => {
        appData['帳本'] = (res.manageData || []).filter(d => d.type === '帳本');
        appData['使用者'] = (res.manageData || []).filter(d => d.type === '使用者');
        appData['大分類'] = (res.manageData || []).filter(d => d.type === '大分類');
        appData['小分類'] = (res.manageData || []).filter(d => d.type === '小分類');
        appData['專案'] = (res.manageData || []).filter(d => d.type === '專案'); 
        
        validExpenseData = [];
        exceptionExpenseData = [];
        
        (res.expenseData || []).forEach(e => {
            if (e.date) {
                if (e.date.includes('T')) {
                    let d = new Date(e.date); 
                    if (!isNaN(d.getTime())) e.date = formatDate(d);
                } else {
                    e.date = e.date.substring(0, 10);
                }
            }
            let reason = getExceptionReason(e);
            if (reason) { e.exceptionReason = reason; exceptionExpenseData.push(e); } 
            else { validExpenseData.push(e); }
        });
        
        renderManageLists();
        updateChartDropdowns(); 
        initCompareDropdowns(); 
        
        renderCalendar();
        renderChart();
        renderDailyList(); 
        renderExceptionList();

        if (document.getElementById('page-chart-detail').classList.contains('active')) {
            openChartDetail(currentChartDetailLabel);
        }
        if (document.getElementById('page-compare').classList.contains('active')) {
            renderComparePage();
        }
    }).catch(err => console.log("背景抓取失敗"));
}

window.renderDailyList = function() {
    const list = document.getElementById('expenseList');
    document.getElementById('dailyDetailTitle').innerText = `${selectedDateStr} 明細`;
    
    const dailyData = validExpenseData.filter(e => e.date === selectedDateStr);
    
    if (dailyData.length === 0) {
        list.innerHTML = `<div class="empty-state">本日無記帳紀錄</div>`;
        return;
    }
    
    let dailyIncome = 0;
    let dailyExpense = 0;
    
    let listHTML = dailyData.map(e => {
        const amt = Number(e.amount) || 0;
        if(e.type === '支出') dailyExpense += amt;
        else if(e.type === '收入') dailyIncome += amt; 
        
        const color = e.type === '收入' ? '#4CAF50' : '#ef5350';
        const sign = e.type === '收入' ? '+' : '-';
        const emoji = getIcon(e.mainCategory, '大分類');
        
        const userStyle = getUserColorStyle(e.user); 
        let targetHtml = '';
        let tUser = e.targetUser || e.user;
        if (tUser !== e.user) {
            const targetStyle = getUserColorStyle(tUser);
            targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span><span class="detail-user-tag" style="${targetStyle}">${tUser}</span>`;
        }

        let projectHtml = e.project ? `<span class="detail-project-tag">★ ${e.project}</span>` : '';

        // 【修改】將備註改為獨立換行，黑色字體，無表情符號
        return `
        <div style="padding:12px 0; border-bottom:1px solid #f0f0f0;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div>
                    <div style="display:flex; align-items:center; flex-wrap:wrap; gap:5px;">
                        <span class="cat-emoji">${emoji}</span>
                        <span class="detail-tag">${e.subCategory}</span>
                        ${projectHtml}
                        <span class="detail-user-tag" style="${userStyle}">${e.user}</span>
                        ${targetHtml}
                    </div>
                    ${e.content ? `<div style="font-size:13.5px; color:#333; margin-top:8px; padding-left:40px; font-weight:500;">${e.content}</div>` : ''}
                </div>
                <div style="text-align:right;">
                    <div style="color:${color}; font-weight:bold; font-size:16px; margin-bottom:4px;">${sign}$${amt}</div>
                    <div style="display:flex; gap:8px; justify-content:flex-end;">
                        <button class="icon-btn-gray" style="font-size:13px; background:#f5f5f5; padding:4px 8px; border-radius:4px;" onclick="editExpense('${e.id}')">✎</button>
                        <button class="icon-btn-red" style="font-size:13px; background:#ffebee; padding:4px 8px; border-radius:4px;" onclick="deleteExpense('${e.id}')">✕</button>
                    </div>
                </div>
            </div>
        </div>`;
    }).join('');
    
    let netAmount = dailyIncome - dailyExpense;
    let netColor = netAmount >= 0 ? '#2e7d32' : '#c62828';
    let netBg = netAmount >= 0 ? '#e8f5e9' : '#ffebee';
    let netSign = netAmount >= 0 ? '' : '-';

    const summaryHTML = `
        <div style="background:${netBg}; padding:10px; border-radius:6px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center; font-weight:bold; color:${netColor};">
            <span>本日淨收支</span>
            <span style="font-size:18px;">${netSign}$${Math.abs(netAmount).toLocaleString()}</span>
        </div>`;
    list.innerHTML = summaryHTML + listHTML;
}

window.renderExceptionList = function() {
    const list = document.getElementById('exceptionList');
    if(!list) return;
    
    if (exceptionExpenseData.length === 0) {
        list.innerHTML = `<div class="empty-state">太棒了！目前無任何異常資料</div>`;
        return;
    }
    
    list.innerHTML = exceptionExpenseData.map(e => {
        const userStyle = getUserColorStyle(e.user);
        let tUser = e.targetUser || e.user;
        let targetHtml = '';
        if (tUser !== e.user) {
            const targetStyle = getUserColorStyle(tUser);
            targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span><span class="detail-user-tag" style="${targetStyle}">${tUser}</span>`;
        }

        return `
        <div style="padding:15px 0; border-bottom:1px solid #f0f0f0;">
            <div style="display:flex; justify-content:space-between; margin-bottom:5px;">
                <div style="font-weight:bold; font-size:15px; text-decoration: line-through; color:#999;">${e.subCategory}</div>
                <div style="color:${e.type === '收入' ? '#4CAF50' : '#ef5350'}; font-weight:bold;">$${e.amount}</div>
            </div>
            <div style="font-size:12px; color:#888; margin-bottom:8px; display:flex; align-items:center; gap:6px;">
                <span style="background:#ffebee; color:#c62828; padding:2px 6px; border-radius:4px; font-weight:bold;">⚠️ ${e.exceptionReason}</span>
                <span>${e.date}</span>
                <span class="detail-user-tag" style="${userStyle}">${e.user}</span>
                ${targetHtml}
            </div>
            <div style="display:flex; gap:8px; justify-content:flex-end;">
                <button class="icon-btn-gray" style="font-size:13px; background:#f5f5f5; padding:6px 12px; border-radius:4px; font-weight:bold;" onclick="editExpense('${e.id}')">✎ 重新編輯</button>
                <button class="icon-btn-red" style="font-size:13px; background:#ffebee; padding:6px 12px; border-radius:4px; font-weight:bold;" onclick="deleteExpense('${e.id}')">✕ 永久刪除</button>
            </div>
        </div>`;
    }).join('');
};

window.quickAdd = function(type) {
    let parentName = '';
    if (type === '小分類') {
        parentName = document.getElementById('mainCategory').value;
        if (!parentName) return alert('請先選擇左側的「大項目」！');
    }
    let displayType = type === '大分類' ? '大項目' : (type === '小分類' ? '小項目' : type);
    let newName = prompt(`請輸入新的 ${displayType} 名稱：`);
    if (!newName || newName.trim() === '') return;

    let icon = '📁';
    let color = '';
    if (type === '大分類') icon = prompt(`請為「${newName}」設定一個 Emoji 圖示：\n(例如：🍔、🚗、🎮)`, '🏷️') || '🏷️';
    else if (type === '使用者') color = getRandomColor(); 
    
    const btn = document.getElementById('submitBtn');
    btn.innerText = '新增中...'; btn.disabled = true;
    const payload = { action: 'manage', operation: 'add', type: type, id: Date.now(), name: newName, parentName: parentName, icon: icon, color: color };

    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) })
    .then(res => res.json())
    .then(res => {
        if(res.status === 'success') {
            appData[type].push({id: payload.id, name: payload.name, parentName: payload.parentName, icon: payload.icon, color: payload.color});
            updateFormDropdowns(); 
            if (type === '使用者') { 
                document.getElementById('user').value = payload.name; 
                document.getElementById('targetUser').value = payload.name; 
            }
            if (type === '大分類') { document.getElementById('mainCategory').value = payload.name; updateFormSubCategory(); }
            if (type === '小分類') document.getElementById('subCategory').value = payload.name;
            if (type === '專案') document.getElementById('project').value = payload.name;
            renderManageLists(); 
        }
    })
    .finally(() => { btn.innerText = document.getElementById('expenseOperation').value === 'edit' ? '儲存修改' : '儲存新增'; btn.disabled = false; });
}

window.renderManageLists = function() {
    const globalSel = document.getElementById('globalLedgerSelect');
    if (globalSel) {
        const currentVal = globalSel.value;
        globalSel.innerHTML = '';
        if (appData['帳本'].length === 0) globalSel.innerHTML = `<option value="日常帳本">日常帳本</option>`;
        else {
            appData['帳本'].forEach(ledger => { globalSel.innerHTML += `<option value="${ledger.name}">${ledger.name}</option>`; });
            if (currentVal && appData['帳本'].find(l => l.name === currentVal)) globalSel.value = currentVal;
            else { globalSel.value = appData['帳本'][0].name; fetchData(); }
        }
    }

    renderSimpleList('ledgerList', '帳本', '📓');
    renderSimpleList('projectList', '專案', '★'); 
    
    const userContainer = document.getElementById('userList');
    if(userContainer) {
        userContainer.innerHTML = appData['使用者'].map(item => {
            const c = item.color || '#1976d2';
            return `
            <li class="cat-item">
                <div class="item-header">
                    <div class="item-icon" style="background-color:${c}20; color:${c}; border:1px solid ${c}50; font-size:14px; font-weight:bold;">色</div>
                    <div class="item-content">${item.name}</div>
                    <div class="item-actions">
                        <button class="icon-btn-gray" onclick="openManageModal('edit', '使用者', '${item.id}', '${item.name}', '', '', '${c}')">✎</button>
                        <button class="icon-btn-red" onclick="deleteItem('使用者', '${item.id}')">✕</button>
                    </div>
                </div>
            </li>`;
        }).join('');
    }

    const catContainer = document.getElementById('categoryList');
    if (catContainer) {
        catContainer.innerHTML = '';
        appData['大分類'].forEach(mainCat => {
            const liMain = document.createElement('li');
            liMain.className = 'cat-item';
            const icon = mainCat.icon || '📁';
            liMain.innerHTML = `
                <div class="item-header" style="background:#fafafa; border-radius:8px; padding:10px;">
                    <div class="item-icon" style="background:transparent;"><span class="cat-emoji">${icon}</span></div>
                    <div class="item-content">${mainCat.name}</div>
                    <div class="item-actions">
                        <button class="add-sub-btn" onclick="openManageModal('add', '小分類', null, '', '${mainCat.name}')">+小分類</button>
                        <button class="icon-btn-gray" onclick="openManageModal('edit', '大分類', '${mainCat.id}', '${mainCat.name}', '', '${icon}')">✎</button>
                        <button class="icon-btn-red" onclick="deleteItem('大分類', '${mainCat.id}')">✕</button>
                    </div>
                </div>`;
            catContainer.appendChild(liMain);

            const subs = appData['小分類'].filter(sub => sub.parentName === mainCat.name);
            subs.forEach(sub => {
                const liSub = document.createElement('li');
                liSub.className = 'cat-item';
                liSub.style.borderBottom = "none";
                liSub.innerHTML = `
                    <div class="item-header" style="padding-left: 40px; margin-bottom: 0;">
                        <div class="item-icon" style="background:transparent;"><span class="cat-emoji" style="transform:scale(0.85);">${icon}</span></div>
                        <div class="item-content">${sub.name}</div>
                        <div class="item-actions">
                            <button class="icon-btn-gray" onclick="openManageModal('edit', '小分類', '${sub.id}', '${sub.name}', '${mainCat.name}')">✎</button>
                            <button class="icon-btn-red" onclick="deleteItem('小分類', '${sub.id}')">✕</button>
                        </div>
                    </div>`;
                catContainer.appendChild(liSub);
            });
        });
    }
    updateFormDropdowns(); 
};

function renderSimpleList(containerId, type, icon) {
    const el = document.getElementById(containerId);
    if(el) {
        el.innerHTML = appData[type].map(item => `
            <li class="cat-item">
                <div class="item-header">
                    <div class="item-icon">${icon}</div>
                    <div class="item-content">${item.name}</div>
                    <div class="item-actions">
                        <button class="icon-btn-gray" onclick="openManageModal('edit', '${type}', '${item.id}', '${item.name}')">✎</button>
                        <button class="icon-btn-red" onclick="deleteItem('${type}', '${item.id}')">✕</button>
                    </div>
                </div>
            </li>
        `).join('');
    }
}

function updateFormDropdowns() {
    const userSel = document.getElementById('user');
    const targetUserSel = document.getElementById('targetUser');
    const mainSel = document.getElementById('mainCategory');
    const projectSel = document.getElementById('project');
    
    const curUser = userSel ? userSel.value : '';
    const curTarget = targetUserSel ? targetUserSel.value : '';
    const curMain = mainSel ? mainSel.value : '';
    const curProject = projectSel ? projectSel.value : '';

    if(userSel) {
        userSel.innerHTML = `<option value="" disabled>請選擇</option>` + appData['使用者'].map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if (curUser && appData['使用者'].find(u => u.name === curUser)) userSel.value = curUser;
        else userSel.value = '';
    }

    if(targetUserSel) {
        targetUserSel.innerHTML = `<option value="" disabled>請選擇</option>` + appData['使用者'].map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if (curTarget && appData['使用者'].find(u => u.name === curTarget)) targetUserSel.value = curTarget;
        else targetUserSel.value = userSel.value; 
    }

    if(mainSel) {
        mainSel.innerHTML = `<option value="" disabled>請選擇</option>` + appData['大分類'].map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if(curMain && appData['大分類'].find(c => c.name === curMain)) mainSel.value = curMain;
        else mainSel.value = '';
    }

    if(projectSel) {
        projectSel.innerHTML = `<option value="">無</option>` + appData['專案'].map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if (curProject && appData['專案'].find(p => p.name === curProject)) projectSel.value = curProject;
        else projectSel.value = '';
    }

    updateFormSubCategory();
}

function updateFormSubCategory() {
    const mainSel = document.getElementById('mainCategory');
    const subSel = document.getElementById('subCategory');
    if(!mainSel || !subSel) return;

    const curSub = subSel.value;
    const validSubs = appData['小分類'].filter(sub => sub.parentName === mainSel.value);
    
    if (validSubs.length > 0) {
        subSel.innerHTML = `<option value="" disabled>請選擇</option>` + validSubs.map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if (curSub && validSubs.find(s => s.name === curSub)) subSel.value = curSub;
        else subSel.value = '';
    } else {
        subSel.innerHTML = `<option value="" disabled selected>無小分類</option>`;
    }
}

// 【更新】統計圖表渲染 (加入階層式：大項目與小項目總和)
window.renderChart = function() {
    const container = document.getElementById('customChartBars');
    const inOutSel = document.getElementById('chartInOutSelect');
    const timeSel = document.getElementById('chartTimeSelect');
    const ysEl = document.getElementById('chartYearSelect');
    const msEl = document.getElementById('chartMonthSelect');
    const userSel = document.getElementById('chartUserSelect');
    const targetUserSel = document.getElementById('chartTargetUserSelect');
    const projectSel = document.getElementById('chartProjectSelect');
    
    if(!container || !inOutSel || !timeSel || !ysEl || !msEl || !userSel || !targetUserSel || !projectSel) return;

    const currentChartInOut = inOutSel.value;
    const currentChartTime = timeSel.value;
    const ys = ysEl.value;
    const ms = msEl.value;
    const selectedUser = userSel.value;
    const selectedTargetUser = targetUserSel.value;
    const selectedProject = projectSel.value;
    
    let filtered = validExpenseData;
    
    if (currentChartInOut === '支出') filtered = filtered.filter(e => e.type === '支出');
    else if (currentChartInOut === '收入') filtered = filtered.filter(e => e.type === '收入');

    if (selectedUser && selectedUser !== '全付款人') {
        filtered = filtered.filter(e => String(e.user).trim() === String(selectedUser).trim());
    }
    if (selectedTargetUser && selectedTargetUser !== '全被使用者') {
        filtered = filtered.filter(e => {
            let tUser = e.targetUser || e.user; 
            return String(tUser).trim() === String(selectedTargetUser).trim();
        });
    }
    if (selectedProject && selectedProject !== '全專案') {
        filtered = filtered.filter(e => String(e.project || '').trim() === String(selectedProject).trim());
    }

    filtered = filtered.filter(e => {
        if(!e.date) return false; 
        let parts = e.date.split('-');
        if (parts.length < 2) return false;
        let y = parts[0], m = parseInt(parts[1], 10);
        
        if (currentChartTime === '本月') return y == ys && m == ms;
        else if (currentChartTime === '當年') return y == ys;
        return true; 
    });

    // 建立雙層資料結構：大項目包小項目
    let groupMap = {};
    let totalIncome = 0;
    let totalExpense = 0;

    filtered.forEach(e => {
        let mainKey = e.mainCategory || '未分類';
        let subKey = e.subCategory || '未分類';
        let amt = Number(e.amount) || 0;
        
        if (!groupMap[mainKey]) groupMap[mainKey] = { total: 0, subs: {} };

        if (e.type === '支出') {
            totalExpense += amt;
            if (currentChartInOut === '支出' || currentChartInOut === '總收支') {
                let val = (currentChartInOut === '總收支') ? -amt : amt;
                groupMap[mainKey].total += val;
                groupMap[mainKey].subs[subKey] = (groupMap[mainKey].subs[subKey] || 0) + val;
            }
        } else {
            totalIncome += amt;
            if (currentChartInOut === '收入' || currentChartInOut === '總收支') {
                groupMap[mainKey].total += amt;
                groupMap[mainKey].subs[subKey] = (groupMap[mainKey].subs[subKey] || 0) + amt;
            }
        }
    });

    const summaryBox = document.querySelector('.chart-summary-box');
    const summaryEl = document.getElementById('chartSummaryText');
    
    let prefixArr = [];
    if (selectedProject !== '全專案') prefixArr.push(`【${selectedProject}】`);
    if (selectedUser !== '全付款人') prefixArr.push(`[${selectedUser}]`);
    let prefixStr = prefixArr.length > 0 ? prefixArr.join('') + ' ' : '';

    if (currentChartInOut === '支出') {
        summaryEl.innerText = `${prefixStr}總支出: $${totalExpense.toLocaleString()}`;
        summaryBox.style.background = '#ffebee'; summaryBox.style.borderColor = '#ffcdd2'; summaryEl.style.color = '#ef5350';
    } else if (currentChartInOut === '收入') {
        summaryEl.innerText = `${prefixStr}總收入: $${totalIncome.toLocaleString()}`;
        summaryBox.style.background = '#e8f5e9'; summaryBox.style.borderColor = '#c8e6c9'; summaryEl.style.color = '#4CAF50';
    } else {
        let net = totalIncome - totalExpense;
        let sign = net >= 0 ? '' : '-';
        summaryEl.innerText = `${prefixStr}淨收支: ${sign}$${Math.abs(net).toLocaleString()}`;
        summaryBox.style.background = '#e3f2fd'; summaryBox.style.borderColor = '#bbdefb'; summaryEl.style.color = net >= 0 ? '#4CAF50' : '#ef5350';
    }

    let mainKeys = Object.keys(groupMap).sort((a, b) => Math.abs(groupMap[b].total) - Math.abs(groupMap[a].total));

    if(mainKeys.length === 0) {
        container.innerHTML = '<div style="text-align:center; padding: 30px; color:#999; font-weight:bold;">這個時段沒有資料</div>';
        return;
    }

    const maxVal = Math.max(...mainKeys.map(k => Math.abs(groupMap[k].total)), 1);
    
    // 生成包含小項目的長條圖結構
    container.innerHTML = mainKeys.map(mainName => {
        let mainObj = groupMap[mainName];
        const percent = Math.max((Math.abs(mainObj.total) / maxVal) * 100, 5); 
        
        let barColor = '#ef5350'; 
        if (currentChartInOut === '收入') barColor = '#4CAF50'; 
        else if (currentChartInOut === '總收支') barColor = mainObj.total >= 0 ? '#4CAF50' : '#ef5350';
        
        let displayVal = mainObj.total >= 0 ? `$${mainObj.total.toLocaleString()}` : `-$${Math.abs(mainObj.total).toLocaleString()}`;
        let displayLabel = `<span class="cat-emoji" style="width:22px;height:22px;font-size:13px;">${getIcon(mainName, '大分類')}</span> ${mainName}`;

        // 排序並產生小項目 HTML
        let subKeys = Object.keys(mainObj.subs).sort((a, b) => Math.abs(mainObj.subs[b]) - Math.abs(mainObj.subs[a]));
        let subHtml = subKeys.map(subName => {
            let subVal = mainObj.subs[subName];
            let subDisplay = subVal >= 0 ? `$${subVal.toLocaleString()}` : `-$${Math.abs(subVal).toLocaleString()}`;
            return `
            <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0; border-bottom: 1px dashed #e0e0e0; font-size:13.5px;">
                <span><span style="color:#ccc; font-size:10px; margin-right:6px;">▶</span>${subName}</span>
                <span style="font-weight:bold; color:#555;">${subDisplay}</span>
            </div>`;
        }).join('');

        return `
            <div class="custom-bar-row" style="margin-bottom:24px;">
                <div class="bar-labels" style="align-items:center; cursor:pointer;" onclick="openChartDetail('${mainName}')">
                    <div style="display:flex; align-items:center;">${displayLabel}</div>
                    <span class="bar-value" style="color:${barColor};">${displayVal} <span style="font-size:12px; color:#999;">▶</span></span>
                </div>
                <div class="bar-track" style="margin-bottom:8px;"><div class="bar-fill" style="width: ${percent}%; background:${barColor};"></div></div>
                <div style="padding-left:35px; margin-top:5px;">
                    ${subHtml}
                </div>
            </div>`;
    }).join('');
};

window.openChartDetail = function(label) {
    currentChartDetailLabel = label;
    const currentChartInOut = document.getElementById('chartInOutSelect').value;
    const currentChartTime = document.getElementById('chartTimeSelect').value;
    const ys = document.getElementById('chartYearSelect').value;
    const ms = document.getElementById('chartMonthSelect').value;
    const selectedUser = document.getElementById('chartUserSelect').value;
    const selectedTargetUser = document.getElementById('chartTargetUserSelect').value;
    const selectedProject = document.getElementById('chartProjectSelect').value;
    
    let filtered = validExpenseData;
    
    if (currentChartInOut === '支出') filtered = filtered.filter(e => e.type === '支出');
    else if (currentChartInOut === '收入') filtered = filtered.filter(e => e.type === '收入');

    if (selectedUser && selectedUser !== '全付款人') {
        filtered = filtered.filter(e => String(e.user).trim() === String(selectedUser).trim());
    }
    if (selectedTargetUser && selectedTargetUser !== '全被使用者') {
        filtered = filtered.filter(e => {
            let tUser = e.targetUser || e.user; 
            return String(tUser).trim() === String(selectedTargetUser).trim();
        });
    }
    if (selectedProject && selectedProject !== '全專案') {
        filtered = filtered.filter(e => String(e.project || '').trim() === String(selectedProject).trim());
    }

    filtered = filtered.filter(e => {
        if(!e.date) return false; 
        let parts = e.date.split('-');
        if(parts.length < 2) return false;
        let y = parts[0], m = parseInt(parts[1], 10);
        
        if (currentChartTime === '本月') return y == ys && m == ms;
        else if (currentChartTime === '當年') return y == ys;
        return true; 
    });

    filtered = filtered.filter(e => (e.mainCategory || '未分類') === label);

    let headerIcon = `<span class="cat-emoji">${getIcon(label, '大分類')}</span>`;
    document.getElementById('pageChartDetailTitle').innerHTML = `${headerIcon} ${label} - 明細`;
    
    const listEl = document.getElementById('pageChartDetailList');

    if(filtered.length === 0) {
        listEl.innerHTML = '<div class="empty-state">無明細資料</div>';
    } else {
        listEl.innerHTML = filtered.map(e => {
            const color = e.type === '收入' ? '#4CAF50' : '#ef5350';
            const sign = e.type === '收入' ? '+' : '-';
            const userStyle = getUserColorStyle(e.user);
            
            let targetHtml = '';
            let tUser = e.targetUser || e.user;
            if (tUser !== e.user) {
                const targetStyle = getUserColorStyle(tUser);
                targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span><span class="detail-user-tag" style="${targetStyle}">${tUser}</span>`;
            }
            let projectHtml = e.project ? `<span class="detail-project-tag">★ ${e.project}</span>` : '';

            // 【修改】備註文字無表情符號，並換行顯示
            return `
            <div style="padding:12px 0; border-bottom:1px solid #f0f0f0;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                    <div>
                        <div style="display:flex; align-items:center; flex-wrap:wrap; gap:5px;">
                            <span class="detail-tag">${e.subCategory}</span>
                            ${projectHtml}
                            <span class="detail-user-tag" style="${userStyle}">${e.user}</span>
                            ${targetHtml}
                        </div>
                        ${e.content ? `<div style="font-size:13.5px; color:#333; margin-top:8px; font-weight:500;">${e.content}</div>` : ''}
                    </div>
                    <div style="text-align:right;">
                        <div style="color:${color}; font-weight:bold; font-size:16px; margin-bottom:8px;">${sign}$${e.amount}</div>
                        <div style="display:flex; gap:8px; justify-content:flex-end;">
                            <button class="icon-btn-gray" style="font-size:13px; background:#f5f5f5; padding:4px 8px; border-radius:4px;" onclick="editExpense('${e.id}')">✎</button>
                            <button class="icon-btn-red" style="font-size:13px; background:#ffebee; padding:4px 8px; border-radius:4px;" onclick="deleteExpense('${e.id}')">✕</button>
                        </div>
                    </div>
                </div>
            </div>`;
        }).join('');
    }

    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById('page-chart-detail').classList.add('active');
};

// ==========================================
// 比較頁面 動態圖表與文字報表邏輯
// ==========================================
window.renderComparePage = function() {
    let sYearStr = document.getElementById('compareStartYear').value;
    let eYearStr = document.getElementById('compareEndYear').value;
    const month = parseInt(document.getElementById('compareMonth').value, 10);
    
    const summaryContainer = document.getElementById('compareSummaryArea');
    const chartContainer = document.getElementById('compareChartArea');
    if(!summaryContainer || !chartContainer) return;

    let sYear = parseInt(sYearStr, 10);
    let eYear = parseInt(eYearStr, 10);

    if (isNaN(sYear) || isNaN(eYear)) {
        let opts = Array.from(document.getElementById('compareStartYear').options);
        if(opts.length > 0 && opts[0].value !== "") {
            sYear = parseInt(opts[opts.length > 1 ? opts.length - 2 : 0].value, 10);
            eYear = parseInt(opts[opts.length - 1].value, 10);
        } else {
            chartContainer.innerHTML = '<div style="text-align:center; padding:30px; color:#999;">沒有記帳資料可供比較</div>';
            summaryContainer.innerHTML = '';
            return;
        }
    }

    const startY = Math.min(sYear, eYear);
    const endY = Math.max(sYear, eYear);
    
    let expenses = validExpenseData.filter(e => e.type === '支出');
    
    let datasets = [];
    let targetTotal = 0; 
    let prevTotal = 0;   
    let cardTitle = '';
    
    let yearsInRange = [];
    for(let y = startY; y <= endY; y++) yearsInRange.push(y);

    let xLabels = yearsInRange.map(y => y + '年');
    let dataArray = new Array(yearsInRange.length).fill(0);

    if (compareMode === 'month') {
        cardTitle = `單月歷史趨勢 (${startY}年~${endY}年 ${month}月)`;
        expenses.forEach(e => {
            if(!e.date) return;
            let p = e.date.split('-');
            let y = parseInt(p[0], 10);
            let m = parseInt(p[1], 10);
            if(y >= startY && y <= endY && m === month) {
                let yIndex = y - startY;
                dataArray[yIndex] += Number(e.amount);
            }
        });
        datasets.push({ label: `${month}月總額`, data: dataArray, color: compareColors[1] });
    } else {
        cardTitle = `年度歷史趨勢 (${startY}年~${endY}年)`;
        expenses.forEach(e => {
            if(!e.date) return;
            let p = e.date.split('-');
            let y = parseInt(p[0], 10);
            if(y >= startY && y <= endY) {
                let yIndex = y - startY;
                dataArray[yIndex] += Number(e.amount);
            }
        });
        datasets.push({ label: `年度總額`, data: dataArray, color: compareColors[0] });
    }

    if (dataArray.length > 0) targetTotal = dataArray[dataArray.length - 1];
    if (dataArray.length > 1) prevTotal = dataArray[dataArray.length - 2];
    else prevTotal = 0;

    let maxVal = Math.max(...dataArray);
    if(maxVal === 0) maxVal = 100;

    const width = chartContainer.clientWidth || 300;
    const height = 260;
    const padX = 40;
    const padY = 40; 
    const chartW = width - padX * 2;
    const chartH = height - padY * 2;

    let legendHtml = `<div style="display:flex; justify-content:center; gap:15px; margin-bottom:10px; flex-wrap:wrap;">`;
    datasets.forEach(ds => {
        legendHtml += `<div style="display:flex; align-items:center; gap:5px; font-size:13px; font-weight:bold; color:#555;">
            <div style="width:16px; height:4px; background-color:${ds.color}; border-radius:2px;"></div>
            ${ds.label}
        </div>`;
    });
    legendHtml += `</div>`;

    let svg = `<svg width="100%" height="${height}">`;

    for(let i=0; i<=4; i++) {
        let y = padY + chartH - (i/4)*chartH;
        let val = Math.round((i/4)*maxVal);
        svg += `<text x="${padX-5}" y="${y+4}" font-size="11" text-anchor="end" fill="#999">${val}</text>`;
    }

    let step = Math.ceil(xLabels.length / 10); 
    xLabels.forEach((lbl, i) => {
        let x = padX + (i / (xLabels.length-1)) * chartW;
        if(i % step === 0 || i === xLabels.length-1) {
            svg += `<text x="${x}" y="${height-10}" font-size="11" text-anchor="middle" fill="#999">${lbl}</text>`;
        }
    });

    datasets.forEach(ds => {
        let points = ds.data.map((val, i) => {
            let x = padX + (i / (xLabels.length-1)) * chartW;
            let y = padY + chartH - (val/maxVal) * chartH;
            return `${x},${y}`;
        }).join(' ');
        svg += `<polyline points="${points}" fill="none" stroke="${ds.color}" stroke-width="2.5" />`;
        
        ds.data.forEach((val, i) => {
            let x = padX + (i / (xLabels.length-1)) * chartW;
            let y = padY + chartH - (val/maxVal) * chartH;
            svg += `<circle cx="${x}" cy="${y}" r="4" fill="white" stroke="${ds.color}" stroke-width="2" />`;
            
            if (val >= 0) {
                svg += `<text x="${x}" y="${y - 12}" font-size="13" text-anchor="middle" fill="#444" font-weight="bold">$${val.toLocaleString()}</text>`;
            }
        });
    });
    svg += `</svg>`;

    chartContainer.innerHTML = legendHtml + svg;

    let diff = targetTotal - prevTotal;
    let diffColorClass = diff > 0 ? 'red' : (diff < 0 ? 'green' : 'gray'); 
    let diffSign = diff > 0 ? '+' : '';
    
    let targetYearStr = `${endY}年`;
    let prevYearStr = yearsInRange.length > 1 ? `${endY-1}年` : '無資料';
    let prevValStr = yearsInRange.length > 1 ? `$${prevTotal.toLocaleString()}` : '$0';

    summaryContainer.innerHTML = `
        <div class="compare-card">
            <div class="compare-card-title">${cardTitle}</div>
            <div class="compare-row main">
                <span class="compare-label">${targetYearStr}</span>
                <span class="compare-val-main">$${targetTotal.toLocaleString()}</span>
            </div>
            <div class="compare-row">
                <span class="compare-sub-label">${prevYearStr}: ${prevValStr}</span>
                <span class="compare-diff ${diffColorClass}">差異: ${diffSign}$${diff.toLocaleString()}</span>
            </div>
        </div>
    `;
};

window.editExpense = function(id) {
    let exp = validExpenseData.find(e => e.id.toString() === id.toString()) || exceptionExpenseData.find(e => e.id.toString() === id.toString());
    if(!exp) return;

    document.getElementById('expenseId').value = exp.id;
    document.getElementById('expenseOperation').value = 'edit';
    document.getElementById('date').value = exp.date; 
    document.querySelector('select[name="type"]').value = exp.type === '收入' ? 'income' : 'expense';
    document.querySelector('input[name="amount"]').value = exp.amount;
    document.querySelector('input[name="content"]').value = exp.content || '';

    let userSel = document.getElementById('user');
    if(appData['使用者'].find(u => u.name === exp.user)) userSel.value = exp.user;
    else userSel.value = ""; 

    let targetSel = document.getElementById('targetUser');
    let tUser = exp.targetUser || exp.user; 
    if(appData['使用者'].find(u => u.name === tUser)) targetSel.value = tUser;
    else targetSel.value = ""; 

    let mainSel = document.getElementById('mainCategory');
    if(appData['大分類'].find(c => c.name === exp.mainCategory)) mainSel.value = exp.mainCategory;
    else mainSel.value = "";

    updateFormSubCategory();
    let subSel = document.getElementById('subCategory');
    if(appData['小分類'].find(c => c.name === exp.subCategory)) subSel.value = exp.subCategory;
    else subSel.value = "";

    let projectSel = document.getElementById('project');
    if(appData['專案'].find(p => p.name === exp.project)) projectSel.value = exp.project;
    else projectSel.value = ""; 

    document.getElementById('submitBtn').innerText = '儲存修改';
    document.getElementById('addModal').style.display = 'flex';         
};

window.deleteExpense = function(id) {
    if(!confirm('確定要永久刪除這筆明細嗎？')) return;
    const ledger = document.getElementById('globalLedgerSelect').value;
    
    validExpenseData = validExpenseData.filter(e => e.id.toString() !== id.toString());
    exceptionExpenseData = exceptionExpenseData.filter(e => e.id.toString() !== id.toString());
    
    renderDailyList();
    renderChart();
    renderExceptionList();
    if (document.getElementById('page-chart-detail').classList.contains('active')) openChartDetail(currentChartDetailLabel);
    if (document.getElementById('page-compare').classList.contains('active')) renderComparePage();
    
    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'expense', operation: 'delete', id: id, ledgerName: ledger }) })
        .then(r => r.json()).then(res => { if(res.status === 'success') fetchData(); });
};

// 【修復】強制從 input 讀取明確的 operation 值，杜絕產生兩筆資料的可能
document.getElementById('expenseForm').onsubmit = function(e) {
    e.preventDefault();
    const btn = document.getElementById('submitBtn');
    btn.innerText = '處理中...'; btn.disabled = true;

    const formData = new FormData(this);
    const data = Object.fromEntries(formData.entries());
    
    data.operation = document.getElementById('expenseOperation').value;
    if (!data.targetUser) data.targetUser = data.user;
    data.action = 'expense'; 
    data.ledgerName = document.getElementById('globalLedgerSelect').value;

    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify(data) })
        .then(res => res.json())
        .then(res => {
            if(res.status === 'success') {
                alert(data.operation === 'add' ? '新增成功！' : '修改成功！');
                document.getElementById('addModal').style.display = 'none';
                fetchData(); 
            } else {
                alert('儲存失敗：' + res.message);
            }
        })
        .catch(err => { alert('網路錯誤或伺服器無回應'); })
        .finally(() => { btn.innerText = '儲存'; btn.disabled = false; });
};

window.openManageModal = function(operation, type, id = null, currentName = '', parentName = '', currentIcon = '', currentColor = '') {
    document.getElementById('manageOperation').value = operation;
    document.getElementById('manageType').value = type;
    document.getElementById('manageId').value = id || Date.now(); 
    document.getElementById('manageNameInput').value = currentName;
    document.getElementById('manageModalTitle').innerText = (operation === 'add' ? '新增' : '編輯') + type;
    
    const parentGroup = document.getElementById('manageParentCatGroup');
    const parentSelect = document.getElementById('manageParentCatInput');
    const iconGroup = document.getElementById('manageIconGroup');
    const iconInput = document.getElementById('manageIconInput');
    const colorGroup = document.getElementById('manageColorGroup');
    const colorInput = document.getElementById('manageColorInput');

    if (type === '大分類') {
        iconGroup.style.display = 'block';
        iconInput.value = currentIcon || '📁';
    } else {
        iconGroup.style.display = 'none';
        iconInput.value = '';
    }

    if (type === '使用者') {
        colorGroup.style.display = 'block';
        colorInput.value = currentColor || getRandomColor();
    } else {
        colorGroup.style.display = 'none';
    }

    if (type === '小分類') {
        parentGroup.style.display = 'block';
        parentSelect.innerHTML = appData['大分類'].map(i => `<option value="${i.name}">${i.name}</option>`).join('');
        if (parentName) parentSelect.value = parentName; 
    } else {
        parentGroup.style.display = 'none';
    }
    document.getElementById('manageItemModal').style.display = 'flex';
};

window.deleteItem = function(type, id) {
    if(!confirm(`確定要刪除這個${type}嗎？\n(相關明細將會被移動到「異常」分頁！)`)) return;
    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'manage', operation: 'delete', type: type, id: id }) })
        .then(r => r.json()).then(res => { if(res.status === 'success') fetchData(); });
};

document.getElementById('manageItemForm').onsubmit = function(e) {
    e.preventDefault();
    const btn = document.getElementById('manageSubmitBtn');
    btn.innerText = '儲存中...'; btn.disabled = true;

    const type = document.getElementById('manageType').value;
    const parentName = type === '小分類' ? document.getElementById('manageParentCatInput').value : "";
    const icon = document.getElementById('manageIconInput').value.trim();
    const color = document.getElementById('manageColorInput').value.trim();

    const payload = {
        action: 'manage', operation: document.getElementById('manageOperation').value,
        type: type, id: document.getElementById('manageId').value,
        name: document.getElementById('manageNameInput').value, parentName: parentName, icon: icon, color: color
    };

    fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) })
        .then(res => res.json())
        .then(res => {
            if(res.status === 'success') {
                document.getElementById('manageItemModal').style.display = 'none';
                fetchData(); 
            } else {
                alert('儲存失敗：' + res.message);
            }
        })
        .catch(err => { alert('網路錯誤或伺服器無回應'); })
        .finally(() => { btn.innerText = '儲存'; btn.disabled = false; });
};