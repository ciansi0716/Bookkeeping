import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { 
    getFirestore, collection, getDocs, doc, setDoc, deleteDoc, query, where 
} from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

// ==========================================
// 你的 Firebase 專案設定
// ==========================================
const firebaseConfig = {
  apiKey: "AIzaSyCJG8IR8ibUWF1m8-4C9VInA2FXXWSQSJ8",
  authDomain: "bookkeeping-80cd1.firebaseapp.com",
  projectId: "bookkeeping-80cd1",
  storageBucket: "bookkeeping-80cd1.firebasestorage.app",
  messagingSenderId: "445068237032",
  appId: "1:445068237032:web:21c27ffbebb283efa2115d",
  measurementId: "G-DT8LFSFRXF"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// 全域變數
let appData = { '帳本': [], '使用者': [], '大分類': [], '小分類': [], '專案': [] };
let validExpenseData = []; 
let exceptionExpenseData = []; 

let currentChartInOut = '支出'; 
let currentChartTime = '本月';
let currentCalDate = new Date();
let selectedDateStr = formatDate(new Date());
let currentChartDetailLabel = '';

let compareMode = 'currentYearMonth'; 
let compareSelectedPeriods = [];
const compareColors = ['#F44336', '#2196F3', '#4CAF50', '#FF9800', '#9C27B0'];

let isGroupedView = false;
let currentAdjustUser = '';
let currentSysBalances = { deposit: 0, credit: 0 };

// ==========================================
// CSV 匯出 / 匯入邏輯
// ==========================================
window.exportDataToCSV = function() {
    const allExpenses = [...validExpenseData, ...exceptionExpenseData];
    if(allExpenses.length === 0) {
        alert('目前帳本沒有任何明細可供匯出！');
        return;
    }

    const headers = ['id', 'date', 'type', 'amount', 'payMethod', 'user', 'targetUser', 'mainCategory', 'subCategory', 'project', 'content', 'ledgerName', 'isRepaid'];
    const csvRows = [];
    csvRows.push(headers.join(','));

    allExpenses.forEach(e => {
        const row = headers.map(header => {
            let val = e[header] !== undefined ? String(e[header]) : '';
            if (val.includes(',') || val.includes('"') || val.includes('\n') || val.includes('\r')) {
                val = `"${val.replace(/"/g, '""')}"`;
            }
            return val;
        });
        csvRows.push(row.join(','));
    });

    const csvString = '\uFEFF' + csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    
    const ledgerName = document.getElementById('globalLedgerSelect').value || '記帳資料';
    link.setAttribute('download', `${ledgerName}_${formatDate(new Date())}.csv`);
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

function parseCSV(text) {
    let result = [];
    let row = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        let char = text[i];
        let nextChar = text[i+1];
        
        if (inQuotes) {
            if (char === '"' && nextChar === '"') {
                cur += '"';
                i++;
            } else if (char === '"') {
                inQuotes = false;
            } else {
                cur += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                row.push(cur);
                cur = '';
            } else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
                if (char === '\r') i++;
                row.push(cur);
                result.push(row);
                row = [];
                cur = '';
            } else {
                cur += char;
            }
        }
    }
    if (cur || text[text.length-1] === ',' || text[text.length-1] === '"') {
        row.push(cur);
    }
    if (row.length > 0) result.push(row);
    return result;
}

window.importDataFromCSV = async function(event) {
    const file = event.target.files[0];
    if (!file) return;

    if(!confirm(`確定要匯入「${file.name}」的資料嗎？\n(若有重複的資料 ID 將會被覆蓋更新)`)) {
        event.target.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = async function(e) {
        const text = e.target.result;
        const rows = parseCSV(text);
        
        if (rows.length <= 1) {
            alert('檔案是空的或格式錯誤！');
            event.target.value = '';
            return;
        }

        let headers = rows[0].map(h => h.trim());
        if (headers[0].charCodeAt(0) === 0xFEFF) {
            headers[0] = headers[0].slice(1);
        }

        const requiredHeaders = ['id', 'date', 'type', 'amount', 'payMethod', 'user', 'mainCategory', 'subCategory'];
        const missingHeaders = requiredHeaders.filter(rh => !headers.includes(rh));
        if (missingHeaders.length > 0) {
            alert(`CSV 檔案缺少必要欄位：${missingHeaders.join(', ')}`);
            event.target.value = '';
            return;
        }

        const btn = document.getElementById('importBtn');
        const originalText = btn.innerText;
        btn.innerText = "匯入中...⏳"; btn.disabled = true;

        try {
            const promises = [];
            for (let i = 1; i < rows.length; i++) {
                const cols = rows[i];
                if (cols.length === 1 && cols[0] === '') continue;

                const item = {};
                headers.forEach((h, index) => {
                    item[h] = cols[index] !== undefined ? cols[index] : '';
                });

                if (!item.id || !item.date) continue; 
                
                item.amount = Number(item.amount) || 0;
                if (!item.targetUser) item.targetUser = item.user;
                if (!item.ledgerName) item.ledgerName = document.getElementById('globalLedgerSelect').value;
                if (item.isRepaid === 'true') item.isRepaid = true;
                if (item.isRepaid === 'false') item.isRepaid = false;

                promises.push(setDoc(doc(db, "expenses", String(item.id)), item));
            }

            await Promise.all(promises);
            alert(`🎉 成功匯入 ${promises.length} 筆資料！`);
            window.fetchData(); 
        } catch (err) {
            console.error(err);
            alert("匯入發生錯誤：" + err.message);
        } finally {
            btn.innerText = originalText; btn.disabled = false;
            event.target.value = ''; 
        }
    };
    reader.readAsText(file);
};

// ==========================================
// 快速切換還款狀態
// ==========================================
window.toggleRepayStatus = async function(id, currentStatus) {
    try {
        await setDoc(doc(db, "expenses", String(id)), { isRepaid: !currentStatus }, { merge: true });
        window.fetchData(); 
    } catch(err) {
        alert("狀態更新失敗：" + err.message);
    }
};

// ==========================================
// 自訂專案多選下拉選單邏輯
// ==========================================
window.toggleProjectDropdown = function() {
    const opts = document.getElementById('projectSelectOptions');
    opts.style.display = opts.style.display === 'none' ? 'block' : 'none';
};

window.updateProjectSelectLabel = function() {
    const checked = Array.from(document.querySelectorAll('.project-checkbox:checked')).map(cb => cb.value);
    const label = document.getElementById('projectSelectLabel');
    if (!label) return;
    if (checked.includes('全專案') || checked.length === 0) {
        label.innerText = '全專案';
    } else {
        label.innerText = checked.join(', ');
    }
};

window.getSelectedProjects = function() {
    const checked = Array.from(document.querySelectorAll('.project-checkbox:checked')).map(cb => cb.value);
    if (checked.length === 0) return ['全專案'];
    return checked;
};

window.onProjectCheckboxChange = function(checkbox) {
    const val = checkbox.value;
    const allCheckboxes = document.querySelectorAll('.project-checkbox');
    
    if (val === '全專案' && checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value !== '全專案') cb.checked = false; });
    } else if (checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value === '全專案') cb.checked = false; });
    }
    
    let anyChecked = Array.from(allCheckboxes).some(cb => cb.checked);
    if (!anyChecked) {
        allCheckboxes.forEach(cb => { if(cb.value === '全專案') cb.checked = true; });
    }
    window.updateProjectSelectLabel();
    window.renderChart();
};

// ==========================================
// 統計頁：被使用者多選下拉選單邏輯
// ==========================================
window.toggleChartTargetDropdown = function() {
    const opts = document.getElementById('chartTargetSelectOptions');
    opts.style.display = opts.style.display === 'none' ? 'block' : 'none';
};

window.updateChartTargetSelectLabel = function() {
    const checked = Array.from(document.querySelectorAll('.chart-target-checkbox:checked')).map(cb => cb.value);
    const label = document.getElementById('chartTargetSelectLabel');
    if (!label) return;
    if (checked.includes('全被使用者') || checked.length === 0) {
        label.innerText = '全被使用者';
    } else {
        label.innerText = checked.join(', ');
    }
};

window.getSelectedChartTargets = function() {
    const checked = Array.from(document.querySelectorAll('.chart-target-checkbox:checked')).map(cb => cb.value);
    if (checked.length === 0) return ['全被使用者'];
    return checked;
};

window.onChartTargetCheckboxChange = function(checkbox) {
    const val = checkbox.value;
    const allCheckboxes = document.querySelectorAll('.chart-target-checkbox');
    
    if (val === '全被使用者' && checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value !== '全被使用者') cb.checked = false; });
    } else if (checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value === '全被使用者') cb.checked = false; });
    }
    
    let anyChecked = Array.from(allCheckboxes).some(cb => cb.checked);
    if (!anyChecked) {
        allCheckboxes.forEach(cb => { if(cb.value === '全被使用者') cb.checked = true; });
    }
    window.updateChartTargetSelectLabel();
    window.renderChart();
};

// ==========================================
// 比較頁：區分使用者多選下拉選單邏輯
// ==========================================
window.toggleCompareUserDropdown = function() {
    const opts = document.getElementById('compareUserSelectOptions');
    opts.style.display = opts.style.display === 'none' ? 'block' : 'none';
};

window.updateCompareUserSelectLabel = function() {
    const checked = Array.from(document.querySelectorAll('.compare-user-checkbox:checked')).map(cb => cb.value);
    const label = document.getElementById('compareUserSelectLabel');
    if (!label) return;
    if (checked.includes('總和 (不區分)') || checked.length === 0) {
        label.innerText = '總和 (不區分)';
    } else {
        label.innerText = checked.join(', ');
    }
};

window.getSelectedCompareUsers = function() {
    const checked = Array.from(document.querySelectorAll('.compare-user-checkbox:checked')).map(cb => cb.value);
    if (checked.length === 0) return ['總和 (不區分)'];
    return checked;
};

window.onCompareUserCheckboxChange = function(checkbox) {
    const val = checkbox.value;
    const allCheckboxes = document.querySelectorAll('.compare-user-checkbox');
    
    if (val === '總和 (不區分)' && checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value !== '總和 (不區分)') cb.checked = false; });
    } else if (checkbox.checked) {
        allCheckboxes.forEach(cb => { if(cb.value === '總和 (不區分)') cb.checked = false; });
    }
    
    let anyChecked = Array.from(allCheckboxes).some(cb => cb.checked);
    if (!anyChecked) {
        allCheckboxes.forEach(cb => { if(cb.value === '總和 (不區分)') cb.checked = true; });
    }
    window.updateCompareUserSelectLabel();
    window.renderComparePage();
};

// ==========================================
// 系統啟動與資料讀取
// ==========================================
document.addEventListener('DOMContentLoaded', function() {
    initUI();
    initDateSelects(); 
    window.fetchData(); 
});

window.fetchData = async function() {
    const globalSel = document.getElementById('globalLedgerSelect');
    const ledger = globalSel ? globalSel.value : '日常帳本';
    
    const list = document.getElementById('expenseList');
    if(list) list.innerHTML = `<div class="empty-state">資料與資料庫同步中...</div>`;

    try {
        const manageSnap = await getDocs(collection(db, "manageData"));
        let manageData = manageSnap.docs.map(d => d.data());

        appData['帳本'] = manageData.filter(d => d.type === '帳本');
        appData['使用者'] = manageData.filter(d => d.type === '使用者');
        appData['大分類'] = manageData.filter(d => d.type === '大分類');
        appData['小分類'] = manageData.filter(d => d.type === '小分類');
        appData['專案'] = manageData.filter(d => d.type === '專案'); 
        
        if (!appData['大分類'].find(c => c.name === '信用卡')) appData['大分類'].push({ id: 'cc_main', name: '信用卡', parentName: '', icon: '💳', color: '' });
        if (!appData['小分類'].find(c => c.name === '繳信用卡')) appData['小分類'].push({ id: 'cc_sub', name: '繳信用卡', parentName: '信用卡', icon: '💳', color: '' });
        if (!appData['大分類'].find(c => c.name === '轉帳')) appData['大分類'].push({ id: 'transfer_main', name: '轉帳', parentName: '', icon: '🔄', color: '' });
        if (!appData['小分類'].find(c => c.name === '轉帳')) appData['小分類'].push({ id: 'transfer_sub', name: '轉帳', parentName: '轉帳', icon: '💸', color: '' });
        if (!appData['大分類'].find(c => c.name === '結帳')) appData['大分類'].push({ id: 'chk_main', name: '結帳', parentName: '', icon: '📝', color: '' });
        if (!appData['小分類'].find(c => c.name === '結帳')) appData['小分類'].push({ id: 'chk_sub', name: '結帳', parentName: '結帳', icon: '📝', color: '' });

        const q = query(collection(db, "expenses"), where("ledgerName", "==", ledger));
        const expenseSnap = await getDocs(q);
        let rawExpenses = expenseSnap.docs.map(d => d.data());

        validExpenseData = [];
        exceptionExpenseData = [];
        
        rawExpenses.forEach(e => {
            if (e.date) {
                let dateStr = String(e.date);
                if (dateStr.includes('T')) {
                    let d = new Date(dateStr); 
                    if (!isNaN(d.getTime())) e.date = formatDate(d);
                } else {
                    e.date = dateStr.substring(0, 10);
                }
            }
            e.payMethod = String(e.payMethod || '存款');

            let reason = getExceptionReason(e);
            if (reason) { e.exceptionReason = reason; exceptionExpenseData.push(e); } 
            else { validExpenseData.push(e); }
        });
        
        window.renderManageLists();
        updateChartDropdowns(); 
        initCompareDropdowns(); 
        
        renderCalendar();
        window.renderChart();
        window.renderDailyList(); 
        window.renderExceptionList();

        if (document.getElementById('page-chart-detail').classList.contains('active')) window.openChartDetail(currentChartDetailLabel);
        if (document.getElementById('page-compare').classList.contains('active')) window.renderComparePage();

    } catch(err) {
        console.error("Fetch Data Error:", err);
        if(list) {
            list.innerHTML = `
            <div class="empty-state" style="color:#e53935; padding: 20px;">
                <div style="font-size:30px; margin-bottom:10px;">⚠️</div>
                <div style="font-weight:bold; font-size:16px;">無法載入資料</div>
                <div style="font-size:13px; color:#555; margin-top:8px;">${err.message}</div>
            </div>`;
        }
    }
};

// ==========================================
// 工具函數
// ==========================================
function formatDate(d) {
    let month = '' + (d.getMonth() + 1), day = '' + d.getDate(), year = d.getFullYear();
    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;
    return [year, month, day].join('-');
}

window.formatSlashDate = function(dateStr) {
    if (!dateStr) return '';
    let parts = dateStr.split('-');
    if (parts.length === 3) {
        return `${parts[0]}/${parseInt(parts[1], 10)}/${parseInt(parts[2], 10)}`;
    }
    return dateStr;
};

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

function isCheckoutRecord(e) {
    return String(e.mainCategory || '').includes('結帳') || String(e.subCategory || '').includes('結帳');
}

function getExceptionReason(exp) {
    let reasons = [];
    if (!appData['使用者'].find(u => u.name === exp.user)) reasons.push("付款人");
    
    let tUsers = (exp.targetUser || exp.user).split(',').map(x=>x.trim()).filter(x=>x);
    let invalidTarget = tUsers.find(tu => !appData['使用者'].find(u => u.name === tu));
    if (invalidTarget) reasons.push("被使用者");
    
    if (String(exp.mainCategory || '').includes('校正') && String(exp.subCategory || '') === '未選擇小項目') return "請為手動校正差額選擇正確的分類";

    if (!appData['大分類'].find(c => c.name === exp.mainCategory)) reasons.push("大分類");
    if (!appData['小分類'].find(c => c.name === exp.subCategory)) reasons.push("小分類");
    if (reasons.length > 0) return `缺少 ${reasons.join('、')} 資料`;
    return null;
}

// ==========================================
// 初始化 UI 與事件綁定
// ==========================================
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
            
            if(target === 'page-charts') window.renderChart();
            if(target === 'page-manage') {
                document.querySelectorAll('.manage-tabs .tab-item').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.manage-content').forEach(c => c.classList.remove('active'));
                document.querySelector('.manage-tabs .tab-item[data-target="manage-content-users"]').classList.add('active');
                document.getElementById('manage-content-users').classList.add('active');
                window.renderManageLists();
            }
            if(target === 'page-exceptions') window.renderExceptionList();
            if(target === 'page-compare') {
                initCompareDropdowns();
                window.renderComparePage();
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
    const chartShowCheckout = document.getElementById('chartShowCheckout');
    const chartShowTransfer = document.getElementById('chartShowTransfer');
    const chartShowUnrepaid = document.getElementById('chartShowUnrepaid');
    
    const chartStartDate = document.getElementById('chartStartDate');
    const chartEndDate = document.getElementById('chartEndDate');
    if(chartStartDate) { chartStartDate.value = formatDate(new Date()); chartStartDate.addEventListener('change', window.renderChart); }
    if(chartEndDate) { chartEndDate.value = formatDate(new Date()); chartEndDate.addEventListener('change', window.renderChart); }

    if(chartShowCheckout) {
        chartShowCheckout.checked = localStorage.getItem('chartShowCheckout') === 'true';
        chartShowCheckout.addEventListener('change', function() {
            localStorage.setItem('chartShowCheckout', this.checked);
            window.renderChart();
        });
    }
    
    if(chartShowTransfer) {
        chartShowTransfer.checked = localStorage.getItem('chartShowTransfer') === 'true';
        chartShowTransfer.addEventListener('change', function() {
            localStorage.setItem('chartShowTransfer', this.checked);
            window.renderChart();
        });
    }
    
    if(chartShowUnrepaid) {
        chartShowUnrepaid.checked = localStorage.getItem('chartShowUnrepaid') === 'true';
        chartShowUnrepaid.addEventListener('change', function() {
            localStorage.setItem('chartShowUnrepaid', this.checked);
            window.renderChart();
        });
    }

    [inOutSel, userSel].forEach(el => {
        if(el) el.addEventListener('change', window.renderChart);
    });

    if(timeSel) {
        timeSel.addEventListener('change', () => {
            currentChartTime = timeSel.value;
            const customRange = document.getElementById('chartCustomDateRange');
            yearSel.style.display = 'none'; 
            monthSel.style.display = 'none';
            if(customRange) customRange.style.display = 'none';

            if (currentChartTime === '歷年') {
            } else if (currentChartTime === '當年') {
                yearSel.style.display = 'inline-block';
            } else if (currentChartTime === '本月') { 
                yearSel.style.display = 'inline-block'; 
                monthSel.style.display = 'inline-block';
            } else if (currentChartTime === '自訂區間') {
                if(customRange) customRange.style.display = 'flex';
            }
            window.renderChart();
        });
    }
    
    if(yearSel) {
        yearSel.addEventListener('change', () => { updateChartMonthDropdown(); window.renderChart(); });
    }
    if(monthSel) {
        monthSel.addEventListener('change', window.renderChart);
    }

    const compModeSel = document.getElementById('compareModeSelect');
    const compBaseYear = document.getElementById('compareBaseYear');
    const compEndYear = document.getElementById('compareEndYear');
    const compRangeLabel = document.getElementById('compareRangeLabel');
    const compMonth = document.getElementById('compareMonth');
    const compareShowCheckout = document.getElementById('compareShowCheckout');
    const compareShowTransfer = document.getElementById('compareShowTransfer');
    const compareShowUnrepaid = document.getElementById('compareShowUnrepaid');

    if(compareShowCheckout) {
        compareShowCheckout.checked = localStorage.getItem('compareShowCheckout') === 'true';
        compareShowCheckout.addEventListener('change', function() { localStorage.setItem('compareShowCheckout', this.checked); window.renderComparePage(); });
    }
    if(compareShowTransfer) {
        compareShowTransfer.checked = localStorage.getItem('compareShowTransfer') === 'true';
        compareShowTransfer.addEventListener('change', function() { localStorage.setItem('compareShowTransfer', this.checked); window.renderComparePage(); });
    }
    if(compareShowUnrepaid) {
        compareShowUnrepaid.checked = localStorage.getItem('compareShowUnrepaid') === 'true';
        compareShowUnrepaid.addEventListener('change', function() { localStorage.setItem('compareShowUnrepaid', this.checked); window.renderComparePage(); });
    }

    if(compModeSel) {
        compModeSel.addEventListener('change', () => {
            compareMode = compModeSel.value;
            if(compRangeLabel) compRangeLabel.style.display = (compareMode === 'year') ? 'inline-block' : 'none';
            if(compEndYear) compEndYear.style.display = (compareMode === 'year') ? 'inline-block' : 'none';
            if(compMonth) compMonth.style.display = (compareMode === 'month') ? 'inline-block' : 'none';
            window.renderComparePage();
        });
    }
    [compBaseYear, compEndYear, compMonth].forEach(el => {
        if(el) el.addEventListener('change', window.renderComparePage);
    });

    const addModalUserSel = document.getElementById('user');
    if (addModalUserSel) {
        addModalUserSel.addEventListener('change', function() {
            document.getElementById('targetUser').value = this.value;
        });
    }

    document.getElementById('fabAddBtn').onclick = () => {
        document.getElementById('expenseForm').reset();
        document.getElementById('expenseId').value = Date.now();
        document.getElementById('expenseOperation').value = 'add';
        document.getElementById('submitBtn').innerText = '儲存新增';
        document.getElementById('date').value = selectedDateStr; 
        document.getElementById('payMethod').value = '存款'; 
        document.getElementById('isRepaidCheckbox').checked = false;
        window.updateFormDropdowns();
        document.getElementById('addModal').style.display = 'flex';
    };
    
    document.getElementById('closeModalBtn').onclick = () => document.getElementById('addModal').style.display = 'none';
    
    window.onclick = function(e) { 
        if (e.target.className === 'modal') e.target.style.display = 'none'; 

        const customProjectSelect = document.getElementById('customProjectSelect');
        if (customProjectSelect && !customProjectSelect.contains(e.target)) {
            const opts = document.getElementById('projectSelectOptions');
            if (opts) opts.style.display = 'none';
        }
        
        const customChartTargetSelect = document.getElementById('customChartTargetSelect');
        if (customChartTargetSelect && !customChartTargetSelect.contains(e.target)) {
            const optsTarget = document.getElementById('chartTargetSelectOptions');
            if (optsTarget) optsTarget.style.display = 'none';
        }

        const customCompareUserSelect = document.getElementById('customCompareUserSelect');
        if (customCompareUserSelect && !customCompareUserSelect.contains(e.target)) {
            const optsComp = document.getElementById('compareUserSelectOptions');
            if (optsComp) optsComp.style.display = 'none';
        }
    };

    const globalSel = document.getElementById('globalLedgerSelect');
    if(globalSel) globalSel.addEventListener('change', window.fetchData);

    const mainCatSel = document.getElementById('mainCategory');
    if(mainCatSel) mainCatSel.addEventListener('change', () => window.updateFormSubCategory());
    
    const adjInput = document.getElementById('adjActualBalance');
    if(adjInput) {
        adjInput.addEventListener('input', function() {
            let actual = Number(this.value) || 0;
            let type = document.getElementById('adjAccountType').value;
            let sysBal = type === '存款' ? currentSysBalances.deposit : currentSysBalances.credit;
            
            let diff = actual - sysBal;
            let display = document.getElementById('adjDiffDisplay');
            if (diff > 0) {
                display.innerHTML = `差額: <span style="color:#4CAF50;">+$${diff.toLocaleString()} (系統少算，將新增一筆異常收入)</span>`;
            } else if (diff < 0) {
                display.innerHTML = `差額: <span style="color:#ef5350;">-$${Math.abs(diff).toLocaleString()} (系統多算，將新增一筆異常支出)</span>`;
            } else {
                display.innerHTML = `差額: $0 (無需調整)`;
            }
        });
    }
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
    const baseYear = document.getElementById('compareBaseYear');
    const endYear = document.getElementById('compareEndYear');
    if(!baseYear || !endYear) return;

    let uniqueYears = Array.from(new Set(validExpenseData.map(e => {
        if(!e.date) return null;
        return String(e.date).split('-')[0];
    }).filter(Boolean))).sort();
    
    if(uniqueYears.length === 0) uniqueYears = [new Date().getFullYear().toString()];

    let yHtml = uniqueYears.map(y => `<option value="${y}">${y}年</option>`).join('');
    baseYear.innerHTML = yHtml;
    endYear.innerHTML = yHtml;

    baseYear.value = uniqueYears[uniqueYears.length - 1];
    
    if (uniqueYears.length > 1) {
        endYear.value = uniqueYears[uniqueYears.length - 2];
    } else {
        endYear.value = uniqueYears[uniqueYears.length - 1];
    }
}

function updateChartDropdowns() {
    const ys = document.getElementById('chartYearSelect');
    const ms = document.getElementById('chartMonthSelect');
    const userSel = document.getElementById('chartUserSelect');
    
    if(!ys || !ms || !userSel) return;

    const currentSelectedUser = userSel.value || '全付款人';
    let uniqueUsers = new Set();
    validExpenseData.forEach(e => { if(e.user) uniqueUsers.add(e.user); });
    userSel.innerHTML = `<option value="全付款人">全付款人</option>`;
    Array.from(uniqueUsers).sort().forEach(u => userSel.innerHTML += `<option value="${u}">${u}</option>`);
    if (uniqueUsers.has(currentSelectedUser)) userSel.value = currentSelectedUser; else userSel.value = '全付款人';

    let currentSelectedTargets = [];
    if (document.querySelectorAll('.chart-target-checkbox').length > 0) {
        currentSelectedTargets = window.getSelectedChartTargets();
    } else {
        currentSelectedTargets = ['全被使用者'];
    }

    let uniqueTargets = new Set();
    validExpenseData.forEach(e => { 
        let t = e.targetUser || e.user; 
        t.split(',').forEach(tu => uniqueTargets.add(tu.trim()));
    });
    
    let targetOptionsHtml = '';
    let isAllTargetChecked = currentSelectedTargets.includes('全被使用者') ? 'checked' : '';
    targetOptionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
        <input type="checkbox" class="chart-target-checkbox" value="全被使用者" onchange="onChartTargetCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isAllTargetChecked}> <span style="flex:1;">全被使用者</span>
    </label>`;
    
    Array.from(uniqueTargets).sort().forEach(u => {
        let isChecked = currentSelectedTargets.includes(u) ? 'checked' : '';
        targetOptionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
            <input type="checkbox" class="chart-target-checkbox" value="${u}" onchange="onChartTargetCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isChecked}> <span style="flex:1;">${u}</span>
        </label>`;
    });

    const targetOptsContainer = document.getElementById('chartTargetSelectOptions');
    if (targetOptsContainer) {
        targetOptsContainer.innerHTML = targetOptionsHtml;
        window.updateChartTargetSelectLabel();
    }

    let currentSelectedCompareUsers = [];
    if (document.querySelectorAll('.compare-user-checkbox').length > 0) {
        currentSelectedCompareUsers = window.getSelectedCompareUsers();
    } else {
        currentSelectedCompareUsers = ['總和 (不區分)'];
    }

    let compareOptionsHtml = '';
    let isAllCompChecked = currentSelectedCompareUsers.includes('總和 (不區分)') ? 'checked' : '';
    compareOptionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
        <input type="checkbox" class="compare-user-checkbox" value="總和 (不區分)" onchange="onCompareUserCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isAllCompChecked}> <span style="flex:1;">總和 (不區分)</span>
    </label>`;
    
    Array.from(uniqueTargets).sort().forEach(u => {
        let isChecked = currentSelectedCompareUsers.includes(u) ? 'checked' : '';
        compareOptionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
            <input type="checkbox" class="compare-user-checkbox" value="${u}" onchange="onCompareUserCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isChecked}> <span style="flex:1;">${u}</span>
        </label>`;
    });

    const compareOptsContainer = document.getElementById('compareUserSelectOptions');
    if (compareOptsContainer) {
        compareOptsContainer.innerHTML = compareOptionsHtml;
        window.updateCompareUserSelectLabel();
    }

    let currentSelectedProjects = [];
    if (document.querySelectorAll('.project-checkbox').length > 0) {
        currentSelectedProjects = window.getSelectedProjects();
    } else {
        currentSelectedProjects = ['全專案'];
    }

    let uniqueProjects = new Set();
    validExpenseData.forEach(e => { if(e.project) uniqueProjects.add(e.project); });
    
    let optionsHtml = '';
    let isAllChecked = currentSelectedProjects.includes('全專案') ? 'checked' : '';
    optionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
        <input type="checkbox" class="project-checkbox" value="全專案" onchange="onProjectCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isAllChecked}> <span style="flex:1;">全專案</span>
    </label>`;
    
    let isNoneChecked = currentSelectedProjects.includes('無專案') ? 'checked' : '';
    optionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
        <input type="checkbox" class="project-checkbox" value="無專案" onchange="onProjectCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isNoneChecked}> <span style="flex:1;">無專案</span>
    </label>`;
    
    Array.from(uniqueProjects).sort().forEach(p => {
        let isChecked = currentSelectedProjects.includes(p) ? 'checked' : '';
        optionsHtml += `<label style="display:flex; align-items:center; padding: 8px 10px; cursor:pointer; border-bottom: 1px solid #eee; font-size: 14px;">
            <input type="checkbox" class="project-checkbox" value="${p}" onchange="onProjectCheckboxChange(this)" style="width:16px; height:16px; margin-right:8px; cursor:pointer;" ${isChecked}> <span style="flex:1;">${p}</span>
        </label>`;
    });

    const optsContainer = document.getElementById('projectSelectOptions');
    if (optsContainer) {
        optsContainer.innerHTML = optionsHtml;
        window.updateProjectSelectLabel();
    }

    let uniqueYears = new Set();
    validExpenseData.forEach(e => {
        if(e.date) { let parts = String(e.date).split('-'); if(parts.length > 0) uniqueYears.add(parts[0]); }
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
        if(e.date && String(e.date).startsWith(selectedYear + '-')) {
            let parts = String(e.date).split('-');
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

// ==========================================
// 畫面渲染函數
// ==========================================
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
            window.renderDailyList(); 
        };
        grid.appendChild(el);
    }
}

window.toggleGroupView = function() {
    isGroupedView = !isGroupedView;
    const btn = document.getElementById('toggleGroupBtn');
    if (isGroupedView) {
        btn.innerText = '取消合併';
        btn.style.background = '#ffe0b2';
        btn.style.color = '#e65100';
        btn.style.borderColor = '#ffcc80';
    } else {
        btn.innerText = '合併同項';
        btn.style.background = '#e3f2fd';
        btn.style.color = '#1565c0';
        btn.style.borderColor = '#bbdefb';
    }
    window.renderDailyList();
};

window.toggleUserTotal = function(username, isChecked) {
    let activeUsers = JSON.parse(localStorage.getItem('activeUsersForTotal')) || [];
    if (isChecked && !activeUsers.includes(username)) {
        activeUsers.push(username);
    } else if (!isChecked) {
        activeUsers = activeUsers.filter(u => u !== username);
    }
    localStorage.setItem('activeUsersForTotal', JSON.stringify(activeUsers));
    window.renderManageLists();
};

window.renderManageLists = function() {
    const globalSel = document.getElementById('globalLedgerSelect');
    if (globalSel) {
        const currentVal = globalSel.value;
        globalSel.innerHTML = '';
        if (appData['帳本'].length === 0) globalSel.innerHTML = `<option value="日常帳本">日常帳本</option>`;
        else {
            appData['帳本'].forEach(ledger => { globalSel.innerHTML += `<option value="${ledger.name}">${ledger.name}</option>`; });
            if (currentVal && appData['帳本'].find(l => l.name === currentVal)) globalSel.value = currentVal;
            else { globalSel.value = appData['帳本'][0].name; window.fetchData(); }
        }
    }

    renderSimpleList('ledgerList', '帳本', '📓');
    renderSimpleList('projectList', '專案', '★'); 
    
    const userContainer = document.getElementById('userList');
    if(userContainer) {
        let totalAllDeposit = 0;
        let totalAllCredit = 0; 
        
        let activeUsers = JSON.parse(localStorage.getItem('activeUsersForTotal'));
        if (!Array.isArray(activeUsers)) {
            activeUsers = appData['使用者'].map(u => u.name); 
            localStorage.setItem('activeUsersForTotal', JSON.stringify(activeUsers));
        }
        
        let userHtml = appData['使用者'].map(item => {
            const c = item.color || '#1976d2';
            const balances = window.getUserBalance(item.name); 
            const isChecked = activeUsers.includes(item.name);
            
            if (isChecked) {
                totalAllDeposit += balances.deposit; 
                totalAllCredit += balances.credit;
            }
            
            return `
            <li class="cat-item">
                <div class="item-header">
                    <input type="checkbox" title="計入系統總額" style="margin-right:12px; width:18px; height:18px; cursor:pointer;" ${isChecked ? 'checked' : ''} onchange="toggleUserTotal('${item.name}', this.checked)">
                    <div class="item-icon" style="background-color:${c}20; color:${c}; border:1px solid ${c}50; font-size:14px; font-weight:bold; margin-right:10px;">色</div>
                    <div class="item-content">
                        <div style="font-size:16px;">${item.name}</div>
                        <div style="font-size:13px; color:#888; margin-top:3px; display:flex; gap:10px;">
                            <span>存款: <span style="color:${balances.deposit>=0?'#4CAF50':'#ef5350'}; font-weight:bold;">$${balances.deposit.toLocaleString()}</span></span>
                            <span>信用卡: <span style="color:${balances.credit>=0?'#1976d2':'#ef5350'}; font-weight:bold;">$${balances.credit.toLocaleString()}</span></span>
                        </div>
                    </div>
                    <div class="item-actions">
                        <button class="icon-btn-gray" style="font-size:13px; background:#fff3e0; color:#e65100; padding:4px 8px; border-radius:4px; font-weight:bold;" onclick="openAdjustBalanceModal('${item.name}')">校正</button>
                        <button class="icon-btn-gray" onclick="openManageModal('edit', '使用者', '${item.id}', '${item.name}', '', '', '${c}')">✎</button>
                        <button class="icon-btn-red" onclick="deleteItem('使用者', '${item.id}')">✕</button>
                    </div>
                </div>
            </li>`;
        }).join('');
        
        let totalHtml = `
        <div style="background:#e3f2fd; border:1px solid #bbdefb; padding:12px; border-radius:8px; margin-bottom:15px; display:flex; justify-content:space-around; align-items:center;">
            <div style="text-align:center;">
                <div style="font-weight:bold; color:#1565c0; font-size:13px;">系統總存款</div>
                <div style="font-size:18px; font-weight:bold; color:#1976d2;">$${totalAllDeposit.toLocaleString()}</div>
            </div>
            <div style="width:1px; background:#bbdefb; height:30px;"></div>
            <div style="text-align:center;">
                <div style="font-weight:bold; color:#1565c0; font-size:13px;">信用卡總可用額度</div>
                <div style="font-size:18px; font-weight:bold; color:#1976d2;">$${totalAllCredit.toLocaleString()}</div>
            </div>
        </div>`;
        
        userContainer.innerHTML = totalHtml + userHtml;
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
    window.updateFormDropdowns(); 
};

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

    let displayData = dailyData;
    if (isGroupedView) {
        let grouped = {};
        dailyData.forEach(e => {
            let key = `${e.type}_${e.mainCategory}_${e.subCategory}_${e.user}_${e.targetUser}_${e.project}_${e.payMethod}_${e.isRepaid}`;
            if (!grouped[key]) {
                grouped[key] = { ...e, amount: 0, count: 0, contents: [] };
            }
            grouped[key].amount += Number(e.amount);
            grouped[key].count += 1;
            if (e.content && e.content.trim() !== '') grouped[key].contents.push(e.content);
        });
        
        displayData = Object.values(grouped).map(g => {
            let combinedContent = g.contents.join(' / ');
            return {
                ...g,
                content: g.count > 1 
                    ? `<span style="color:#e65100; font-weight:bold;">(合併 ${g.count} 筆)</span> ${combinedContent}`
                    : g.content,
                isGrouped: g.count > 1
            };
        });
    }
    
    let listHTML = displayData.map(e => {
        const amt = Number(e.amount) || 0;
        
        let isPayingCreditCardBill = (e.mainCategory === '信用卡' && e.subCategory === '繳信用卡' && e.type === '支出');
        let isTransfer = (e.mainCategory === '轉帳'); 
        let isCheckout = isCheckoutRecord(e); 
        
        if (!isPayingCreditCardBill && !isTransfer && !isCheckout) {
            if(e.type === '支出') dailyExpense += amt;
            else if(e.type === '收入') dailyIncome += amt; 
        }
        
        let color = e.type === '收入' ? '#4CAF50' : '#ef5350';
        let sign = e.type === '收入' ? '+' : '-';
        if (isTransfer || isCheckout) {
            color = '#8e24aa'; 
            sign = '';
        }

        const emoji = getIcon(e.mainCategory, '大分類');
        const userStyle = getUserColorStyle(e.user); 
        
        let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
        let isShared = tUsers.length > 1 || (tUsers.length === 1 && tUsers[0] !== e.user);
        
        let targetHtml = '';
        let nonSelfTargets = tUsers.filter(tu => tu !== e.user);
        if (nonSelfTargets.length > 0) {
            targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span>` +
                         nonSelfTargets.map(tu => `<span class="detail-user-tag" style="${getUserColorStyle(tu)}">${tu}</span>`).join('<span style="margin:0 2px;"></span>');
        }

        let projectHtml = e.project ? `<span class="detail-project-tag">★ ${e.project}</span>` : '';
        
        let payBadge = '';
        if (isTransfer) {
            payBadge = `<span class="detail-project-tag" style="background:#f3e5f5; color:#7b1fa2; border-color:#e1bee7;">🔄 內部轉帳</span>`;
        } else if (isCheckout) {
            payBadge = `<span class="detail-project-tag" style="background:#e0f7fa; color:#006064; border-color:#b2ebf2;">📝 系統結帳</span>`;
        } else {
            payBadge = e.payMethod === '信用卡' ? `<span class="detail-project-tag" style="background:#e3f2fd; color:#1565c0; border-color:#bbdefb;">💳 信用卡</span>` : `<span class="detail-project-tag" style="background:#e8f5e9; color:#2e7d32; border-color:#c8e6c9;">💵 存款</span>`;
        }

        // 🌟 勾選框還款按鈕處理 (無框無底色，原生 checkbox 搭配有顏色的文字)
        let repayBadge = '';
        if (isShared && !isTransfer && !isCheckout) {
            let isRepaid = e.isRepaid === true || e.isRepaid === 'true';
            let statusText = isRepaid ? '已還' : '未還';
            let statusColor = isRepaid ? '#4CAF50' : '#ef5350';
            
            if (e.isGrouped) {
                repayBadge = `<span style="color:#999; font-size:13.5px; font-weight:bold; margin-left:6px;">[${statusText}]</span>`;
            } else {
                repayBadge = `
                <label style="display:inline-flex; align-items:center; gap:4px; cursor:pointer; margin:0 0 0 6px;">
                    <input type="checkbox" style="width:14px; height:14px; margin:0; cursor:pointer;" ${isRepaid ? 'checked' : ''} onchange="toggleRepayStatus('${e.id}', ${isRepaid})">
                    <span style="color:${statusColor}; font-weight:bold; font-size:13.5px;">${statusText}</span>
                </label>`;
            }
        }

        // 🌟 單純白底黑字備註，完全靠左 (去除 padding-left: 40px;)
        let contentHtml = '';
        if (e.content || repayBadge) {
            let textPart = e.content ? `<span style="font-size:13.5px; color:#333;">${e.content}</span>` : '';
            let divider = (e.content && repayBadge) ? `<span style="color:#ccc; font-size:13.5px; margin:0 4px;">|</span>` : '';
            
            contentHtml = `<div style="margin-top:6px;">
                <div style="display:inline-flex; align-items:center; flex-wrap:wrap; max-width:100%;">
                    ${textPart}
                    ${divider}
                    ${repayBadge}
                </div>
            </div>`;
        }

        let actionsHtml = '';
        if (e.isGrouped) {
            actionsHtml = `<span style="font-size:12px; color:#999; margin-right:5px; font-weight:bold;">取消合併以編輯</span>`;
        } else {
            actionsHtml = `
                <button class="icon-btn-gray" style="font-size:13px; background:#f5f5f5; padding:4px 8px; border-radius:4px;" onclick="editExpense('${e.id}')">✎</button>
                <button class="icon-btn-red" style="font-size:13px; background:#ffebee; padding:4px 8px; border-radius:4px;" onclick="deleteExpense('${e.id}')">✕</button>
            `;
        }

        return `
        <div style="padding:12px 0; border-bottom:1px solid #f0f0f0;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div style="flex:1;">
                    <div style="display:flex; align-items:center; flex-wrap:wrap; gap:5px;">
                        <span class="cat-emoji">${emoji}</span>
                        <span class="detail-tag">${e.subCategory}</span>
                        ${payBadge}
                        ${projectHtml}
                        <span class="detail-user-tag" style="${userStyle}">${e.user}</span>
                        ${targetHtml}
                    </div>
                    ${contentHtml}
                </div>
                <div style="text-align:right; min-width:70px;">
                    <div style="color:${color}; font-weight:bold; font-size:16px; margin-bottom:4px;">${sign}$${amt.toLocaleString()}</div>
                    <div style="display:flex; gap:8px; justify-content:flex-end; align-items:center;">
                        ${actionsHtml}
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
            <span>本日淨收支 (不含轉帳/結帳/繳卡)</span>
            <span style="font-size:18px;">${netSign}$${Math.abs(netAmount).toLocaleString()}</span>
        </div>`;
    list.innerHTML = summaryHTML + listHTML;
};

window.renderExceptionList = function() {
    const list = document.getElementById('exceptionList');
    if(!list) return;
    
    if (exceptionExpenseData.length === 0) {
        list.innerHTML = `<div class="empty-state">太棒了！目前無任何異常資料</div>`;
        return;
    }
    
    list.innerHTML = exceptionExpenseData.map(e => {
        const userStyle = getUserColorStyle(e.user);
        
        let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
        let targetHtml = '';
        let nonSelfTargets = tUsers.filter(tu => tu !== e.user);
        if (nonSelfTargets.length > 0) {
            targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span>` +
                         nonSelfTargets.map(tu => `<span class="detail-user-tag" style="${getUserColorStyle(tu)}">${tu}</span>`).join('<span style="margin:0 2px;"></span>');
        }

        return `
        <div style="padding:15px 0; border-bottom:1px solid #f0f0f0;">
            <div style="display:flex; justify-content:space-between; margin-bottom:5px;">
                <div style="font-weight:bold; font-size:15px; text-decoration: line-through; color:#999;">${e.subCategory}</div>
                <div style="color:${e.type === '收入' ? '#4CAF50' : '#ef5350'}; font-weight:bold;">$${Number(e.amount).toLocaleString()}</div>
            </div>
            <div style="font-size:12px; color:#888; margin-bottom:8px; display:flex; align-items:center; gap:6px;">
                <span style="background:#ffebee; color:#c62828; padding:2px 6px; border-radius:4px; font-weight:bold;">⚠️ ${e.exceptionReason}</span>
                <span>${window.formatSlashDate(e.date)}</span>
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

window.renderChart = function() {
    const container = document.getElementById('customChartBars');
    const inOutSel = document.getElementById('chartInOutSelect');
    const timeSel = document.getElementById('chartTimeSelect');
    const ysEl = document.getElementById('chartYearSelect');
    const msEl = document.getElementById('chartMonthSelect');
    const userSel = document.getElementById('chartUserSelect');
    
    if(!container || !inOutSel) return;

    const currentChartInOut = inOutSel.value;
    const currentChartTime = timeSel.value;
    const ys = ysEl ? ysEl.value : '';
    const ms = msEl ? msEl.value : '';
    const selectedUser = userSel ? userSel.value : '全付款人';
    
    const selectedProjects = window.getSelectedProjects();
    const selectedTargets = window.getSelectedChartTargets();
    
    let filtered = validExpenseData;
    
    const showCheckout = document.getElementById('chartShowCheckout') ? document.getElementById('chartShowCheckout').checked : false;
    if (!showCheckout) {
        filtered = filtered.filter(e => !isCheckoutRecord(e));
    }

    const showTransfer = document.getElementById('chartShowTransfer') ? document.getElementById('chartShowTransfer').checked : false;
    if (!showTransfer) {
        filtered = filtered.filter(e => e.mainCategory !== '轉帳');
        filtered = filtered.filter(e => !(e.mainCategory === '信用卡' && e.subCategory === '繳信用卡' && e.type === '支出'));
    }

    // 解決變數重複宣告 (ts2451)
    const isChartUnrepaid = document.getElementById('chartShowUnrepaid') ? document.getElementById('chartShowUnrepaid').checked : false;
    if (isChartUnrepaid) {
        filtered = filtered.filter(e => {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
            let isShared = tUsers.length > 1 || (tUsers.length === 1 && tUsers[0] !== e.user);
            let isRepaid = e.isRepaid === true || e.isRepaid === 'true';
            return isShared && !isRepaid;
        });
    }

    if (currentChartInOut === '支出') filtered = filtered.filter(e => e.type === '支出');
    else if (currentChartInOut === '收入') filtered = filtered.filter(e => e.type === '收入');

    if (selectedUser && selectedUser !== '全付款人') {
        filtered = filtered.filter(e => String(e.user).trim() === String(selectedUser).trim());
    }
    
    if (selectedTargets.length > 0 && !selectedTargets.includes('全被使用者')) {
        filtered = filtered.filter(e => {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()); 
            return tUsers.some(tu => selectedTargets.includes(tu));
        });
    }
    
    if (selectedProjects.length > 0 && !selectedProjects.includes('全專案')) {
        filtered = filtered.filter(e => {
            let p = String(e.project || '').trim();
            if (selectedProjects.includes('無專案') && p === '') return true;
            return selectedProjects.includes(p);
        });
    }

    filtered = filtered.filter(e => {
        if(!e.date) return false; 
        
        if (currentChartTime === '自訂區間') {
            const sDate = document.getElementById('chartStartDate').value;
            const eDate = document.getElementById('chartEndDate').value;
            if(sDate && e.date < sDate) return false;
            if(eDate && e.date > eDate) return false;
            return true;
        }

        let parts = e.date.split('-');
        if (parts.length < 2) return false;
        let y = parts[0], m = parseInt(parts[1], 10);
        
        if (currentChartTime === '本月') return y == ys && m == ms;
        else if (currentChartTime === '當年') return y == ys;
        return true; 
    });

    let groupMap = {};
    let totalIncome = 0;
    let totalExpense = 0;

    filtered.forEach(e => {
        let mainKey = e.mainCategory || '未分類';
        let subKey = e.subCategory || '未分類';
        
        let amt = Number(e.amount) || 0;
        let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
        let splitAmt = amt / Math.max(tUsers.length, 1);
        let effectiveAmt = amt;

        if (selectedTargets.length > 0 && !selectedTargets.includes('全被使用者')) {
            effectiveAmt = 0;
            let matchedTargets = tUsers.filter(tu => selectedTargets.includes(tu));
            matchedTargets.forEach(tu => {
                if (tu === '共用') {
                    effectiveAmt += amt / 2; // 共用的金額強制除以2算入
                } else {
                    effectiveAmt += splitAmt;
                }
            });
        }
        
        if (!groupMap[mainKey]) groupMap[mainKey] = { total: 0, subs: {} };

        if (e.type === '支出') {
            if (mainKey !== '轉帳' && !isCheckoutRecord(e)) totalExpense += effectiveAmt;
            if (currentChartInOut === '支出' || currentChartInOut === '總收支') {
                let val = (currentChartInOut === '總收支') ? -effectiveAmt : effectiveAmt;
                groupMap[mainKey].total += val;
                groupMap[mainKey].subs[subKey] = (groupMap[mainKey].subs[subKey] || 0) + val;
            }
        } else {
            if (mainKey !== '轉帳' && !isCheckoutRecord(e)) totalIncome += effectiveAmt;
            if (currentChartInOut === '收入' || currentChartInOut === '總收支') {
                groupMap[mainKey].total += effectiveAmt;
                groupMap[mainKey].subs[subKey] = (groupMap[mainKey].subs[subKey] || 0) + effectiveAmt;
            }
        }
    });

    const summaryBox = document.querySelector('.chart-summary-box');
    const summaryEl = document.getElementById('chartSummaryText');
    
    let prefixArr = [];
    if (selectedProjects.length > 0 && !selectedProjects.includes('全專案')) {
        prefixArr.push(`【${selectedProjects.join(' + ')}】`);
    }
    if (selectedUser !== '全付款人') prefixArr.push(`[${selectedUser}]`);
    let prefixStr = prefixArr.length > 0 ? prefixArr.join('') + ' ' : '';

    let labelPrefix = isChartUnrepaid ? "未還" : "總";

    if (currentChartInOut === '支出') {
        summaryEl.innerText = `${prefixStr}${labelPrefix}支出: $${totalExpense.toLocaleString()}`;
        summaryBox.style.background = '#ffebee'; summaryBox.style.borderColor = '#ffcdd2'; summaryEl.style.color = '#ef5350';
    } else if (currentChartInOut === '收入') {
        summaryEl.innerText = `${prefixStr}${labelPrefix}收入: $${totalIncome.toLocaleString()}`;
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
    
    container.innerHTML = mainKeys.map(mainName => {
        let mainObj = groupMap[mainName];
        const percent = Math.max((Math.abs(mainObj.total) / maxVal) * 100, 5); 
        
        let barColor = '#ef5350'; 
        if (currentChartInOut === '收入') barColor = '#4CAF50'; 
        else if (currentChartInOut === '總收支') barColor = mainObj.total >= 0 ? '#4CAF50' : '#ef5350';
        
        let displayVal = mainObj.total >= 0 ? `$${mainObj.total.toLocaleString()}` : `-$${Math.abs(mainObj.total).toLocaleString()}`;
        let displayLabel = `<span class="cat-emoji" style="width:22px;height:22px;font-size:13px;">${getIcon(mainName, '大分類')}</span> ${mainName}`;

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
    
    const selectedProjects = window.getSelectedProjects();
    const selectedTargets = window.getSelectedChartTargets();
    
    let filtered = validExpenseData;
    
    const showCheckout = document.getElementById('chartShowCheckout') ? document.getElementById('chartShowCheckout').checked : false;
    if (!showCheckout) {
        filtered = filtered.filter(e => !isCheckoutRecord(e));
    }

    const showTransfer = document.getElementById('chartShowTransfer') ? document.getElementById('chartShowTransfer').checked : false;
    if (!showTransfer) {
        filtered = filtered.filter(e => e.mainCategory !== '轉帳');
        filtered = filtered.filter(e => !(e.mainCategory === '信用卡' && e.subCategory === '繳信用卡' && e.type === '支出'));
    }

    // 解決變數重複宣告 (ts2451)
    const isDetailUnrepaid = document.getElementById('chartShowUnrepaid') ? document.getElementById('chartShowUnrepaid').checked : false;
    if (isDetailUnrepaid) {
        filtered = filtered.filter(e => {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
            let isShared = tUsers.length > 1 || (tUsers.length === 1 && tUsers[0] !== e.user);
            let isRepaid = e.isRepaid === true || e.isRepaid === 'true';
            return isShared && !isRepaid;
        });
    }

    if (currentChartInOut === '支出') filtered = filtered.filter(e => e.type === '支出');
    else if (currentChartInOut === '收入') filtered = filtered.filter(e => e.type === '收入');

    if (selectedUser && selectedUser !== '全付款人') {
        filtered = filtered.filter(e => String(e.user).trim() === String(selectedUser).trim());
    }
    
    if (selectedTargets.length > 0 && !selectedTargets.includes('全被使用者')) {
        filtered = filtered.filter(e => {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()); 
            return tUsers.some(tu => selectedTargets.includes(tu));
        });
    }
    
    if (selectedProjects.length > 0 && !selectedProjects.includes('全專案')) {
        filtered = filtered.filter(e => {
            let p = String(e.project || '').trim();
            if (selectedProjects.includes('無專案') && p === '') return true;
            return selectedProjects.includes(p);
        });
    }

    filtered = filtered.filter(e => {
        if(!e.date) return false; 
        
        if (currentChartTime === '自訂區間') {
            const sDate = document.getElementById('chartStartDate').value;
            const eDate = document.getElementById('chartEndDate').value;
            if(sDate && e.date < sDate) return false;
            if(eDate && e.date > eDate) return false;
            return true;
        }

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
            
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
            let isShared = tUsers.length > 1 || (tUsers.length === 1 && tUsers[0] !== e.user);

            let targetHtml = '';
            let nonSelfTargets = tUsers.filter(tu => tu !== e.user);
            if (nonSelfTargets.length > 0) {
                targetHtml = `<span style="color:#aaa; font-size:10px; margin:0 2px;">▶</span>` +
                             nonSelfTargets.map(tu => `<span class="detail-user-tag" style="${getUserColorStyle(tu)}">${tu}</span>`).join('<span style="margin:0 2px;"></span>');
            }
            
            let projectHtml = e.project ? `<span class="detail-project-tag">★ ${e.project}</span>` : '';

            let splitAmt = Number(e.amount) / Math.max(tUsers.length, 1);
            let displayAmt = Number(e.amount);

            if (selectedTargets.length > 0 && !selectedTargets.includes('全被使用者')) {
                displayAmt = 0;
                let matchedTargets = tUsers.filter(tu => selectedTargets.includes(tu));
                matchedTargets.forEach(tu => {
                    if (tu === '共用') {
                        displayAmt += Number(e.amount) / 2;
                    } else {
                        displayAmt += splitAmt;
                    }
                });
            }

            // 🌟 統計明細中的還款勾選框 (原生 checkbox, 白底黑字)
            let repayBadge = '';
            if (isShared && e.mainCategory !== '轉帳' && !isCheckoutRecord(e)) {
                let isRepaid = e.isRepaid === true || e.isRepaid === 'true';
                let statusText = isRepaid ? '已還' : '未還';
                let statusColor = isRepaid ? '#4CAF50' : '#ef5350';
                
                repayBadge = `
                <label style="display:inline-flex; align-items:center; gap:4px; cursor:pointer; margin:0 0 0 6px;">
                    <input type="checkbox" style="width:14px; height:14px; margin:0; cursor:pointer;" ${isRepaid ? 'checked' : ''} onchange="toggleRepayStatus('${e.id}', ${isRepaid})">
                    <span style="color:${statusColor}; font-weight:bold; font-size:13.5px;">${statusText}</span>
                </label>`;
            }

            // 🌟 備註純文字靠左對齊，無圖示無外框 (去除 padding-left)
            let contentHtml = '';
            if (e.content || repayBadge) {
                let textPart = e.content ? `<span style="font-size:13.5px; color:#333;">${e.content}</span>` : '';
                let divider = (e.content && repayBadge) ? `<span style="color:#ccc; font-size:13.5px; margin:0 4px;">|</span>` : '';
                
                contentHtml = `<div style="margin-top:6px;">
                    <div style="display:inline-flex; align-items:center; flex-wrap:wrap; max-width:100%;">
                        ${textPart}
                        ${divider}
                        ${repayBadge}
                    </div>
                </div>`;
            }

            return `
            <div style="padding:12px 0; border-bottom:1px solid #f0f0f0;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                    <div style="flex:1;">
                        <div style="display:flex; align-items:center; flex-wrap:wrap; gap:5px;">
                            <span style="font-size:13px; font-weight:bold; color:#666; margin-right:5px;">${window.formatSlashDate(e.date)}</span>
                            <span class="detail-tag">${e.subCategory}</span>
                            ${projectHtml}
                            <span class="detail-user-tag" style="${userStyle}">${e.user}</span>
                            ${targetHtml}
                        </div>
                        ${contentHtml}
                    </div>
                    <div style="text-align:right; min-width:70px;">
                        <div style="color:${color}; font-weight:bold; font-size:16px; margin-bottom:8px;">
                            ${sign}$${displayAmt.toLocaleString()}
                            ${displayAmt !== Number(e.amount) ? `<div style="font-size:10px; color:#999; margin-top:2px;">(總額 $${Number(e.amount).toLocaleString()})</div>` : ''}
                        </div>
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
// 比較頁面渲染邏輯
// ==========================================
window.renderComparePage = function() {
    const baseYearStr = document.getElementById('compareBaseYear').value;
    const endYearStr = document.getElementById('compareEndYear').value;
    const month = parseInt(document.getElementById('compareMonth').value, 10);
    const summaryContainer = document.getElementById('compareSummaryArea');
    const chartContainer = document.getElementById('compareChartArea');
    
    if(!summaryContainer || !chartContainer) return;

    const selectedCompareUsers = window.getSelectedCompareUsers();
    const isTotalMode = selectedCompareUsers.includes('總和 (不區分)');

    let xLabels = [];
    let cardTitle = "";
    let datasets = [];

    let expenses = validExpenseData.filter(e => e.type === '支出'); 

    const showCheckout = document.getElementById('compareShowCheckout') ? document.getElementById('compareShowCheckout').checked : false;
    if (!showCheckout) {
        expenses = expenses.filter(e => !isCheckoutRecord(e));
    }
    
    const showTransfer = document.getElementById('compareShowTransfer') ? document.getElementById('compareShowTransfer').checked : false;
    if (!showTransfer) {
        expenses = expenses.filter(e => e.mainCategory !== '轉帳');
        expenses = expenses.filter(e => !(e.mainCategory === '信用卡' && e.subCategory === '繳信用卡'));
    }

    // 解決變數重複宣告 (ts2451)
    const isCompareUnrepaid = document.getElementById('compareShowUnrepaid') ? document.getElementById('compareShowUnrepaid').checked : false;
    if (isCompareUnrepaid) {
        expenses = expenses.filter(e => {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim()).filter(x=>x);
            let isShared = tUsers.length > 1 || (tUsers.length === 1 && tUsers[0] !== e.user);
            let isRepaid = e.isRepaid === true || e.isRepaid === 'true';
            return isShared && !isRepaid;
        });
    }

    let uniqueYears = [];
    if (compareMode === 'currentYearMonth') {
        cardTitle = `${baseYearStr}年 月份支出趨勢`;
        xLabels = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
    } else if (compareMode === 'year') {
        let sY = Math.min(parseInt(baseYearStr), parseInt(endYearStr));
        let eY = Math.max(parseInt(baseYearStr), parseInt(endYearStr));
        cardTitle = `${sY}年 - ${eY}年 年度總支出`;
        for(let y = sY; y <= eY; y++) xLabels.push(y + "年");
    } else if (compareMode === 'month') {
        cardTitle = `歷年 ${month}月 支出比較`;
        uniqueYears = Array.from(new Set(validExpenseData.map(e => String(e.date).split('-')[0]))).sort();
        xLabels = uniqueYears.map(y => y + "年");
    }

    const now = new Date();
    const currentY = now.getFullYear();
    const currentM = now.getMonth() + 1;
    let validIndices = xLabels.map((lbl, i) => {
        if (compareMode === 'currentYearMonth') {
            let y = parseInt(baseYearStr, 10);
            let m = i + 1;
            return (y < currentY) || (y === currentY && m < currentM);
        } else if (compareMode === 'year') {
            let y = parseInt(lbl.replace('年', ''), 10);
            return y < currentY;
        } else if (compareMode === 'month') {
            let y = parseInt(lbl.replace('年', ''), 10);
            let m = month;
            return (y < currentY) || (y === currentY && m < currentM);
        }
        return true;
    });

    let userMap = {}; 
    if (isTotalMode) {
        userMap['總計支出'] = new Array(xLabels.length).fill(0);
    } else {
        selectedCompareUsers.forEach(u => userMap[u] = new Array(xLabels.length).fill(0));
    }

    expenses.forEach(e => {
        if(!e.date) return;
        
        let p = e.date.split('-');
        let y = p[0];
        let m = parseInt(p[1], 10);
        
        let index = -1;
        if (compareMode === 'currentYearMonth' && y === baseYearStr) {
            index = m - 1;
        } else if (compareMode === 'year') {
            let sY = Math.min(parseInt(baseYearStr), parseInt(endYearStr));
            let eY = Math.max(parseInt(baseYearStr), parseInt(endYearStr));
            let eYear = parseInt(y, 10);
            if(eYear >= sY && eYear <= eY) index = eYear - sY;
        } else if (compareMode === 'month' && m === month) {
            index = uniqueYears.indexOf(y);
        }

        if(index !== -1) {
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim());
            if (isTotalMode) {
                userMap['總計支出'][index] += Number(e.amount);
            } else {
                let splitAmt = Number(e.amount) / Math.max(tUsers.length, 1);
                tUsers.forEach(tUser => {
                    if (userMap[tUser]) {
                        if (tUser === '共用') {
                            userMap[tUser][index] += Number(e.amount) / 2;
                        } else {
                            userMap[tUser][index] += splitAmt;
                        }
                    }
                });
            }
        }
    });

    let globalTotal = 0;
    let globalDataArr = new Array(xLabels.length).fill(0);

    for (const [key, dataArr] of Object.entries(userMap)) {
        let color = '#2196F3';
        if (!isTotalMode) {
            let userObj = appData['使用者'].find(u => u.name === key);
            color = (userObj && userObj.color) ? userObj.color : getRandomColor();
        } else {
            if (compareMode === 'currentYearMonth') color = '#ff9bbb';
            else if (compareMode === 'year') color = '#2196F3';
            else color = '#4CAF50';
        }
        datasets.push({ label: key, data: dataArr, color: color });
        
        let sum = 0;
        dataArr.forEach((val, i) => {
            sum += val;
            globalDataArr[i] += val;
        });
        globalTotal += sum;
    }

    renderSvgChart(chartContainer, xLabels, datasets);

    let globalAvgTotal = 0;
    let globalAvgCount = 0;
    globalDataArr.forEach((v, i) => {
        if (validIndices[i] && v > 0) { 
            globalAvgTotal += v;
            globalAvgCount++;
        }
    });
    if (globalAvgCount === 0 && globalTotal > 0) {
        globalDataArr.forEach((v, i) => {
            if (v > 0) {
                globalAvgTotal += v;
                globalAvgCount++;
            }
        });
    }
    let avg = globalAvgCount > 0 ? Math.round(globalAvgTotal / globalAvgCount) : 0;
    
    let labelPrefix = isCompareUnrepaid ? "未還" : "總";

    let summaryHtml = `
        <div class="compare-card" style="margin-bottom:10px;">
            <div class="compare-card-title">${cardTitle}</div>
            <div class="compare-row main">
                <span class="compare-label">圖表${labelPrefix}支出</span>
                <span class="compare-val-main">$${globalTotal.toLocaleString()}</span>
            </div>
            <div class="compare-row" style="${!isTotalMode ? 'margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px dashed #ccc;' : ''}">
                <span class="compare-sub-label">時間節點平均支出</span>
                <span class="compare-diff gray">$${avg.toLocaleString()}</span>
            </div>
    `;

    if (!isTotalMode && datasets.length > 0) {
        summaryHtml += `<div style="font-size: 13px; color: #666; font-weight: bold; margin-bottom: 8px;">各別統計：</div>`;
        datasets.forEach(ds => {
            let dsTotal = ds.data.reduce((a, b) => a + b, 0);
            
            let dsAvgTotal = 0;
            let dsAvgCount = 0;
            ds.data.forEach((v, i) => {
                if (validIndices[i] && v > 0) { 
                    dsAvgTotal += v;
                    dsAvgCount++;
                }
            });
            if (dsAvgCount === 0 && dsTotal > 0) {
                ds.data.forEach(v => {
                    if (v > 0) {
                        dsAvgTotal += v;
                        dsAvgCount++;
                    }
                });
            }
            let dsAvg = dsAvgCount > 0 ? Math.round(dsAvgTotal / dsAvgCount) : 0;
            
            summaryHtml += `
                <div style="display:flex; justify-content:space-between; align-items:center; padding: 4px 0; font-size: 13px;">
                    <span style="display:flex; align-items:center;">
                        <span style="display:inline-block; width:10px; height:10px; background:${ds.color}; border-radius:50%; margin-right:6px;"></span>
                        ${ds.label}
                    </span>
                    <span style="font-weight:bold; color:#333;">總計: $${dsTotal.toLocaleString()} <span style="color:#aaa; font-weight:normal; margin-left:4px;">(均 $${dsAvg.toLocaleString()})</span></span>
                </div>
            `;
        });
    }

    summaryHtml += `</div>`;
    summaryContainer.innerHTML = summaryHtml;
};

// SVG 圖表繪製
function renderSvgChart(container, xLabels, datasets) {
    const width = container.clientWidth || 300;
    const height = 240;
    const padX = 45;
    const padY = 30;
    const chartW = width - padX * 2;
    const chartH = height - padY * 2;

    let allValues = [];
    datasets.forEach(ds => allValues.push(...ds.data));
    let maxVal = Math.max(...allValues, 100);

    let svg = `<svg width="100%" height="${height}">`;

    for(let i=0; i<=4; i++) {
        let y = padY + chartH - (i/4)*chartH;
        let val = Math.round((i/4)*maxVal);
        svg += `<line x1="${padX}" y1="${y}" x2="${width-padX}" y2="${y}" stroke="#eee" />`;
        svg += `<text x="${padX-5}" y="${y+4}" font-size="10" text-anchor="end" fill="#999">${val.toLocaleString()}</text>`;
    }

    let step = Math.ceil(xLabels.length / 6);
    xLabels.forEach((lbl, i) => {
        let x = padX + (i / (xLabels.length - 1)) * chartW;
        if(i % step === 0 || i === xLabels.length - 1) {
            svg += `<text x="${x}" y="${height-5}" font-size="10" text-anchor="middle" fill="#999">${lbl}</text>`;
        }
    });

    datasets.forEach(ds => {
        let points = ds.data.map((val, i) => {
            let x = padX + (i / (xLabels.length - 1)) * chartW;
            let y = padY + chartH - (val/maxVal) * chartH;
            return `${x},${y}`;
        }).join(' ');
        
        svg += `<polyline points="${points}" fill="none" stroke="${ds.color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />`;
        
        ds.data.forEach((val, i) => {
            if(val === 0 && ds.data.length > 12) return; 
            let x = padX + (i / (xLabels.length - 1)) * chartW;
            let y = padY + chartH - (val/maxVal) * chartH;
            svg += `<circle cx="${x}" cy="${y}" r="4" fill="white" stroke="${ds.color}" stroke-width="2" />`;
        });
    });

    if (datasets.length > 0 && datasets[0].data.length <= 12) {
        let labelsToDraw = [];
        
        datasets.forEach(ds => {
            ds.data.forEach((val, i) => {
                if(val === 0 && ds.data.length > 12) return;
                if(!labelsToDraw[i]) labelsToDraw[i] = [];
                let x = padX + (i / (xLabels.length - 1)) * chartW;
                let y = padY + chartH - (val/maxVal) * chartH;
                labelsToDraw[i].push({ val: val, x: x, y: y, color: ds.color });
            });
        });

        labelsToDraw.forEach(colLabels => {
            if(!colLabels) return;
            colLabels.sort((a,b) => b.y - a.y); 
            
            for(let j = 1; j < colLabels.length; j++) {
                let prev = colLabels[j-1];
                let curr = colLabels[j];
                if((prev.y - curr.y) < 14) {
                    curr.y = prev.y - 14;
                }
            }
            
            colLabels.forEach(lbl => {
                svg += `<text x="${lbl.x}" y="${lbl.y-10}" font-size="10" text-anchor="middle" font-weight="bold" fill="${lbl.color}" style="text-shadow: 1px 1px 0 #fff, -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff;">$${lbl.val.toLocaleString()}</text>`;
            });
        });
    }

    svg += `</svg>`;

    let legendHtml = '';
    if (datasets.length > 1) {
        legendHtml = '<div style="display:flex; justify-content:center; flex-wrap:wrap; gap:12px; margin-top:10px; padding:0 15px;">';
        datasets.forEach(ds => {
            legendHtml += `<div style="font-size:13px; display:flex; align-items:center; font-weight:bold; color:#555;"><span style="display:inline-block; width:12px; height:12px; background:${ds.color}; border-radius:50%; margin-right:6px;"></span>${ds.label}</div>`;
        });
        legendHtml += '</div>';
    }

    container.innerHTML = svg + legendHtml;
}

// ==========================================
// 表單操作與 Firestore 資料寫入
// ==========================================
window.quickAdd = async function(type) {
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
    
    const id = Date.now().toString();
    const payload = { type: type, id: id, name: newName, parentName: parentName, icon: icon, color: color };

    const btn = document.getElementById('submitBtn');
    const oldText = btn.innerText;
    btn.innerText = '新增中...'; btn.disabled = true;

    try {
        await setDoc(doc(db, "manageData", id), payload);
        appData[type].push(payload);
        window.updateFormDropdowns(); 
        if (type === '使用者') { 
            document.getElementById('user').value = payload.name; 
            document.getElementById('targetUser').value = payload.name;
        }
        if (type === '大分類') { document.getElementById('mainCategory').value = payload.name; window.updateFormSubCategory(); }
        if (type === '小分類') document.getElementById('subCategory').value = payload.name;
        if (type === '專案') document.getElementById('project').value = payload.name;
        window.renderManageLists(); 
    } catch(err) {
        alert("新增分類失敗");
    } finally {
        btn.innerText = oldText; btn.disabled = false;
    }
}

window.editExpense = function(id) {
    let exp = validExpenseData.find(e => e.id.toString() === id.toString()) || exceptionExpenseData.find(e => e.id.toString() === id.toString());
    if(!exp) return;

    document.getElementById('expenseId').value = exp.id;
    document.getElementById('expenseOperation').value = 'edit';
    document.getElementById('date').value = exp.date; 
    document.querySelector('select[name="type"]').value = exp.type === '收入' ? 'income' : 'expense';
    document.querySelector('input[name="amount"]').value = exp.amount;
    document.getElementById('payMethod').value = exp.payMethod || '存款'; 
    document.querySelector('input[name="content"]').value = exp.content || '';
    
    // 讀取是否已還款
    document.getElementById('isRepaidCheckbox').checked = (exp.isRepaid === true || exp.isRepaid === 'true');

    let userSel = document.getElementById('user');
    if(appData['使用者'].find(u => u.name === exp.user)) userSel.value = exp.user;
    else userSel.value = ""; 

    let targetSel = document.getElementById('targetUser');
    let tUserStr = exp.targetUser || exp.user; 
    let firstTarget = tUserStr.split(',')[0].trim();
    if(appData['使用者'].find(u => u.name === firstTarget)) targetSel.value = firstTarget;
    else targetSel.value = ""; 

    let mainSel = document.getElementById('mainCategory');
    if(appData['大分類'].find(c => c.name === exp.mainCategory)) mainSel.value = exp.mainCategory;
    else mainSel.value = "";

    window.updateFormSubCategory();
    let subSel = document.getElementById('subCategory');
    if(appData['小分類'].find(c => c.name === exp.subCategory)) subSel.value = exp.subCategory;
    else subSel.value = "";

    let projectSel = document.getElementById('project');
    if(appData['專案'].find(p => p.name === exp.project)) projectSel.value = exp.project;
    else projectSel.value = ""; 

    document.getElementById('submitBtn').innerText = '儲存修改';
    window.updateFormDropdowns();
    document.getElementById('addModal').style.display = 'flex';         
};

window.deleteExpense = async function(id) {
    if(!confirm('確定要永久刪除這筆明細嗎？')) return;
    
    validExpenseData = validExpenseData.filter(e => e.id.toString() !== id.toString());
    exceptionExpenseData = exceptionExpenseData.filter(e => e.id.toString() !== id.toString());
    window.renderDailyList();
    window.renderChart();
    window.renderExceptionList();
    
    if (document.getElementById('page-chart-detail').classList.contains('active')) window.openChartDetail(currentChartDetailLabel);
    if (document.getElementById('page-compare').classList.contains('active')) window.renderComparePage();
    
    try {
        await deleteDoc(doc(db, "expenses", id.toString()));
        window.fetchData(); 
    } catch(err) {
        alert('刪除失敗');
    }
};

document.getElementById('expenseForm').onsubmit = async function(e) {
    e.preventDefault();
    const btn = document.getElementById('submitBtn');
    btn.innerText = '處理中...'; btn.disabled = true;

    const formData = new FormData(this);
    const data = Object.fromEntries(formData.entries());
    
    if (!data.targetUser) data.targetUser = data.user;
    data.ledgerName = document.getElementById('globalLedgerSelect').value;
    // 儲存已還款狀態
    data.isRepaid = document.getElementById('isRepaidCheckbox').checked;

    const docId = data.id.toString(); 

    data.type = data.type === 'income' ? '收入' : '支出';
    data.amount = Number(data.amount) || 0;

    try {
        await setDoc(doc(db, "expenses", docId), data);
        alert(data.operation === 'add' ? '新增成功！' : '修改成功！');
        document.getElementById('addModal').style.display = 'none';
        window.fetchData(); 
    } catch(err) {
        alert('儲存失敗：' + err.message);
    } finally {
        btn.innerText = '儲存'; btn.disabled = false;
    }
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

window.deleteItem = async function(type, id) {
    if(!confirm(`確定要刪除這個${type}嗎？\n(相關明細將會被移動到「異常」分頁！)`)) return;
    try {
        await deleteDoc(doc(db, "manageData", id.toString()));
        window.fetchData();
    } catch(err) {
        alert('刪除失敗');
    }
};

document.getElementById('manageItemForm').onsubmit = async function(e) {
    e.preventDefault();
    const btn = document.getElementById('manageSubmitBtn');
    btn.innerText = '儲存中...'; btn.disabled = true;

    const type = document.getElementById('manageType').value;
    const parentName = type === '小分類' ? document.getElementById('manageParentCatInput').value : "";
    const id = document.getElementById('manageId').value.toString();

    const payload = {
        type: type, 
        id: id,
        name: document.getElementById('manageNameInput').value, 
        parentName: parentName, 
        icon: document.getElementById('manageIconInput').value.trim(), 
        color: document.getElementById('manageColorInput').value.trim()
    };

    try {
        await setDoc(doc(db, "manageData", id), payload);
        document.getElementById('manageItemModal').style.display = 'none';
        window.fetchData(); 
    } catch(err) {
        alert('儲存失敗：' + err.message);
    } finally {
        btn.innerText = '儲存'; btn.disabled = false;
    }
};

// 🌟 管理分頁邏輯餘額計算 (真實反映付款人的現金流)
window.getUserBalance = function(userName) {
    let deposit = 0;
    let credit = 0;
    
    validExpenseData.forEach(e => {
        let amt = Number(e.amount) || 0;
        let method = e.payMethod || '存款';
        let isTransfer = (e.mainCategory === '轉帳'); 

        if (isTransfer) {
            if (e.user === userName) deposit -= amt;
            let tUsers = (e.targetUser || e.user).split(',').map(x=>x.trim());
            if (tUsers.includes(userName) && e.user !== userName) {
                deposit += (amt / Math.max(tUsers.length, 1));
            }
        } else {
            if (e.user === userName) {
                if (method === '存款') {
                    if (e.type === '收入') deposit += amt;
                    else if (e.type === '支出') {
                        deposit -= amt;
                        if (e.mainCategory === '信用卡' && e.subCategory === '繳信用卡') {
                            credit += amt; 
                        }
                    }
                } else if (method === '信用卡') {
                    if (e.type === '收入') credit += amt; 
                    else if (e.type === '支出') credit -= amt; 
                }
            }
        }
    });
    return { deposit, credit };
};

window.openAdjustBalanceModal = function(userName) {
    currentAdjustUser = userName;
    currentSysBalances = window.getUserBalance(userName);
    
    document.getElementById('adjUserName').innerText = userName;
    document.getElementById('adjAccountType').value = '存款'; 
    
    window.updateAdjustDisplay();
    document.getElementById('adjustBalanceModal').style.display = 'flex';
};

window.updateAdjustDisplay = function() {
    let type = document.getElementById('adjAccountType').value;
    let sysBal = type === '存款' ? currentSysBalances.deposit : currentSysBalances.credit;
    
    document.getElementById('adjSysBalance').innerText = sysBal.toLocaleString();
    document.getElementById('adjActualBalance').value = sysBal;
    document.getElementById('adjDiffDisplay').innerText = '差額: $0 (無需調整)';
};

window.submitBalanceAdjustment = async function() {
    let actual = Number(document.getElementById('adjActualBalance').value) || 0;
    let type = document.getElementById('adjAccountType').value;
    let sysBal = type === '存款' ? currentSysBalances.deposit : currentSysBalances.credit;
    
    let diff = actual - sysBal;
    
    if (diff === 0) {
        document.getElementById('adjustBalanceModal').style.display = 'none';
        return;
    }

    const btn = document.getElementById('adjSubmitBtn');
    btn.innerText = '處理中...'; btn.disabled = true;

    const ledger = document.getElementById('globalLedgerSelect') ? document.getElementById('globalLedgerSelect').value : '日常帳本';
    const id = Date.now().toString();

    const data = {
        id: id,
        date: formatDate(new Date()),
        type: diff > 0 ? '收入' : '支出',
        amount: Math.abs(diff),
        payMethod: type, 
        user: currentAdjustUser,
        targetUser: currentAdjustUser,
        mainCategory: `未選擇大項目(${type}校正)`, 
        subCategory: '未選擇小項目',
        project: '',
        content: `手動${type}校正差額`,
        ledgerName: ledger,
        isRepaid: false
    };

    try {
        await setDoc(doc(db, "expenses", id), data);
        alert(`已將${type}差額新增至「異常」明細！\n請前往異常分頁補齊分類。`);
        document.getElementById('adjustBalanceModal').style.display = 'none';
        window.fetchData(); 
    } catch(err) {
        alert('校正失敗：' + err.message);
    } finally {
        btn.innerText = '確認新增差額'; btn.disabled = false;
    }
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

window.updateFormDropdowns = function() {
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

    window.updateFormSubCategory();
};

window.updateFormSubCategory = function() {
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
};