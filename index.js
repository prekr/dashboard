// State Management
let items = ['예스24', '알라딘', '쿠팡'];
const months = Array.from({ length: 12 }, (_, i) => `${i + 1}월`);

let data = {
    goals: {
        '예스24': Array(12).fill(0),
        '알라딘': Array(12).fill(0),
        '쿠팡': Array(12).fill(0)
    },
    actuals: {
        '예스24': Array(12).fill(0),
        '알라딘': Array(12).fill(0),
        '쿠팡': Array(12).fill(0)
    }
};

// LocalStorage Integration
function saveToLocalStorage() {
    localStorage.setItem('dashboardItems', JSON.stringify(items));
    localStorage.setItem('dashboardData', JSON.stringify(data));
}

function loadFromLocalStorage() {
    const savedItems = localStorage.getItem('dashboardItems');
    const savedData = localStorage.getItem('dashboardData');
    
    if (savedItems && savedData) {
        items = JSON.parse(savedItems);
        data = JSON.parse(savedData);
    } else {
        // Fallback to demo data
        data.goals['예스24'] = [500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500];
        data.actuals['예스24'] = [100, 150, 200, 200, 250, 300, 0, 0, 0, 0, 0, 0];

        data.goals['알라딘'] = [800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800];
        data.actuals['알라딘'] = [500, 600, 750, 800, 850, 1000, 0, 0, 0, 0, 0, 0];

        data.goals['쿠팡'] = [1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000];
        data.actuals['쿠팡'] = [300, 400, 500, 600, 650, 800, 0, 0, 0, 0, 0, 0];
        saveToLocalStorage();
    }
}

let currentUnit = 1;

function formatDisplayValue(val) {
    if (val === undefined || val === null || isNaN(val)) return '0';
    if (currentUnit === 1) {
        return Math.round(val).toLocaleString();
    } else {
        let scaled = val / currentUnit;
        return scaled.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    }
}

function parseInputString(str) {
    let raw = String(str).replace(/[^0-9.-]/g, '');
    let val = Number(raw) || 0;
    return Math.round(val * currentUnit);
}

// Charts references
let monthlyChart = null;
let cumulativeChart = null;

// Initialize Chart Defaults Dark mode
Chart.defaults.color = 'rgba(255, 255, 255, 0.7)';
Chart.defaults.borderColor = 'rgba(255, 255, 255, 0.1)';
Chart.defaults.font.family = 'Inter, sans-serif';

// Computed Data
function getComputedData() {
    const monthlyTotalGoals = Array(12).fill(0);
    const monthlyTotal = Array(12).fill(0);
    const achievementRates = Array(12).fill(0);
    const cumulativeGoals = Array(12).fill(0);
    const cumulativeTotals = Array(12).fill(0);
    const cumulativeRates = Array(12).fill(0);

    const itemAnnualGoals = {};
    const itemAnnualActuals = {};

    items.forEach(item => {
        itemAnnualGoals[item] = 0;
        itemAnnualActuals[item] = 0;
    });

    let runGoal = 0;
    let runTotal = 0;

    for (let i = 0; i < 12; i++) {
        // Calculate monthly total
        let actual = 0;
        let goal = 0;
        items.forEach(item => {
            let a = Math.round(Number(data.actuals[item]?.[i] || 0));
            let g = Math.round(Number(data.goals[item]?.[i] || 0));
            itemAnnualActuals[item] += a;
            itemAnnualGoals[item] += g;
            actual += a;
            goal += g;
        });
        monthlyTotal[i] = actual;
        monthlyTotalGoals[i] = goal;

        // Calculate achievement rate
        achievementRates[i] = goal > 0 ? (actual / goal) * 100 : 0;

        // Cumulative
        runGoal += goal;
        runTotal += actual;
        cumulativeGoals[i] = runGoal;
        cumulativeTotals[i] = runTotal;
        cumulativeRates[i] = runGoal > 0 ? (runTotal / runGoal) * 100 : 0;
    }

    const totalAnnualGoal = runGoal;
    const totalAnnualActual = runTotal;
    const totalAnnualRate = runGoal > 0 ? (runTotal / runGoal) * 100 : 0;

    return {
        monthlyTotalGoals,
        monthlyTotal,
        achievementRates,
        cumulativeGoals,
        cumulativeTotals,
        cumulativeRates,
        itemAnnualGoals,
        itemAnnualActuals,
        totalAnnualGoal,
        totalAnnualActual,
        totalAnnualRate
    };
}

// Render DOM Table Table
function renderTable() {
    const computed = getComputedData();
    
    // 1. Render Headers
    const headerRow = document.getElementById('table-header-row');
    headerRow.innerHTML = '<th class="fixed-col">항목</th><th class="fixed-col-sub">구분</th>';
    months.forEach(m => {
        headerRow.innerHTML += `<th>${m}</th>`;
    });
    headerRow.innerHTML += `<th>합계</th>`;

    // 2. Render Body
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    // Items Row (Goals and Actuals for each item)
    items.forEach(item => {
        let goalHtml = `<td class="fixed-col" rowspan="2">${item}</td>`;
        goalHtml += `<td class="fixed-col-sub">목표</td>`;
        for (let i = 0; i < 12; i++) {
            let val = data.goals[item]?.[i] || 0;
            goalHtml += `<td><input type="text" class="data-input" data-type="goal" data-item="${item}" data-index="${i}" value="${formatDisplayValue(val)}" /></td>`;
        }
        goalHtml += `<td class="total-col" data-total-goal="${item}" style="color:var(--text-primary); font-weight:600">${formatDisplayValue(computed.itemAnnualGoals[item])}</td>`;
        const goalRow = document.createElement('tr');
        goalRow.className = 'goal-row';
        goalRow.innerHTML = goalHtml;
        tbody.appendChild(goalRow);

        let itemHtml = `<td class="fixed-col-sub">실적</td>`;
        for (let i = 0; i < 12; i++) {
            let val = data.actuals[item]?.[i] || 0;
            itemHtml += `<td><input type="text" class="data-input" data-type="actual" data-item="${item}" data-index="${i}" value="${formatDisplayValue(val)}" /></td>`;
        }
        itemHtml += `<td class="total-col" data-total-actual="${item}" style="color:var(--accent-primary); font-weight:600">${formatDisplayValue(computed.itemAnnualActuals[item])}</td>`;
        const itemRow = document.createElement('tr');
        itemRow.innerHTML = itemHtml;
        tbody.appendChild(itemRow);
    });

    // 3. Render Footer (Totals)
    const tfoot = document.getElementById('table-foot');
    tfoot.innerHTML = '';
    
    // Total Goal Row
    let totalGoalHtml = `<td class="fixed-col-span" colspan="2" style="color:var(--text-primary)">총 목표 (Total Goal)</td>`;
    for (let i = 0; i < 12; i++) {
        totalGoalHtml += `<td style="color:var(--text-primary)">${formatDisplayValue(computed.monthlyTotalGoals[i])}</td>`;
    }
    totalGoalHtml += `<td style="color:var(--text-primary); font-weight:600">${formatDisplayValue(computed.totalAnnualGoal)}</td>`;
    const totalGoalRow = document.createElement('tr');
    totalGoalRow.className = 'summary-row';
    totalGoalRow.innerHTML = totalGoalHtml;
    tfoot.appendChild(totalGoalRow);

    // Total Row
    let totalHtml = `<td class="fixed-col-span" colspan="2" style="color:var(--accent-primary)">총 실적 (Total Actual)</td>`;
    for (let i = 0; i < 12; i++) {
        totalHtml += `<td style="color:var(--accent-primary)">${formatDisplayValue(computed.monthlyTotal[i])}</td>`;
    }
    totalHtml += `<td style="color:var(--accent-primary); font-weight:600">${formatDisplayValue(computed.totalAnnualActual)}</td>`;
    const totalRow = document.createElement('tr');
    totalRow.className = 'summary-row';
    totalRow.innerHTML = totalHtml;
    tfoot.appendChild(totalRow);

    // Achiev. Rate Row
    let rateHtml = `<td class="fixed-col-span" colspan="2" style="color:var(--accent-success)">달성율 (%)</td>`;
    for (let i = 0; i < 12; i++) {
        let rate = computed.achievementRates[i];
        rateHtml += `<td style="color:var(--accent-success)">${rate.toFixed(1)}%</td>`;
    }
    rateHtml += `<td style="color:var(--accent-success); font-weight:600">${computed.totalAnnualRate.toFixed(1)}%</td>`;
    const rateRow = document.createElement('tr');
    rateRow.className = 'summary-row';
    rateRow.innerHTML = rateHtml;
    tfoot.appendChild(rateRow);

    attachInputListeners();
    updateCharts(computed);
}

// Event Listeners for data binding
function attachInputListeners() {
    const inputs = document.querySelectorAll('.data-input');
    inputs.forEach(input => {
        input.addEventListener('input', (e) => {
            let caretPos = e.target.selectionStart;
            let originalLen = e.target.value.length;
            
            let rawStr = e.target.value.replace(/[^0-9.-]/g, '');
            let val = parseInputString(rawStr);
            
            let formatted;
            if (currentUnit === 1) {
                let intVal = Number(rawStr.replace(/-/g, '')) || 0;
                let sign = rawStr.startsWith('-') ? '-' : '';
                formatted = rawStr === '' ? '' : rawStr === '-' ? '-' : sign + intVal.toLocaleString();
            } else {
                formatted = rawStr; // Let user type decimals freely
            }
            e.target.value = formatted;
            
            let newLen = formatted.length;
            caretPos = caretPos - (originalLen - newLen);
            e.target.setSelectionRange(caretPos, caretPos);

            const index = e.target.getAttribute('data-index');
            const type = e.target.getAttribute('data-type');
            
            if (type === 'goal') {
                const item = e.target.getAttribute('data-item');
                data.goals[item][index] = val;
            } else if (type === 'actual') {
                const item = e.target.getAttribute('data-item');
                data.actuals[item][index] = val;
            }
            
            saveToLocalStorage();
            updateFooterAndChartsOnly();
        });
        
        input.addEventListener('blur', (e) => {
            let rawStr = e.target.value.replace(/[^0-9.-]/g, '');
            if (rawStr === '' || rawStr === '-') { rawStr = '0'; }
            let val = parseInputString(rawStr);
            e.target.value = formatDisplayValue(val);
        });
    });
}

function updateFooterAndChartsOnly() {
    const computed = getComputedData();
    
    // Update individual item row totals
    items.forEach(item => {
        const goalTd = document.querySelector(`td[data-total-goal="${item}"]`);
        const actualTd = document.querySelector(`td[data-total-actual="${item}"]`);
        if (goalTd) goalTd.innerText = formatDisplayValue(computed.itemAnnualGoals[item]);
        if (actualTd) actualTd.innerText = formatDisplayValue(computed.itemAnnualActuals[item]);
    });

    const tfoot = document.getElementById('table-foot');
    
    // Update footer total row inner html loosely
    const goalTds = tfoot.children[0].children;
    const totalTds = tfoot.children[1].children;
    const rateTds = tfoot.children[2].children;

    for (let i = 0; i < 12; i++) {
        goalTds[i + 1].innerText = formatDisplayValue(computed.monthlyTotalGoals[i]);
        totalTds[i + 1].innerText = formatDisplayValue(computed.monthlyTotal[i]);
        rateTds[i + 1].innerText = computed.achievementRates[i].toFixed(1) + '%';
    }

    // Update column 13 (합계) for the footer
    goalTds[13].innerText = formatDisplayValue(computed.totalAnnualGoal);
    totalTds[13].innerText = formatDisplayValue(computed.totalAnnualActual);
    rateTds[13].innerText = computed.totalAnnualRate.toFixed(1) + '%';

    updateCharts(computed);
}

function getChartOptions() {
    return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: {
        mode: 'index',
        intersect: false,
    },
    scales: {
        y: {
            type: 'linear',
            display: true,
            position: 'left',
            grid: { color: 'rgba(255,255,255,0.05)' },
            ticks: {
                callback: function(value) { return formatDisplayValue(value); }
            }
        },
        y1: {
            type: 'linear',
            display: true,
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: {
                callback: function(value) { return value + '%' }
            }
        }
    },
    plugins: {
        legend: { labels: { usePointStyle: true } },
        tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.9)',
            titleColor: '#fff',
            bodyColor: '#cbd5e1',
            borderColor: 'rgba(255,255,255,0.1)',
            borderWidth: 1,
            callbacks: {
                label: function(context) {
                    let label = context.dataset.label || '';
                    if (label) {
                        label += ': ';
                    }
                    if (context.parsed.y !== null) {
                        if (context.dataset.yAxisID === 'y1') {
                            label += context.parsed.y.toFixed(1) + '%';
                        } else {
                            label += formatDisplayValue(context.parsed.y);
                        }
                    }
                    return label;
                }
            }
        }
    }
    };
}

function initCharts() {
    const ctxMonthly = document.getElementById('monthlyChart').getContext('2d');
    const ctxCumulative = document.getElementById('cumulativeChart').getContext('2d');

    monthlyChart = new Chart(ctxMonthly, {
        type: 'bar',
        data: { labels: months, datasets: [] },
        options: getChartOptions()
    });

    cumulativeChart = new Chart(ctxCumulative, {
        type: 'bar',
        data: { labels: months, datasets: [] },
        options: getChartOptions()
    });
}

const colors = [
    '#6366f1', // Indigo
    '#8b5cf6', // Violet
    '#ec4899', // Pink
    '#14b8a6', // Teal
    '#f59e0b'  // Amber
];

function updateCharts(computed) {
    if (!monthlyChart || !cumulativeChart) return;

    // Monthly Chart Update
    // Datasets: 1. Goal (Line), 2. Items (Stacked Bars), 3. Achievement Rate (Line on y1)
    const monthlyDatasets = [];
    
    monthlyDatasets.push({
        label: '총 목표 (Total Goal)',
        data: computed.monthlyTotalGoals,
        type: 'bar',
        backgroundColor: 'rgba(255, 255, 255, 0.1)',
        borderColor: 'rgba(255, 255, 255, 0.3)',
        borderWidth: 1,
        yAxisID: 'y',
        order: 2
    });

    items.forEach((item, idx) => {
        monthlyDatasets.push({
            label: `${item} 실적`,
            data: data.actuals[item],
            type: 'bar',
            backgroundColor: colors[idx % colors.length],
            yAxisID: 'y',
            stack: 'Stack 1',
            order: 2
        });
    });

    monthlyDatasets.push({
        label: '달성율 (%)',
        data: computed.achievementRates,
        type: 'line',
        borderColor: '#10b981',
        backgroundColor: '#10b981',
        borderWidth: 3,
        yAxisID: 'y1',
        tension: 0.4,
        order: 1
    });

    monthlyChart.data.datasets = monthlyDatasets;
    monthlyChart.update();

    // Cumulative Chart Update
    const cumulativeDatasets = [
        {
            label: '누적 목표',
            data: computed.cumulativeGoals,
            type: 'bar',
            backgroundColor: 'rgba(255, 255, 255, 0.1)',
            borderWidth: 1,
            yAxisID: 'y',
            order: 2
        },
        {
            label: '누적 실적',
            data: computed.cumulativeTotals,
            type: 'bar',
            backgroundColor: '#3b82f6',
            yAxisID: 'y',
            order: 2
        },
        {
            label: '누적 달성율 (%)',
            data: computed.cumulativeRates,
            type: 'line',
            borderColor: '#f59e0b',
            backgroundColor: '#f59e0b',
            borderWidth: 3,
            yAxisID: 'y1',
            tension: 0.4,
            order: 1
        }
    ];

    cumulativeChart.data.datasets = cumulativeDatasets;
    cumulativeChart.update();
}

// Add Item Modal Logic
document.getElementById('add-item-btn').addEventListener('click', () => {
    document.getElementById('add-item-modal').classList.add('active');
    document.getElementById('new-item-name').focus();
});

document.getElementById('cancel-add-btn').addEventListener('click', () => {
    document.getElementById('add-item-modal').classList.remove('active');
    document.getElementById('new-item-name').value = '';
});

document.getElementById('confirm-add-btn').addEventListener('click', () => {
    const input = document.getElementById('new-item-name');
    const name = input.value.trim();
    if (name && !items.includes(name)) {
        items.push(name);
        data.goals[name] = Array(12).fill(0);
        data.actuals[name] = Array(12).fill(0);
        saveToLocalStorage();
        renderTable();
        document.getElementById('add-item-modal').classList.remove('active');
        input.value = '';
    } else if (items.includes(name)) {
        alert('이미 존재하는 항목입니다.');
    }
});

// Excel Upload Logic
document.getElementById('upload-btn').addEventListener('click', () => {
    document.getElementById('excel-upload').click();
});

document.getElementById('excel-upload').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const fileData = new Uint8Array(e.target.result);
        const workbook = XLSX.read(fileData, {type: 'array'});
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        
        if (json.length > 1) {
            let tempItems = [];
            let tempGoals = {};
            let tempActuals = {};

            for (let r = 1; r < json.length; r++) {
                const row = json[r];
                if (!row || !row[0]) continue;
                
                const rowLabel = String(row[0]).trim();
                
                if (rowLabel.includes('총 실적') || rowLabel.includes('Total') || rowLabel.includes('달성율') || rowLabel.includes('%')) continue;

                if (rowLabel.includes('목표')) {
                    let itemName = rowLabel.replace('목표', '').trim();
                    if (!tempItems.includes(itemName)) { tempItems.push(itemName); }
                    if (!tempGoals[itemName]) { tempGoals[itemName] = Array(12).fill(0); }
                    for (let i = 1; i <= 12; i++) {
                        tempGoals[itemName][i-1] = Math.round(Number(row[i])) || 0;
                    }
                } else if (rowLabel.includes('실적')) {
                    let itemName = rowLabel.replace('실적', '').trim();
                    if (!tempItems.includes(itemName)) { tempItems.push(itemName); }
                    if (!tempActuals[itemName]) { tempActuals[itemName] = Array(12).fill(0); }
                    for (let i = 1; i <= 12; i++) {
                        tempActuals[itemName][i-1] = Math.round(Number(row[i])) || 0;
                    }
                }
            }

            if (tempItems.length > 0) {
                tempItems.forEach(item => {
                    if (!tempGoals[item]) tempGoals[item] = Array(12).fill(0);
                    if (!tempActuals[item]) tempActuals[item] = Array(12).fill(0);
                });

                items = [];
                tempItems.forEach(item => items.push(item));
                data.goals = tempGoals;
                data.actuals = tempActuals;
                saveToLocalStorage();
                renderTable();
            } else {
                alert('업로드한 엑셀 파일 형식을 확인해주세요. (1열: 구분, 각 행마다 항목의 목표/실적)');
            }
        }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input to allow re-upload
});

// Add Unit Selector Logic
document.querySelectorAll('.unit-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.unit-btn').forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        currentUnit = Number(e.target.getAttribute('data-unit'));
        renderTable();
    });
});

// App Entry Point
document.addEventListener('DOMContentLoaded', () => {
    loadFromLocalStorage();
    initCharts();
    renderTable();
});
