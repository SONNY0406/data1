// app.js - 100% In-Browser In-Memory Excel Analysis & Visualization
// Zero network transmission, pure client-side processing using HTML5 FileReader + SheetJS + Chart.js

(function() {
  'use strict';

  // Global State in Browser Memory (Cleaned when page reloads)
  let parsedDataset = null; // { filename, dateKey, companies: [{key, name, color}], rows: [{date, ...values}] }
  let currentChart = null;
  let activeTechId = 1;

  // Preset theme colors for detected companies
  const COLOR_PALETTE = [
    { main: '#3b82f6', bg: 'rgba(59, 130, 246, 0.25)', border: '#60a5fa' },
    { main: '#10b981', bg: 'rgba(16, 185, 129, 0.25)', border: '#34d399' },
    { main: '#f43f5e', bg: 'rgba(244, 63, 94, 0.25)', border: '#fb7185' },
    { main: '#a855f7', bg: 'rgba(168, 85, 247, 0.25)', border: '#c084fc' },
    { main: '#f59e0b', bg: 'rgba(245, 158, 11, 0.25)', border: '#fbbf24' }
  ];

  // Mathematical & Statistical Utilities
  const stats = {
    mean: (arr) => arr.reduce((a, b) => a + b, 0) / arr.length,
    median: (arr) => {
      const sorted = [...arr].sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    },
    variance: (arr) => {
      const m = stats.mean(arr);
      return arr.reduce((sum, v) => sum + Math.pow(v - m, 2), 0) / (arr.length - 1);
    },
    stdDev: (arr) => Math.sqrt(stats.variance(arr)),
    cv: (arr) => (stats.stdDev(arr) / stats.mean(arr)) * 100, // Coefficient of Variation in %
    min: (arr) => Math.min(...arr),
    max: (arr) => Math.max(...arr),
    quantile: (arr, q) => {
      const sorted = [...arr].sort((a, b) => a - b);
      const pos = (sorted.length - 1) * q;
      const base = Math.floor(pos);
      const rest = pos - base;
      return sorted[base + 1] !== undefined ? sorted[base] + rest * (sorted[base + 1] - sorted[base]) : sorted[base];
    },
    linearRegression: (xVals, yVals) => {
      const n = xVals.length;
      const mx = stats.mean(xVals);
      const my = stats.mean(yVals);
      let num = 0, den = 0;
      for (let i = 0; i < n; i++) {
        num += (xVals[i] - mx) * (yVals[i] - my);
        den += Math.pow(xVals[i] - mx, 2);
      }
      const slope = num / den;
      const intercept = my - slope * mx;

      // Calculate R2 (Coefficient of Determination)
      const ssTot = yVals.reduce((acc, y) => acc + Math.pow(y - my, 2), 0);
      const ssRes = yVals.reduce((acc, y, i) => acc + Math.pow(y - (slope * xVals[i] + intercept), 2), 0);
      const r2 = 1 - (ssRes / ssTot);
      return { slope, intercept, r2 };
    }
  };

  // DOM Elements
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const browseBtn = document.getElementById('browseBtn');
  const uploadSection = document.getElementById('uploadSection');
  const dashboardSection = document.getElementById('dashboardSection');
  const headerActions = document.getElementById('headerActions');
  const reuploadBtn = document.getElementById('reuploadBtn');

  // Initialize
  function init() {
    setupUploadHandlers();
    setupTabNavigation();
  }

  // Setup Drag-and-Drop and File Dialog Events
  function setupUploadHandlers() {
    // Click to browse
    browseBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });
    dropZone.addEventListener('click', () => fileInput.click());

    // File input change
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) processLocalExcelFile(file);
    });

    // Drag & Drop effects
    ['dragenter', 'dragover'].forEach(name => {
      dropZone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dropzone-active');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropZone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dropzone-active');
      });
    });

    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const file = dt.files[0];
      if (file) processLocalExcelFile(file);
    });

    // Re-upload button in header
    reuploadBtn.addEventListener('click', () => {
      dashboardSection.classList.add('hidden');
      headerActions.classList.add('hidden');
      uploadSection.classList.remove('hidden');
      fileInput.value = '';
      if (currentChart) {
        currentChart.destroy();
        currentChart = null;
      }
    });
  }

  // 100% In-Memory Processing: Parse Excel with SheetJS
  function processLocalExcelFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert('엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다.');
      return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const buffer = new Uint8Array(e.target.result);
        // Parse workbook purely inside RAM
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        // Priority to '월간매출' or first sheet
        let sheetName = workbook.SheetNames.includes('월간매출') ? '월간매출' : workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!rawJson || rawJson.length < 2) {
          alert('엑셀 파일에 유효한 데이터 행이 부족합니다.');
          return;
        }

        // Header detection (Row 0)
        const headerRow = rawJson[0];
        const dateKey = headerRow[0] || '기준월';
        
        // Extract company columns (from Col 1 onwards)
        const companies = [];
        for (let c = 1; c < headerRow.length; c++) {
          const colName = String(headerRow[c] || '').trim();
          if (colName) {
            // Clean up name like '한빛테크 매출(억원)' -> '한빛테크'
            const cleanName = colName.replace(/\s*(매출|연매출|\(억원\)|\(원\)).*$/, '').trim() || colName;
            const palette = COLOR_PALETTE[(c - 1) % COLOR_PALETTE.length];
            companies.push({
              index: c,
              key: `comp_${c}`,
              rawHeader: colName,
              name: cleanName,
              color: palette.main,
              bg: palette.bg,
              border: palette.border
            });
          }
        }

        if (companies.length === 0) {
          alert('분석할 기업 매출 열을 찾을 수 없습니다. (1열: 기준월, 2열 이후: 기업별 매출)');
          return;
        }

        // Extract data rows
        const rows = [];
        for (let r = 1; r < rawJson.length; r++) {
          const rowData = rawJson[r];
          if (!rowData || rowData[0] === undefined || rowData[0] === null || String(rowData[0]).trim() === '') continue;

          const rowObj = { date: String(rowData[0]).trim() };
          let hasValidNum = false;

          companies.forEach(comp => {
            const val = parseFloat(rowData[comp.index]);
            const num = isNaN(val) ? 0 : val;
            rowObj[comp.key] = num;
            if (!isNaN(val)) hasValidNum = true;
          });

          if (hasValidNum) rows.push(rowObj);
        }

        if (rows.length === 0) {
          alert('유효한 수치 데이터 행이 존재하지 않습니다.');
          return;
        }

        // Save parsed data in client-side memory
        parsedDataset = {
          fileName: file.name,
          dateKey: dateKey,
          companies: companies,
          rows: rows
        };

        // Render dashboard
        displayDashboard();

      } catch (err) {
        console.error(err);
        alert('엑셀 파일을 읽는 도중 오류가 발생했습니다: ' + err.message);
      }
    };

    reader.readAsArrayBuffer(file);
  }

  // Display Dashboard Section and Set Overview Meta
  function displayDashboard() {
    uploadSection.classList.add('hidden');
    dashboardSection.classList.remove('hidden');
    headerActions.classList.remove('hidden');

    const { fileName, rows, companies } = parsedDataset;
    document.getElementById('metaFileName').innerText = fileName;
    document.getElementById('metaRowCount').innerText = `${rows.length}개 기간(월)`;
    document.getElementById('metaPeriod').innerText = `분석 기간: ${rows[0].date} ~ ${rows[rows.length - 1].date}`;

    // Render Quick KPIs in header
    renderQuickKPIs();

    // Render Raw Data Table
    renderRawTable();

    // Trigger Initial Technique (Technique 1: Descriptive Stats)
    loadTechnique(1);
  }

  // Render Top Quick KPIs
  function renderQuickKPIs() {
    const { rows, companies } = parsedDataset;
    const kpiContainer = document.getElementById('quickKpis');

    // Calculate sum for each company
    const compSums = companies.map(comp => {
      const sum = rows.reduce((acc, r) => acc + r[comp.key], 0);
      return { comp, sum };
    });
    const grandSum = compSums.reduce((acc, c) => acc + c.sum, 0);

    kpiContainer.innerHTML = `
      <div class="px-3 py-1.5 rounded-lg bg-blue-500/10 border border-blue-500/20 flex items-center gap-2">
        <span class="text-slate-400">시장 총매출:</span>
        <span class="font-bold text-blue-400 font-mono text-sm">${grandSum.toFixed(1)} 억원</span>
      </div>
      ${compSums.map(item => `
        <div class="px-2.5 py-1.5 rounded-lg bg-slate-800/80 border border-slate-700/80 flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full" style="background:${item.comp.color}"></span>
          <span class="text-slate-300 font-medium">${item.comp.name}:</span>
          <span class="font-bold text-white font-mono">${item.sum.toFixed(1)}억</span>
          <span class="text-[10px] text-slate-400">(${(item.sum / grandSum * 100).toFixed(1)}%)</span>
        </div>
      `).join('')}
    `;
  }

  // Setup 3 Core Statistical Technique Tabs
  function setupTabNavigation() {
    const tabs = document.querySelectorAll('#statTabs .tab-btn');
    tabs.forEach(btn => {
      btn.addEventListener('click', () => {
        const techId = parseInt(btn.getAttribute('data-tech'), 10);
        loadTechnique(techId);
      });
    });
  }

  // Load Technique by ID
  function loadTechnique(techId) {
    activeTechId = techId;

    // Update active tab styles
    const tabs = document.querySelectorAll('#statTabs .tab-btn');
    tabs.forEach(btn => {
      const id = parseInt(btn.getAttribute('data-tech'), 10);
      if (id === techId) {
        btn.className = "tab-btn active px-4 py-2 rounded-lg text-xs font-semibold transition flex items-center gap-2";
      } else {
        btn.className = "tab-btn bg-slate-800/80 hover:bg-slate-700/80 text-slate-300 px-4 py-2 rounded-lg text-xs font-semibold transition flex items-center gap-2";
      }
    });

    if (currentChart) {
      currentChart.destroy();
      currentChart = null;
    }

    if (techId === 1) renderTech1_Descriptive();
    else if (techId === 2) renderTech2_LinearTrend();
    else if (techId === 3) renderTech3_CVVolatility();
  }

  // ========================================================
  // TECHNIQUE 1: 기초 기술통계량 분석 (Descriptive Statistics)
  // ========================================================
  function renderTech1_Descriptive() {
    const { rows, companies } = parsedDataset;
    const ctx = document.getElementById('statChart').getContext('2d');

    document.getElementById('chartBadge').innerText = "핵심 통계 기법 1";
    document.getElementById('chartHeading').innerText = "1. 기초 기술통계량 비교 분석 (Descriptive Statistics)";
    document.getElementById('commentaryFormula').innerText = "평균 \\bar{X} = \\frac{1}{n}\\sum X_i, \\quad 표준편차 s = \\sqrt{\\frac{1}{n-1}\\sum(X_i - \\bar{X})^2}";
    document.getElementById('chartControls').innerHTML = '';

    // Calculate stats for each company
    const statsList = companies.map(comp => {
      const arr = rows.map(r => r[comp.key]);
      return {
        comp,
        mean: stats.mean(arr),
        median: stats.median(arr),
        std: stats.stdDev(arr),
        min: stats.min(arr),
        max: stats.max(arr),
        range: stats.max(arr) - stats.min(arr)
      };
    });

    // Render Chart
    currentChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ['평균 (Mean)', '중앙값 (Median)', '표준편차 (StdDev)', '최솟값 (Min)', '최댓값 (Max)'],
        datasets: statsList.map(st => ({
          label: `${st.comp.name} (억원)`,
          data: [st.mean, st.median, st.std, st.min, st.max].map(v => Number(v.toFixed(2))),
          backgroundColor: st.comp.bg,
          borderColor: st.comp.color,
          borderWidth: 2,
          borderRadius: 6
        }))
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#cbd5e1', font: { size: 12 } } },
          tooltip: {
            callbacks: {
              afterBody: () => '단위: 억원'
            }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#94a3b8' } },
          y: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#94a3b8' }, title: { display: true, text: '매출액 (억원)', color: '#64748b' } }
        }
      }
    });

    // Sub Table
    document.getElementById('chartSubTable').innerHTML = `
      <table class="w-full text-xs text-left text-slate-300">
        <thead class="text-[11px] text-slate-400 border-b border-slate-700">
          <tr>
            <th class="py-1.5 px-2">기업명</th>
            <th class="py-1.5 px-2 text-right">월평균(Mean)</th>
            <th class="py-1.5 px-2 text-right">중앙값(Median)</th>
            <th class="py-1.5 px-2 text-right">표준편차(SD)</th>
            <th class="py-1.5 px-2 text-right">최솟값</th>
            <th class="py-1.5 px-2 text-right">최댓값</th>
            <th class="py-1.5 px-2 text-right">변동범위(Max-Min)</th>
          </tr>
        </thead>
        <tbody>
          ${statsList.map(st => `
            <tr class="border-b border-slate-800">
              <td class="py-2 px-2 font-bold" style="color:${st.comp.color}">● ${st.comp.name}</td>
              <td class="py-2 px-2 text-right font-mono">${st.mean.toFixed(2)}억</td>
              <td class="py-2 px-2 text-right font-mono">${st.median.toFixed(2)}억</td>
              <td class="py-2 px-2 text-right font-mono">${st.std.toFixed(2)}억</td>
              <td class="py-2 px-2 text-right font-mono">${st.min.toFixed(2)}억</td>
              <td class="py-2 px-2 text-right font-mono">${st.max.toFixed(2)}억</td>
              <td class="py-2 px-2 text-right font-mono font-bold text-amber-400">${st.range.toFixed(2)}억</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    // Commentary
    document.getElementById('commentaryConcept').innerHTML = `
      데이터의 중심 위치(평균, 중앙값)와 흩어진 정도(표준편차, 최솟값, 최댓값)를 종합하여 각 기업의 <strong>기본 체급과 실적 분포 스케일</strong>을 가장 직관적으로 비교하는 통계 기법입니다.
    `;

    const highestMean = [...statsList].sort((a,b) => b.mean - a.mean)[0];
    const highestRange = [...statsList].sort((a,b) => b.range - a.range)[0];

    document.getElementById('commentaryResult').innerHTML = `
      <p>• <strong>최고 평균 매출 기업</strong>: <strong>${highestMean.comp.name}</strong>이(가) 월평균 <strong>${highestMean.mean.toFixed(2)}억원</strong>을 기록하며 전체적인 실적 체급에서 가장 앞서 있습니다.</p>
      <p>• <strong>외형 확장 폭(변동 범위)</strong>: <strong>${highestRange.comp.name}</strong>은(는) 최저 ${highestRange.min}억원에서 최고 ${highestRange.max}억원까지 격차가 무려 <strong>${highestRange.range.toFixed(2)}억원</strong>에 달해 가장 역동적인 외형 성장을 겪었습니다.</p>
      <p>• <strong>평균과 중앙값 일치도</strong>: 평균과 중앙값의 편차가 크지 않다는 것은 특정 월의 일시적 왜곡 없이 꾸준한 실적 분포를 유지하고 있음을 시사합니다.</p>
    `;

    document.getElementById('commentaryAction').innerHTML = `
      체급(평균)이 가장 높은 기업은 시장 선도자로서 규모의 경제를 활용하고, 변동 범위가 넓은 기업은 스케일업에 따른 인력 및 자금의 급격한 변동을 체계적으로 흡수할 수 있는 재무 버퍼가 필수적입니다.
    `;
  }

  // ========================================================
  // TECHNIQUE 2: 시계열 및 선형 회귀 추세 분석 (OLS Linear Trend)
  // ========================================================
  function renderTech2_LinearTrend() {
    const { rows, companies } = parsedDataset;
    const ctx = document.getElementById('statChart').getContext('2d');

    document.getElementById('chartBadge').innerText = "핵심 통계 기법 2";
    document.getElementById('chartHeading').innerText = "2. 시계열 및 선형 회귀 추세선 분석 (OLS Linear Trend)";
    document.getElementById('commentaryFormula').innerText = "y = \\beta_1 x + \\beta_0, \\quad \\beta_1 = \\frac{\\sum(x-\\bar{x})(y-\\bar{y})}{\\sum(x-\\bar{x})^2}, \\quad R^2 = 1 - \\frac{SS_{res}}{SS_{tot}}";
    document.getElementById('chartControls').innerHTML = '';

    const xVals = rows.map((_, i) => i + 1);
    const dates = rows.map(r => r.date);

    const regList = companies.map(comp => {
      const yVals = rows.map(r => r[comp.key]);
      const reg = stats.linearRegression(xVals, yVals);
      const trendLine = xVals.map(x => reg.slope * x + reg.intercept);
      return { comp, yVals, reg, trendLine };
    });

    const datasets = [];
    regList.forEach(item => {
      // Raw line
      datasets.push({
        label: `${item.comp.name} 실적`,
        data: item.yVals,
        borderColor: item.comp.color,
        backgroundColor: item.comp.color,
        borderWidth: 1.5,
        pointRadius: 1.5,
        tension: 0.2
      });
      // Trendline
      datasets.push({
        label: `${item.comp.name} 추세 (월 +${(item.reg.slope * 100).toFixed(0)}백만, R²=${item.reg.r2.toFixed(3)})`,
        data: item.trendLine,
        borderColor: item.comp.border,
        borderDash: [5, 4],
        borderWidth: 2,
        pointRadius: 0
      });
    });

    currentChart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: dates,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#cbd5e1', font: { size: 11 } } },
          tooltip: { mode: 'index', intersect: false }
        },
        scales: {
          x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', maxTicksLimit: 12 } },
          y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' }, title: { display: true, text: '매출액 (억원)', color: '#64748b' } }
        }
      }
    });

    // Sub Table
    document.getElementById('chartSubTable').innerHTML = `
      <div class="grid grid-cols-${Math.min(companies.length, 3)} gap-2 text-center text-xs">
        ${regList.map(item => `
          <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
            <div class="font-bold" style="color:${item.comp.color}">${item.comp.name} 회귀식</div>
            <div class="font-mono text-slate-200 mt-1">y = ${item.reg.slope.toFixed(4)}x + ${item.reg.intercept.toFixed(2)}</div>
            <div class="text-[11px] mt-0.5 ${item.reg.slope >= 0 ? 'text-emerald-400' : 'text-rose-400'} font-semibold">
              월평균 성장: ${item.reg.slope >= 0 ? '+' : ''}${(item.reg.slope * 100).toFixed(1)}백만원 / R²: ${(item.reg.r2 * 100).toFixed(1)}%
            </div>
          </div>
        `).join('')}
      </div>
    `;

    document.getElementById('commentaryConcept').innerHTML = `
      시간(x축: 개월 경과)에 따른 매출(y축)의 구조적 변화 방향을 <strong>최소자승법(OLS Linear Regression)</strong>으로 추정하여, 일시적 등락을 걷어내고 <strong>순수 장기 성장 동력과 성장 속도(기울기 $\\beta_1$)</strong>를 측정합니다.
    `;

    const topGrower = [...regList].sort((a,b) => b.reg.slope - a.reg.slope)[0];
    const decliner = regList.find(r => r.reg.slope < 0);

    document.getElementById('commentaryResult').innerHTML = `
      <p>• <strong>최고 성장률 기업</strong>: <strong>${topGrower.comp.name}</strong>은(는) 매월 약 <strong>${(topGrower.reg.slope * 100).toFixed(1)}백만원</strong>씩 매출이 우상향하고 있으며, 설명력($R^2$)이 <strong>${(topGrower.reg.r2 * 100).toFixed(1)}%</strong>에 달해 성장 추세가 매우 견고합니다.</p>
      ${decliner ? `<p>• <strong>역성장 관측 기업</strong>: <strong>${decliner.comp.name}</strong>은(는) 월평균 <strong>${(decliner.reg.slope * 100).toFixed(1)}백만원</strong>의 감소 기울기를 나타내고 있어 시장 내 수요 축소 또는 경쟁 심화에 따른 구조적 침식을 겪고 있습니다.</p>` : ''}
    `;

    document.getElementById('commentaryAction').innerHTML = `
      성장 기울기가 우수한 기업은 생산 인프라 확충과 고객 락인(Lock-in)에 공격적으로 투자해야 하며, 역성장 추세선이 확인된 기업은 기존 수익 모델의 한계를 인정하고 신규 사업으로의 피봇(Pivot) 전략을 수립해야 합니다.
    `;
  }

  // ========================================================
  // TECHNIQUE 3: 변동계수(CV) 리스크 및 안정성 분석
  // ========================================================
  function renderTech3_CVVolatility() {
    const { rows, companies } = parsedDataset;
    const ctx = document.getElementById('statChart').getContext('2d');

    document.getElementById('chartBadge').innerText = "핵심 통계 기법 3";
    document.getElementById('chartHeading').innerText = "3. 변동계수(CV) 및 리스크·안정성 비교 분석 (Volatility)";
    document.getElementById('commentaryFormula').innerText = "CV(\\%) = \\left(\\frac{s}{\\bar{X}}\\right) \\times 100 \\quad (s: 표준편차, \\; \\bar{X}: 평균)";
    document.getElementById('chartControls').innerHTML = '';

    const cvList = companies.map(comp => {
      const arr = rows.map(r => r[comp.key]);
      const mean = stats.mean(arr);
      const std = stats.stdDev(arr);
      const cvVal = stats.cv(arr);
      return { comp, mean, std, cv: cvVal };
    });

    currentChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: cvList.map(c => c.comp.name),
        datasets: [
          {
            label: '변동계수 CV (%) [상대적 변동 위험]',
            data: cvList.map(c => Number(c.cv.toFixed(2))),
            backgroundColor: cvList.map(c => c.comp.bg),
            borderColor: cvList.map(c => c.comp.color),
            borderWidth: 2,
            borderRadius: 8,
            yAxisID: 'y'
          },
          {
            label: '표준편차 (억원) [절대 변동폭]',
            data: cvList.map(c => Number(c.std.toFixed(2))),
            borderColor: '#f8fafc',
            backgroundColor: 'rgba(255, 255, 255, 0.2)',
            borderWidth: 2,
            type: 'line',
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#cbd5e1' } }
        },
        scales: {
          x: { ticks: { color: '#f1f5f9', font: { weight: 'bold' } } },
          y: {
            title: { display: true, text: '변동계수 CV (%)', color: '#60a5fa' },
            ticks: { color: '#60a5fa' },
            grid: { color: 'rgba(255,255,255,0.06)' }
          },
          y1: {
            position: 'right',
            title: { display: true, text: '표준편차 (억원)', color: '#f8fafc' },
            ticks: { color: '#f8fafc' },
            grid: { drawOnChartArea: false }
          }
        }
      }
    });

    // Sub Table
    document.getElementById('chartSubTable').innerHTML = `
      <div class="grid grid-cols-${Math.min(companies.length, 3)} gap-2 text-center text-xs">
        ${cvList.map(item => `
          <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
            <div class="font-bold" style="color:${item.comp.color}">${item.comp.name}</div>
            <div class="text-lg font-black text-white mt-1">${item.cv.toFixed(1)}%</div>
            <div class="text-[11px] text-slate-400 mt-0.5">표준편차: ${item.std.toFixed(2)}억 / 평균: ${item.mean.toFixed(2)}억</div>
          </div>
        `).join('')}
      </div>
    `;

    document.getElementById('commentaryConcept').innerHTML = `
      매출 규모가 상이한 기업 간에는 단순 표준편차만으로 위험도를 비교할 수 없습니다. 따라서 평균 대비 표준편차 비율인 <strong>변동계수(Coefficient of Variation, CV)</strong>를 산출하여 <strong>실적의 상대적 안정성과 예측 가능성</strong>을 공정하게 평가합니다.
    `;

    const sortedByCV = [...cvList].sort((a,b) => a.cv - b.cv);
    const mostStable = sortedByCV[0];
    const leastStable = sortedByCV[sortedByCV.length - 1];

    document.getElementById('commentaryResult').innerHTML = `
      <p>• <strong>최우수 안정성 기업</strong>: <strong>${mostStable.comp.name}</strong>이(가) 변동계수 <strong>${mostStable.cv.toFixed(1)}%</strong>로 가장 낮아, 외부 충격이나 계절적 부침에 흔들리지 않는 가장 예측 가능한 현금흐름을 증명했습니다.</p>
      <p>• <strong>상대적 고변동성 기업</strong>: <strong>${leastStable.comp.name}</strong>은(는) 변동계수가 <strong>${leastStable.cv.toFixed(1)}%</strong>로 가장 높게 나타났습니다. (고속 성장에 따른 스케일업 편차 또는 실적 불안정 요인 복합 작용)</p>
    `;

    document.getElementById('commentaryAction').innerHTML = `
      변동계수가 낮은 기업은 대외 신용도 평가 및 장기 자금 조달에 유리하므로 보수적 차입 경영에 적합하며, 변동계수가 높은 기업은 분기별 재고 부담 및 현금 유동성 위기를 예방하기 위한 안전마진 확보가 최우선입니다.
    `;
  }

  // ========================================================
  // Raw Data Table Render
  // ========================================================
  function renderRawTable() {
    const { rows, companies } = parsedDataset;
    const thead = document.getElementById('rawTableHead');
    const tbody = document.getElementById('rawTableBody');

    // Headers
    thead.innerHTML = `
      <th class="py-2.5 px-4 font-semibold text-slate-300">기준월 (Date)</th>
      ${companies.map(c => `
        <th class="py-2.5 px-4 font-semibold text-right" style="color:${c.color}">${c.name} (억원)</th>
      `).join('')}
      <th class="py-2.5 px-4 font-semibold text-right text-amber-300">합계 (억원)</th>
    `;

    // Rows
    tbody.innerHTML = rows.map(r => {
      const sum = companies.reduce((acc, c) => acc + r[c.key], 0);
      return `
        <tr class="hover:bg-slate-800/40 transition">
          <td class="py-2 px-4 text-slate-300 font-medium">${r.date}</td>
          ${companies.map(c => `
            <td class="py-2 px-4 text-right" style="color:${c.color}">${r[c.key].toFixed(2)}</td>
          `).join('')}
          <td class="py-2 px-4 text-right text-amber-300 font-bold">${sum.toFixed(2)}</td>
        </tr>
      `;
    }).join('');
  }

  // Run on DOM ready
  document.addEventListener('DOMContentLoaded', init);
})();
