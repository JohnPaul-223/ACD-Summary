// =====================================================
//  ACD INTELLIGENCE ENGINE (PREMIUM DARK SYSTEM)
// =====================================================

let uploadedFiles = [];
const chartInstances = {};
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// --- Color System ---
const COLORS = {
  primary: '#ef4444',
  secondary: '#fafafa',
  accent1: '#3b82f6',
  accent2: '#10b981',
  text: '#fafafa',
  textMuted: '#a1a1aa',
  border: 'rgba(255, 255, 255, 0.08)',
  surface: '#121217'
};

const QUEUE_ALIAS = {
  'INFO MANDAUE QUEUE': 'MADAUE EXT/INFO', 'INFO MANDAUE': 'MADAUE EXT/INFO',
  'EXT_MANDAUE_QUEUE': 'MADAUE EXT/INFO', 'EXT MANDAUE': 'MADAUE EXT/INFO',
  'INT_MANDAUE_QUEUE': 'MANDAUE INT', 'INT MANDAUE': 'MANDAUE INT',
  'INT_MADAUE_QUEUE': 'MANDAUE INT', 'INT MADAUE': 'MANDAUE INT',
  'EXT_FUENTE_QUEUE': 'FUENTE EXT', 'EXT FUENTE': 'FUENTE EXT',
  'INT_FUENTE_QUEUE': 'FUENTE INT', 'INT FUENTE': 'FUENTE INT',
};
const SELECTED_QUEUES = ['MADAUE EXT/INFO','MANDAUE INT','FUENTE EXT','FUENTE INT'];

function canonicalize(raw) {
  const key = raw.trim().toUpperCase();
  for (const [k, v] of Object.entries(QUEUE_ALIAS)) {
    if (key === k.toUpperCase()) return v;
  }
  return raw.trim().replace(/_QUEUE$/i,'').replace(/_/g,' ').replace(/\bQUEUE\b/gi,'').trim().toUpperCase();
}

// --- logic ---
const dz = document.getElementById('dropZone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag-over');
  addFiles([...e.dataTransfer.files]);
});
document.getElementById('fileInput').addEventListener('change', e => {
  addFiles([...e.target.files]); e.target.value='';
});

function addFiles(files) {
  files.forEach(f => {
    if (!uploadedFiles.find(x => x.name===f.name && x.size===f.size)) uploadedFiles.push(f);
  });
  renderFileList();
}
function removeFile(i) { uploadedFiles.splice(i,1); renderFileList(); }
function clearFiles() {
  uploadedFiles=[];
  renderFileList();
  document.getElementById('reportSection').style.display='none';
  document.getElementById('uploadSection').style.display='flex';
}
function renderFileList() {
  const list=document.getElementById('fileList');
  document.getElementById('clearBtn').style.display=uploadedFiles.length?'inline-block':'none';
  document.getElementById('generateBtn').disabled=!uploadedFiles.length;
  list.innerHTML='';
  uploadedFiles.forEach((f,i)=>{
    const sz=f.size>1024*1024?(f.size/1024/1024).toFixed(1)+' MB':(f.size/1024).toFixed(1)+' KB';
    const d=document.createElement('div'); d.className='file-item';
    d.innerHTML=`<span class="file-name">${f.name}</span><span class="file-size">${sz}</span><button class="file-remove" onclick="removeFile(${i})">✕</button>`;
    list.appendChild(d);
  });
}

function parseFile(buf) {
  const wb = XLSX.read(buf, {type:'array', raw:false});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  let startDate=null, endDate=null;
  const queues=[];
  for (const row of rows) {
    const cells = row.map(c => String(c).trim());
    const line = cells.join('|');
    const dates = line.match(/\d{4}-\d{2}-\d{2}/g);
    if (dates && dates.length>=2 && !startDate) { startDate = new Date(dates[0]); endDate = new Date(dates[1]); }
    let name='';
    for (const c of cells) {
      if (!c) continue;
      const cl = c.toLowerCase();
      if (cl.includes('period')||cl==='queue'||cl.includes('acd incoming')) break;
      if (isNaN(c) && !/^\d+:\d+/.test(c) && c.length>2) { name=c; break; }
    }
    if (!name || name.toLowerCase()==='summary') continue;
    let calls=0, totalTime='';
    for (const c of cells) {
      if (/^\d{1,6}$/.test(c) && +c>0 && !calls) { calls=+c; continue; }
      if (/^\d+:\d{2}:\d{2}$/.test(c) && !totalTime) { totalTime=c; continue; }
    }
    if (calls>0) queues.push({ raw:name, canonical:canonicalize(name), calls });
  }
  return { startDate, endDate, queues };
}

function spreadMonths(start, end, calls) {
  const out={}; if (!start||!end||!calls) return out;
  const months=[];
  let cur=new Date(start.getFullYear(), start.getMonth(), 1);
  const last=new Date(end.getFullYear(), end.getMonth(), 1);
  while (cur<=last) { months.push(new Date(cur)); cur.setMonth(cur.getMonth()+1); }
  const weights=months.map(ms=>{
    const me=new Date(ms.getFullYear(), ms.getMonth()+1, 0);
    const os=ms<start?start:ms, oe=me>end?end:me;
    return Math.max(0,(oe-os)/86400000+1);
  });
  const tot=weights.reduce((a,b)=>a+b,0)||1;
  months.forEach((ms,i)=>{
    const mk=MON[ms.getMonth()]+' '+ms.getFullYear();
    out[mk]=(out[mk]||0)+Math.round(calls*weights[i]/tot);
  });
  return out;
}

function sortMonths(keys) {
  return [...new Set(keys)].sort((a,b)=>{
    const [ma,ya]=a.split(' '), [mb,yb]=b.split(' ');
    return (+ya - +yb) || (MON.indexOf(ma)-MON.indexOf(mb));
  });
}

// =====================================================
async function generateReport() {
  if (!uploadedFiles.length) return;
  document.getElementById('loadingOverlay').style.display='flex';
  document.getElementById('reportSection').style.display='none';
  await new Promise(r=>setTimeout(r,60));

  const rawPeriods=[];
  for (const f of uploadedFiles) {
    const buf=await f.arrayBuffer();
    rawPeriods.push(parseFile(buf));
  }

  const qmap={};
  const monthSet=new Set();
  for (const p of rawPeriods) {
    for (const q of p.queues) {
      const cn=q.canonical;
      if (!qmap[cn]) qmap[cn]={};
      const slices=spreadMonths(p.startDate, p.endDate, q.calls);
      for (const [mk,v] of Object.entries(slices)) {
        qmap[cn][mk]=(qmap[cn][mk]||0)+v;
        monthSet.add(mk);
      }
    }
  }
  const allMonths=sortMonths([...monthSet]);
  for (const q of SELECTED_QUEUES) { if (!qmap[q]) qmap[q] = {}; }

  // Clean old charts
  Object.keys(chartInstances).forEach(k=>{ try{chartInstances[k].destroy();}catch(e){} delete chartInstances[k]; });

  // Build Viz
  buildComparisonChart('extCompareChart', qmap, allMonths, ['MADAUE EXT/INFO', 'FUENTE EXT'], 'External Call Volume');
  buildComparisonChart('intCompareChart', qmap, allMonths, ['MANDAUE INT', 'FUENTE INT'], 'Internal Call Volume');
  buildFullTrendGraph(qmap, allMonths, SELECTED_QUEUES);
  buildDataTable(qmap, allMonths, SELECTED_QUEUES);

  document.getElementById('loadingOverlay').style.display='none';
  document.getElementById('reportSection').style.display='block';
  document.getElementById('uploadSection').style.display='none';
  window.scrollTo({top:0, behavior:'smooth'});
}

// --- Visual Components ---

Chart.register(ChartDataLabels);

function buildComparisonChart(canvasId, qmap, labels, queues, title) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  const datasets = queues.map((q, i) => ({
    label: q,
    data: labels.map(m => qmap[q][m] || 0),
    backgroundColor: i === 0 ? COLORS.primary : COLORS.secondary,
    borderRadius: 6,
    barPercentage: 0.7,
    datalabels: {
      align: 'end',
      anchor: 'end',
      color: i === 1 ? COLORS.surface : COLORS.text,
      font: { family: 'Outfit', weight: 'bold', size: 10 },
      formatter: (val) => val > 0 ? val.toLocaleString() : ''
    }
  }));

  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 20 } },
      plugins: {
        legend: { position: 'top', labels: { color: COLORS.text, font: { family: 'Outfit', size: 12 } } },
        tooltip: { backgroundColor: COLORS.surface, titleColor: COLORS.primary, bodyColor: COLORS.text, borderColor: COLORS.border, borderWidth: 1 },
        datalabels: { display: true }
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: COLORS.textMuted, font: { family: 'Outfit' } } },
        y: { 
          beginAtZero: true, 
          grid: { color: COLORS.border }, 
          ticks: { color: COLORS.textMuted, font: { family: 'Outfit' }, callback: v => v.toLocaleString() } 
        }
      }
    }
  });
}

function buildFullTrendGraph(qmap, labels, queues) {
  const ctx = document.getElementById('full-graph').getContext('2d');
  const pal = [COLORS.primary, COLORS.secondary, COLORS.accent1, COLORS.accent2];
  
  const datasets = queues.map((q, i) => {
    const data = labels.map(m => qmap[q][m] || 0);
    const maxVal = Math.max(...data);
    
    return {
      label: q,
      data: data,
      borderColor: pal[i % pal.length],
      borderWidth: 3,
      pointBackgroundColor: pal[i % pal.length],
      tension: 0.3,
      pointRadius: 4,
      fill: false,
      datalabels: {
        display: (context) => {
          // Only show label if it's the peak value
          return context.dataset.data[context.dataIndex] === maxVal && maxVal > 0;
        },
        align: 'top',
        offset: 8,
        color: pal[i % pal.length] === COLORS.secondary ? COLORS.surface : '#fff',
        backgroundColor: pal[i % pal.length],
        borderRadius: 4,
        padding: 4,
        font: { family: 'Outfit', weight: 'bold', size: 11 },
        formatter: (val) => `PEAK: ${val.toLocaleString()}`
      }
    };
  });

  chartInstances['full-graph'] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 25 } },
      plugins: {
        legend: { position: 'bottom', labels: { color: COLORS.text, font: { family: 'Outfit', size: 12 }, padding: 20 } },
        tooltip: { mode: 'index', intersect: false, backgroundColor: COLORS.surface, titleColor: COLORS.primary, bodyColor: COLORS.text },
        datalabels: { display: true }
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: COLORS.textMuted, font: { family: 'Outfit' } } },
        y: { 
          beginAtZero: true, 
          grid: { color: COLORS.border }, 
          ticks: { color: COLORS.textMuted, font: { family: 'Outfit' }, callback: v => v.toLocaleString() } 
        }
      }
    }
  });
}

function buildDataTable(qmap, allMonths, allQueues) {
  const el = document.getElementById('dataTable');
  let h = `<table class="data-table"><thead><tr><th>Queue</th>`;
  allMonths.forEach(m => h += `<th class="num">${m}</th>`);
  h += `<th class="num">Total</th></tr></thead><tbody>`;

  const mTotals = {};
  allMonths.forEach(m => mTotals[m] = allQueues.reduce((s, q) => s + (qmap[q][m] || 0), 0));

  allQueues.forEach(q => {
    const qTotal = allMonths.reduce((s, m) => s + (qmap[q][m] || 0), 0);
    h += `<tr><td><span style="font-weight:600">${q}</span></td>`;
    allMonths.forEach(m => h += `<td class="num">${(qmap[q][m] || 0).toLocaleString()}</td>`);
    h += `<td class="num"><strong>${qTotal.toLocaleString()}</strong></td></tr>`;
  });

  const grand = Object.values(mTotals).reduce((a, b) => a + b, 0);
  h += `<tr><td><strong>SYSTEM TOTAL</strong></td>`;
  allMonths.forEach(m => h += `<td class="num"><strong>${mTotals[m].toLocaleString()}</strong></td>`);
  h += `<td class="num" style="color:var(--primary)"><strong>${grand.toLocaleString()}</strong></td></tr></tbody></table>`;
  
  el.innerHTML = h;
}

// --- Exporters ---
function saveChartPng(canvasId, title) {
  const canvas = document.getElementById(canvasId);
  const temp = document.createElement('canvas');
  temp.width = canvas.width; temp.height = canvas.height;
  const tctx = temp.getContext('2d');
  tctx.fillStyle = COLORS.surface;
  tctx.fillRect(0, 0, temp.width, temp.height);
  tctx.drawImage(canvas, 0, 0);
  const a = document.createElement('a');
  a.download = title + '.png';
  a.href = temp.toDataURL('image/png', 1.0);
  a.click();
}

async function saveTablePng(elementId, title) {
  const el = document.getElementById(elementId);
  const clone = el.cloneNode(true);
  const container = document.createElement('div');
  container.style.position = 'absolute'; container.style.left = '-9999px'; container.style.width = 'fit-content';
  container.style.background = COLORS.surface; container.style.padding = '30px'; container.style.color = COLORS.text;
  document.body.appendChild(container); container.appendChild(clone);
  await new Promise(r => setTimeout(r, 100));
  const canvas = await html2canvas(clone, { scale: 2, backgroundColor: COLORS.surface, width: clone.scrollWidth, height: clone.scrollHeight, windowWidth: clone.scrollWidth + 500 });
  const a = document.createElement('a'); a.download = title + '.png'; a.href = canvas.toDataURL('image/png', 1.0); a.click();
  document.body.removeChild(container);
}
