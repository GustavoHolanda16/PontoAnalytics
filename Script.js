// ─────────────────────────────────────────
// STATE
// ─────────────────────────────────────────
let employees = [];
let sortKey = 'he50';
let sortDir = -1;
let currentFileName = '';

// ─────────────────────────────────────────
// FILE HANDLING
// ─────────────────────────────────────────
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragging'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragging'));
dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragging');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
});

fileInput.addEventListener('change', e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
});

function processFile(file) {
    currentFileName = file.name;
    showLoading();

    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'html' || ext === 'htm' || ext === 'xls') {
        reader.onload = e => {
            try {
                const text = e.target.result;
                employees = parseHTMLTimesheet(text);
                if (employees.length === 0) throw new Error('Nenhum funcionário encontrado');
                showDashboard(file);
            } catch (err) {
                alert('Erro ao processar: ' + err.message);
                resetApp();
            }
        };
        reader.readAsText(file, 'UTF-8');
    } else {
        reader.onload = e => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const raw = XLSX.utils.sheet_to_csv(ws);
                employees = parseCSVTimesheet(raw);
                if (employees.length === 0) throw new Error('Nenhum funcionário encontrado');
                showDashboard(file);
            } catch (err) {
                alert('Erro ao processar: ' + err.message);
                resetApp();
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

// ─────────────────────────────────────────
// PARSER — HTML/XLS (HTML format exported from ponto system)
// ─────────────────────────────────────────
function parseHTMLTimesheet(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = doc.querySelectorAll('table');
    const results = [];

    for (let i = 0; i < tables.length; i += 4) {
        if (i + 3 >= tables.length) break;

        const infoTable = tables[i];
        const pontoTable = tables[i + 1];
        const summaryTable = tables[i + 3];

        // ── Info table ─────────────────────────────────────────────────
        // Structure: row[0]=empresa (1 td), row[1]=empty, row[2]=employee (1 td), row[3]=empty
        const infoRows = infoTable.querySelectorAll('tr');
        // Collect ALL text from ALL tds across all rows (concat everything)
        let infoText = '';
        infoRows.forEach(row => {
            row.querySelectorAll('td').forEach(td => { infoText += ' ' + td.textContent; });
        });
        infoText = infoText.replace(/\s+/g, ' ').trim();

        // Extract Nome — text between "Nome:" and "Matrícula:" (handle accent variants)
        const nomeMatch = infoText.match(/Nome:\s*([^]+?)\s*Matr[íi]cula:/i)
            || infoText.match(/Nome:\s*([^]+?)\s*Cargo:/i);
        const nome = nomeMatch ? nomeMatch[1].trim() : 'N/A';

        // Extract Cargo — between "Cargo:" and "PIS:"
        const cargoMatch = infoText.match(/Cargo:\s*(.*?)\s*PIS:/i);
        const cargo = cargoMatch ? cargoMatch[1].trim() : '';

        // Extract Admissão
        const admissaoMatch = infoText.match(/admiss[ãa]o:\s*(\d{2}\/\d{2}\/\d{4})/i);
        const admissao = admissaoMatch ? admissaoMatch[1] : '';

        // ── Summary table ──────────────────────────────────────────────
        // row[0]: single td with all stats concatenated
        // row[2]: single td with extras info
        const sumRows = summaryTable.querySelectorAll('tr');
        const summaryText = Array.from(sumRows[0]?.querySelectorAll('td') || []).map(td => td.textContent).join(' ').replace(/\s+/g, ' ');
        const extrasText = Array.from(sumRows[2]?.querySelectorAll('td') || []).map(td => td.textContent).join(' ').replace(/\s+/g, ' ');

        const dias = parseInt((summaryText.match(/Dias trabalhados:\s*(\d+)/i) || [])[1] || '0');
        const faltas = (summaryText.match(/Faltas:\s*([\d:]+)/i) || [])[1] || '00:00';
        const atrasos = (summaryText.match(/Atrasos:\s*([\d:]+)/i) || [])[1] || '00:00';
        const dsr = (summaryText.match(/DSR:\s*([\d:]+)/i) || [])[1] || '00:00';
        const he50 = (extrasText.match(/([\d:]+)\s*-\s*Hora Extra 50%/i) || [])[1] || '00:00';
        const folgaTrab = (extrasText.match(/([\d:]+)\s*-\s*Folga Trabalhada/i) || [])[1] || '00:00';

        // ── Ponto table ────────────────────────────────────────────────
        // row[0] and row[1] = header rows with <th> (0 <td>s) → skip rows that have no <td>
        // Data rows start from row[2]
        // Cols: 0=Data, 1=Horário, 2=Marcações, 3=H.Trab, 4=H.E., 8=Descontos
        const pontoRows = pontoTable.querySelectorAll('tr');
        let htrabTotal = '00:00';
        let sabCount = 0, domCount = 0;
        const dailyRows = [];

        pontoRows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells.length < 4) return; // skip header rows (have <th>, no <td>)

            const c = idx => (cells[idx]?.textContent?.trim() || '');
            const dateStr = c(0);
            const horario = c(1);
            const marcacoes = c(2);
            const htrab = c(3);
            const he = c(4);
            const descontos = c(8);

            // Totais row: col[2] = "Totais:", col[3] = total H.Trab
            if (marcacoes.includes('Totais')) {
                htrabTotal = htrab || '00:00';
                return;
            }

            if (!dateStr) return;

            const isSab = /s[áa]b/i.test(dateStr);
            const isDom = /\bdom\b/i.test(dateStr);
            const worked = htrab !== '';

            if (isSab && worked) sabCount++;
            if (isDom && worked) domCount++;

            dailyRows.push({
                date: dateStr,
                tipo: isSab ? 'Sábado' : (isDom ? 'Domingo' : 'Dia Útil'),
                horario: horario || '-',
                marcacoes: marcacoes || '-',
                htrab: worked ? htrab : '-',
                he: he || '-',
                descontos: descontos || '-'
            });
        });

        results.push({
            nome, cargo: cargo || '-', admissao,
            dias, htrab: htrabTotal, he50, folgaTrab,
            faltas, atrasos, dsr,
            sab: sabCount, dom: domCount,
            daily: dailyRows
        });
    }

    return results;
}

// ─────────────────────────────────────────
// PARSER — Generic XLSX/CSV
// ─────────────────────────────────────────
function parseCSVTimesheet(csv) {
    // Fallback: try to extract basic info
    const lines = csv.split('\n');
    const employees = [];
    // If the file is already structured as a summary table
    // look for header row
    let headerIdx = -1;
    for (let i = 0; i < Math.min(20, lines.length); i++) {
        if (lines[i].toLowerCase().includes('nome') && lines[i].toLowerCase().includes('h')) {
            headerIdx = i; break;
        }
    }

    if (headerIdx >= 0) {
        const headers = lines[headerIdx].split(',').map(h => h.trim().toLowerCase());
        for (let i = headerIdx + 1; i < lines.length; i++) {
            const cols = lines[i].split(',');
            if (cols.length < 3) continue;
            const obj = {};
            headers.forEach((h, idx) => obj[h] = (cols[idx] || '').trim().replace(/^"|"$/g, ''));
            if (obj.nome) {
                employees.push({
                    nome: obj.nome || 'N/A',
                    cargo: obj.cargo || '-',
                    admissao: obj['admissão'] || obj.admissao || '-',
                    dias: parseInt(obj['dias trabalhados'] || obj.dias || 0),
                    htrab: obj['h. trabalhadas'] || obj.htrab || '00:00',
                    he50: obj['h. extras (50%)'] || obj.he50 || '00:00',
                    folgaTrab: obj['h. folga trabalhada'] || obj.folga || '00:00',
                    faltas: obj.faltas || '00:00',
                    atrasos: obj.atrasos || '00:00',
                    dsr: obj.dsr || '00:00',
                    sab: parseInt(obj['sáb. trabalhados'] || obj.sab || 0),
                    dom: parseInt(obj['dom. trabalhados'] || obj.dom || 0),
                    daily: []
                });
            }
        }
    }

    return employees;
}

// ─────────────────────────────────────────
// UI TRANSITIONS
// ─────────────────────────────────────────
function showLoading() {
    document.getElementById('upload-section').style.display = 'none';
    document.getElementById('loading').classList.add('active');
    document.getElementById('dashboard').classList.remove('active');
}

function showDashboard(file) {
    setTimeout(() => {
        document.getElementById('loading').classList.remove('active');
        document.getElementById('dashboard').classList.add('active');

        document.getElementById('file-name-display').textContent = file.name;
        document.getElementById('file-meta-display').textContent =
            `${employees.length} funcionários · ${(file.size / 1024).toFixed(1)} KB · Processado em ${new Date().toLocaleTimeString('pt-BR')}`;

        renderKPIs();
        renderTable();
    }, 600);
}

function resetApp() {
    employees = [];
    document.getElementById('upload-section').style.display = '';
    document.getElementById('loading').classList.remove('active');
    document.getElementById('dashboard').classList.remove('active');
    document.getElementById('file-input').value = '';
    document.getElementById('search-box').value = '';
    document.getElementById('filter-extras').value = '';
}

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function toMin(s) {
    try {
        const p = String(s).split(':');
        return parseInt(p[0]) * 60 + parseInt(p[1]);
    } catch { return 0; }
}

function toHHMM(m) {
    m = Math.max(0, Math.round(m));
    return `${Math.floor(m / 60).toString().padStart(2, '0')}:${(m % 60).toString().padStart(2, '0')}`;
}

// ─────────────────────────────────────────
// KPIs
// ─────────────────────────────────────────
function renderKPIs() {
    const totalHE = employees.reduce((s, e) => s + toMin(e.he50) + toMin(e.folgaTrab), 0);
    const totalHT = employees.reduce((s, e) => s + toMin(e.htrab), 0);
    const totalSab = employees.reduce((s, e) => s + e.sab, 0);
    const totalDom = employees.reduce((s, e) => s + e.dom, 0);
    const avgDias = (employees.reduce((s, e) => s + e.dias, 0) / employees.length).toFixed(1);
    const comFaltas = employees.filter(e => toMin(e.faltas) > 0).length;

    const kpis = [
        { label: 'Funcionários', value: employees.length, sub: 'no período', cls: 'teal' },
        { label: 'H. Extras Total', value: toHHMM(totalHE), sub: '50% + folgas trabalhadas', cls: 'amber' },
        { label: 'H. Trabalhadas', value: toHHMM(totalHT), sub: 'soma da equipe', cls: 'teal' },
        { label: 'Sábados Trab.', value: totalSab, sub: 'ocorrências', cls: 'amber' },
        { label: 'Domingos Trab.', value: totalDom, sub: 'ocorrências', cls: 'red' },
        { label: 'Média Dias Trab.', value: avgDias, sub: 'dias por funcionário', cls: 'blue' },
        { label: 'Com Faltas', value: comFaltas, sub: `de ${employees.length} funcionários`, cls: 'red' },
    ];

    document.getElementById('kpi-grid').innerHTML = kpis.map(k => `
    <div class="kpi-card ${k.cls}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>
  `).join('');
}

// ─────────────────────────────────────────
// TABLE
// ─────────────────────────────────────────
function sortBy(key) {
    if (sortKey === key) sortDir *= -1;
    else { sortKey = key; sortDir = -1; }
    renderTable();
}

function renderTable() {
    const q = document.getElementById('search-box').value.toLowerCase();
    const f = document.getElementById('filter-extras').value;

    let data = employees.filter(e => {
        if (q && !e.nome.toLowerCase().includes(q) && !e.cargo.toLowerCase().includes(q)) return false;
        if (f === 'extras' && toMin(e.he50) + toMin(e.folgaTrab) === 0) return false;
        if (f === 'sabado' && e.sab === 0) return false;
        if (f === 'domingo' && e.dom === 0) return false;
        if (f === 'faltas' && toMin(e.faltas) === 0) return false;
        return true;
    });

    const keyMap = {
        nome: e => e.nome,
        cargo: e => e.cargo,
        dias: e => e.dias,
        htrab: e => toMin(e.htrab),
        he50: e => toMin(e.he50) + toMin(e.folgaTrab),
        folga: e => toMin(e.folgaTrab),
        faltas: e => toMin(e.faltas),
        atrasos: e => toMin(e.atrasos),
        sab: e => e.sab,
        dom: e => e.dom
    };

    data.sort((a, b) => {
        const va = keyMap[sortKey]?.(a) ?? 0;
        const vb = keyMap[sortKey]?.(b) ?? 0;
        if (typeof va === 'string') return sortDir * va.localeCompare(vb);
        return sortDir * (va - vb);
    });

    const tbody = document.getElementById('table-body');

    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" class="empty-rows">Nenhum funcionário encontrado</td></tr>`;
        return;
    }

    tbody.innerHTML = data.map((e, idx) => {
        const totalExtra = toMin(e.he50) + toMin(e.folgaTrab);
        const temFalta = toMin(e.faltas) > 0;
        const temAtraso = toMin(e.atrasos) > 60;

        const extraPill = totalExtra > 0
            ? `<span class="pill pill-amber">${toHHMM(totalExtra)}</span>`
            : `<span class="pill pill-neutral">00:00</span>`;

        const faltaPill = temFalta
            ? `<span class="pill pill-red">${e.faltas}</span>`
            : `<span class="pill pill-neutral">—</span>`;

        const sabPill = e.sab > 0
            ? `<span class="pill pill-amber">${e.sab}</span>`
            : `<span class="pill pill-neutral">—</span>`;

        const domPill = e.dom > 0
            ? `<span class="pill pill-red">${e.dom}</span>`
            : `<span class="pill pill-neutral">—</span>`;

        return `
      <tr>
        <td>
          <div class="td-name" title="${e.nome}">${e.nome}</div>
        </td>
        <td><div class="td-cargo" title="${e.cargo}">${e.cargo}</div></td>
        <td><span class="pill pill-teal">${e.dias}</span></td>
        <td style="font-family:'DM Mono',monospace;font-size:13px;">${e.htrab}</td>
        <td>${extraPill}</td>
        <td style="font-family:'DM Mono',monospace;font-size:12px;color:var(--text-dim);">${e.folgaTrab}</td>
        <td>${faltaPill}</td>
        <td style="font-family:'DM Mono',monospace;font-size:12px;color:${temAtraso ? 'var(--amber)' : 'var(--text-dim)'};">${e.atrasos}</td>
        <td>${sabPill}</td>
        <td>${domPill}</td>
        <td>
          <button onclick="openModal(${employees.indexOf(e)})" style="background:rgba(0,200,160,0.1);border:1px solid rgba(0,200,160,0.2);color:var(--teal);padding:5px 12px;border-radius:6px;cursor:pointer;font-size:12px;font-family:'Sora',sans-serif;transition:all 0.2s;" onmouseover="this.style.background='rgba(0,200,160,0.2)'" onmouseout="this.style.background='rgba(0,200,160,0.1)'">Ver</button>
        </td>
      </tr>
    `;
    }).join('');
}

// ─────────────────────────────────────────
// MODAL
// ─────────────────────────────────────────
function openModal(idx) {
    const e = employees[idx];
    document.getElementById('modal-title').textContent = e.nome;
    document.getElementById('modal-sub').textContent = `${e.cargo} · Admissão: ${e.admissao}`;

    const totalExtra = toMin(e.he50) + toMin(e.folgaTrab);

    document.getElementById('modal-kpis').innerHTML = [
        { label: 'Dias Trabalhados', val: e.dias, color: 'var(--teal)' },
        { label: 'H. Trabalhadas', val: e.htrab, color: 'var(--text)' },
        { label: 'H. Extras 50%', val: e.he50, color: 'var(--amber)' },
        { label: 'Folga Trabalhada', val: e.folgaTrab, color: 'var(--amber)' },
        { label: 'Total Extras', val: toHHMM(totalExtra), color: 'var(--green)' },
        { label: 'Faltas', val: e.faltas, color: toMin(e.faltas) > 0 ? 'var(--red)' : 'var(--text-dim)' },
        { label: 'Atrasos', val: e.atrasos, color: toMin(e.atrasos) > 30 ? 'var(--amber)' : 'var(--text-dim)' },
        { label: 'Sáb. Trabalhados', val: e.sab, color: e.sab > 0 ? 'var(--amber)' : 'var(--text-dim)' },
        { label: 'Dom. Trabalhados', val: e.dom, color: e.dom > 0 ? 'var(--red)' : 'var(--text-dim)' },
    ].map(k => `
    <div class="modal-kpi">
      <div class="modal-kpi-label">${k.label}</div>
      <div class="modal-kpi-val" style="color:${k.color}">${k.val}</div>
    </div>
  `).join('');

    document.getElementById('modal-days').innerHTML = e.daily.length > 0
        ? e.daily.map(d => {
            const isSab = d.tipo === 'Sábado';
            const isDom = d.tipo === 'Domingo';
            const isFalta = d.descontos.includes('Falta');
            const cardCls = isDom ? 'sunday' : (isSab ? 'weekend' : (isFalta ? 'falta' : ''));
            const tagCls = isDom ? 'dom' : (isSab ? 'sab' : 'util');
            return `
          <div class="day-card ${cardCls}">
            <div class="day-date">
              <span>${d.date.split(' ').slice(0, 2).join(' ')}</span>
              <span class="day-tag ${tagCls}">${d.tipo}</span>
            </div>
            <div class="day-detail">
              ${d.htrab !== '-' ? `H. Trab: <span>${d.htrab}</span><br>` : ''}
              ${d.he !== '-' ? `H. Extra: <span>${d.he}</span><br>` : ''}
              ${d.descontos !== '-' ? `<span style="color:var(--amber)">${d.descontos}</span>` : ''}
            </div>
          </div>
        `;
        }).join('')
        : '<div style="color:var(--text-dim);font-size:13px;">Sem detalhes diários disponíveis.</div>';

    document.getElementById('modal-overlay').classList.add('active');
    document.body.style.overflow = 'hidden';
}

function closeModal(e) {
    if (e.target === document.getElementById('modal-overlay')) closeModalDirect();
}

function closeModalDirect() {
    document.getElementById('modal-overlay').classList.remove('active');
    document.body.style.overflow = '';
}

// ─────────────────────────────────────────
// EXPORT CSV
// ─────────────────────────────────────────
function exportCSV() {
    const headers = ['Nome', 'Cargo', 'Admissão', 'Dias Trabalhados', 'H. Trabalhadas', 'H. Extras 50%', 'H. Folga Trabalhada', 'Total H. Extras', 'Faltas', 'Atrasos', 'DSR', 'Sábados Trabalhados', 'Domingos Trabalhados'];
    const rows = employees.map(e => [
        e.nome, e.cargo, e.admissao, e.dias, e.htrab, e.he50, e.folgaTrab,
        toHHMM(toMin(e.he50) + toMin(e.folgaTrab)),
        e.faltas, e.atrasos, e.dsr, e.sab, e.dom
    ]);

    const csv = [headers, ...rows].map(r => r.map(v => `"${v}"`).join(',')).join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'analise_ponto.csv';
    a.click();
    URL.revokeObjectURL(url);
}