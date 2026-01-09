// app.js
(() => {
  'use strict';

  const LOGO_URL = 'https://tasatop.com.pe/wp-content/uploads/elementor/thumbs/logos-17-r320c27cra7m7te2fafiia4mrbqd3aqj7ifttvy33g.png';

  // ====== DOM ======
  const el = {
    logoImg: document.getElementById('logoImg'),
    emailLogo: document.getElementById('emailLogo'),
    fileInput: document.getElementById('fileInput'),
    fileName: document.getElementById('fileName'),
    statusBadge: document.getElementById('statusBadge'),

    companySelect: document.getElementById('companySelect'),
    monthSelect: document.getElementById('monthSelect'),
    yearSelect: document.getElementById('yearSelect'),

    alertBox: document.getElementById('alertBox'),
    warnBox: document.getElementById('warnBox'),

    rowsCount: document.getElementById('rowsCount'),
    totalDeposit: document.getElementById('totalDeposit'),
    warnCount: document.getElementById('warnCount'),

    exportExcelBtn: document.getElementById('exportExcelBtn'),
    exportPdfBtn: document.getElementById('exportPdfBtn'),
    copyEmailBtn: document.getElementById('copyEmailBtn'),

    tableHead: document.getElementById('tableHead'),
    tableBody: document.getElementById('tableBody'),

    prevPageBtn: document.getElementById('prevPageBtn'),
    nextPageBtn: document.getElementById('nextPageBtn'),
    pageInfo: document.getElementById('pageInfo'),

    addRowGrid: document.getElementById('addRowGrid'),
    addRowBtn: document.getElementById('addRowBtn'),

    emailText: document.getElementById('emailText')
  };

  // ====== State ======
  const STATE = {
    rawRows: [],
    normalizedRows: [],   // rows with parsed columns and types
    companies: [],
    selectedCompany: '',
    reportRows: [],       // editable rows for selected company
    warnings: [],

    page: 1,
    pageSize: 50,

    logoDataUrl: null,
    logoFailed: false
  };

  const REPORT_COLS = [
    { key: 'codigo', label: 'Código', type: 'text', editable: true, required: true },
    { key: 'inversionista', label: 'Inversionista', type: 'text', editable: true, required: true },
    { key: 'fechas_pago', label: 'Fechas de pago', type: 'dateText', editable: true, required: true },
    { key: 'moneda', label: 'Moneda', type: 'text', editable: true, required: true },
    { key: 'devolucion_capital', label: 'Dev. capital', type: 'money', editable: true, required: false },
    { key: 'interes_depositar', label: 'Interés', type: 'money', editable: true, required: true },
    { key: 'mora', label: 'Mora', type: 'money', editable: true, required: false },
    { key: 'retencion', label: 'Retención', type: 'money', editable: true, required: false },
    { key: 'total_depositar', label: 'Total a depositar', type: 'money', editable: true, required: true },
  ];

  // Inputs para agregar fila
  const ADD_ROW_FIELDS = [
    { key: 'codigo', label: 'Código', placeholder: 'Ej: OP-12345' },
    { key: 'inversionista', label: 'Inversionista', placeholder: 'Ej: Juan Pérez' },
    { key: 'fechas_pago', label: 'Fechas de pago', placeholder: 'Ej: 15/01/2026' },
    { key: 'moneda', label: 'Moneda (PEN/USD)', placeholder: 'PEN' },
    { key: 'devolucion_capital', label: 'Dev. capital', placeholder: '0.00' },
    { key: 'interes_depositar', label: 'Interés', placeholder: '125.50' },
    { key: 'mora', label: 'Mora', placeholder: '0.00' },
    { key: 'retencion', label: 'Retención', placeholder: '0.00' },
    { key: 'total_depositar', label: 'Total', placeholder: '125.50' },
  ];

  // ====== Init ======
  init();

  function init() {
    bootMonthYear();
    buildAddRowInputs();
    loadLogo();

    el.fileInput.addEventListener('change', onFileSelected);
    el.companySelect.addEventListener('change', () => {
      STATE.selectedCompany = el.companySelect.value || '';
      STATE.page = 1;
      buildReportForSelectedCompany();
    });
    el.monthSelect.addEventListener('change', () => updateEmailPreview());
    el.yearSelect.addEventListener('change', () => updateEmailPreview());

    el.exportExcelBtn.addEventListener('click', exportExcel);
    el.exportPdfBtn.addEventListener('click', exportPdf);
    el.copyEmailBtn.addEventListener('click', copyEmail);

    el.prevPageBtn.addEventListener('click', () => {
      if (STATE.page > 1) {
        STATE.page--;
        renderTable();
      }
    });
    el.nextPageBtn.addEventListener('click', () => {
      const totalPages = Math.max(1, Math.ceil(STATE.reportRows.length / STATE.pageSize));
      if (STATE.page < totalPages) {
        STATE.page++;
        renderTable();
      }
    });

    el.addRowBtn.addEventListener('click', onAddRow);
    setStatus('Sin archivo', 'neutral');
  }

  function setStatus(text, kind) {
    el.statusBadge.textContent = text;
    el.statusBadge.style.borderColor =
      kind === 'good' ? 'rgba(34,197,94,.35)' :
      kind === 'warn' ? 'rgba(245,158,11,.35)' :
      kind === 'bad'  ? 'rgba(239,68,68,.35)' : 'rgba(255,255,255,.10)';
    el.statusBadge.style.background =
      kind === 'good' ? 'rgba(34,197,94,.08)' :
      kind === 'warn' ? 'rgba(245,158,11,.08)' :
      kind === 'bad'  ? 'rgba(239,68,68,.08)' : 'rgba(255,255,255,.04)';
    el.statusBadge.style.color =
      kind === 'good' ? '#bff3cf' :
      kind === 'warn' ? '#ffe3b0' :
      kind === 'bad'  ? '#fecaca' : 'var(--muted)';
  }

  function showAlert(msg) {
    el.alertBox.textContent = msg;
    el.alertBox.classList.remove('hidden');
  }
  function hideAlert() {
    el.alertBox.textContent = '';
    el.alertBox.classList.add('hidden');
  }
  function showWarn(msg) {
    el.warnBox.textContent = msg;
    el.warnBox.classList.remove('hidden');
  }
  function hideWarn() {
    el.warnBox.textContent = '';
    el.warnBox.classList.add('hidden');
  }

  async function loadLogo() {
    // 1) En pantalla: SIEMPRE cargar desde URL (esto NO requiere CORS)
  el.logoImg.src = LOGO_URL;
  el.emailLogo.src = LOGO_URL;

  // 2) Para el PDF: intentamos convertir a DataURL (puede fallar por CORS en file://)
  try {
    const res = await fetch(LOGO_URL, { mode: 'cors', cache: 'no-store' });
    if (!res.ok) throw new Error('Logo fetch failed');
    const blob = await res.blob();
    const dataUrl = await blobToDataUrl(blob);
    STATE.logoDataUrl = dataUrl;
  } catch (e) {
    STATE.logoFailed = true;
    // Solo avisamos: la UI ya lo mostrará igual por URL
    showWarn('Advertencia: el navegador bloqueó la descarga del logo (CORS) para incrustarlo en PDF. En pantalla sí se verá; el PDF se generará sin logo.');
  }

  // 3) Si el URL fallara (raro), ocultamos el ícono roto
  el.logoImg.onerror = () => { el.logoImg.style.display = 'none'; };
  el.emailLogo.onerror = () => { el.emailLogo.style.display = 'none'; };
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result));
      r.readAsDataURL(blob);
    });
  }

  function bootMonthYear() {
    const monthNames = [
      'Enero','Febrero','Marzo','Abril','Mayo','Junio',
      'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
    ];
    el.monthSelect.innerHTML = monthNames.map((m, i) => `<option value="${i+1}">${m}</option>`).join('');

    const now = new Date();
    const y = now.getFullYear();
    const years = [];
    for (let i = y - 2; i <= y + 3; i++) years.push(i);
    el.yearSelect.innerHTML = years.map(v => `<option value="${v}">${v}</option>`).join('');

    el.monthSelect.value = String(now.getMonth() + 1);
    el.yearSelect.value = String(y);

    el.monthSelect.disabled = true;
    el.yearSelect.disabled = true;
  }

  function enableControls(enabled) {
    el.companySelect.disabled = !enabled;
    el.monthSelect.disabled = !enabled;
    el.yearSelect.disabled = !enabled;

    el.exportExcelBtn.disabled = !enabled;
    el.exportPdfBtn.disabled = !enabled;
    el.copyEmailBtn.disabled = !enabled;

    el.addRowBtn.disabled = !enabled;
  }

  // ====== File -> Parse -> Normalize ======
  async function onFileSelected() {
    hideAlert();
    hideWarn();
    STATE.warnings = [];
    STATE.rawRows = [];
    STATE.normalizedRows = [];
    STATE.companies = [];
    STATE.selectedCompany = '';
    STATE.reportRows = [];
    STATE.page = 1;

    const file = el.fileInput.files?.[0];
    if (!file) return;

    el.fileName.textContent = file.name;
    setStatus('Leyendo Excel…', 'warn');
    enableControls(false);

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheetName = wb.SheetNames?.[0];
      if (!sheetName) throw new Error('El archivo no tiene hojas.');

      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });

      if (!rows.length) throw new Error('La hoja está vacía.');

      STATE.rawRows = rows;

      const { normalized, missingCritical, warnings } = normalizeAndParseRows(rows);
      STATE.normalizedRows = normalized;
      STATE.warnings = warnings;

      if (missingCritical.length) {
        setStatus('Columnas faltantes', 'bad');
        showAlert(`Faltan columnas críticas: ${missingCritical.join(', ')}`);
        return;
      }

      STATE.companies = getCompanies(normalized);
      if (!STATE.companies.length) {
        setStatus('Sin empresas', 'bad');
        showAlert('No se encontraron empresas en el archivo.');
        return;
      }

      fillCompanyDropdown(STATE.companies);

      setStatus('Listo', 'good');
      enableControls(true);

      // Auto-selección primera empresa
      el.companySelect.value = STATE.companies[0];
      STATE.selectedCompany = STATE.companies[0];
      buildReportForSelectedCompany();

      // Warnings
      if (STATE.warnings.length) {
        showWarn(`Advertencias detectadas: ${STATE.warnings.length}. Revisa “Moneda” y montos.`);
      }
    } catch (err) {
      setStatus('Error', 'bad');
      showAlert(String(err?.message || err));
    }
  }

  function normalizeAndParseRows(rows) {
    const warnings = [];
    const normalized = [];

    // Mapeo de cabeceras -> claves internas
    const requiredMap = {
      codigo: ['codigo', 'código', 'cod', 'code'],
      inversionista: ['inversionista', 'inversor', 'investor', 'cliente', 'titular'],
      dni: ['numero de documento beneficiario', 'número de documento beneficiario', 'nro de documento beneficiario', 'documento beneficiario', 'dni', 'documento', 'numero documento'],
      empresa: ['empresa', 'razon social', 'razón social', 'compania', 'compañía', 'company'],
      fecha_inicio: ['fecha inicio inversion', 'fecha inicio inversión', 'inicio inversion', 'inicio inversión', 'fecha inicio', 'start date'],
      fechas_pago: ['fechas de pago', 'fecha pago', 'fechas pago', 'pago', 'payment dates'],
      devolucion_capital: ['devolucion de capital', 'devolución de capital', 'capital', 'capital a devolver', 'principal repayment'],
      interes_depositar: ['interes a depositar', 'interés a depositar', 'interes', 'interés', 'interest to deposit'],
      mora: ['mora', 'late fee', 'penalidad'],
      retencion: ['retencion', 'retención', 'withholding', 'retention'],
      total_depositar: ['total a depositar', 'total depositar', 'total', 'monto total', 'total to deposit']
    };

    const critical = ['codigo', 'inversionista', 'empresa', 'fechas_pago', 'interes_depositar', 'total_depositar'];

    // Resolver columnas reales (según primera fila de keys existentes)
    const headerKeys = new Set();
    for (const r of rows) Object.keys(r || {}).forEach(k => headerKeys.add(k));
    const headerList = Array.from(headerKeys);

    const resolved = {};
    for (const [k, alts] of Object.entries(requiredMap)) {
      resolved[k] = findHeaderMatch(headerList, alts);
    }

    const missingCritical = critical
      .filter(k => !resolved[k])
      .map(k => {
        const pretty = REPORT_COLS.find(c => c.key === k)?.label || k;
        return pretty;
      });

    if (missingCritical.length) {
      return { normalized: [], missingCritical, warnings: [] };
    }

    // Parse rows
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || {};
      const empresa = safeText(r[resolved.empresa]);
      if (!empresa) continue;

      const codigo = safeText(r[resolved.codigo]);
      const inversionista = safeText(r[resolved.inversionista]);
      const dni = safeText(r[resolved.dni]);

      const fechasPago = formatAnyDate(r[resolved.fechas_pago]);
      const fechaInicio = resolved.fecha_inicio ? formatAnyDate(r[resolved.fecha_inicio]) : '';

      // Moneda desde "Interés a depositar"
      const interesRaw = r[resolved.interes_depositar];
      const moneda = detectCurrency(interesRaw);
      const interes = parseMoney(interesRaw).value;

      // Otros montos
      const devol = resolved.devolucion_capital ? parseMoney(r[resolved.devolucion_capital]).value : '';
      const mora = resolved.mora ? parseMoney(r[resolved.mora]).value : '';
      const ret = resolved.retencion ? parseMoney(r[resolved.retencion]).value : '';
      const total = parseMoney(r[resolved.total_depositar]).value;

      // Warnings de moneda
      if (moneda === 'PEN_DEFAULTED') {
        warnings.push(`Fila ${i + 2}: moneda no identificada en "Interés a depositar" (default PEN).`);
      }

      normalized.push({
        __idx: i,
        __warnCurrency: moneda === 'PEN_DEFAULTED',
        empresa,
        codigo,
        inversionista,
        dni,
        fecha_inicio: fechaInicio,
        fechas_pago: fechasPago,

        moneda: moneda === 'PEN_DEFAULTED' ? 'PEN' : moneda,

        devolucion_capital: toFixed2(devol),
        interes_depositar: toFixed2(interes),
        mora: toFixed2(mora),
        retencion: toFixed2(ret),
        total_depositar: toFixed2(total),
      });
    }

    return { normalized, missingCritical: [], warnings };
  }

  function findHeaderMatch(headers, alternatives) {
    const normHeaders = headers.map(h => ({ raw: h, norm: normHeader(h) }));
    const altNorms = alternatives.map(a => normHeader(a));

    for (const alt of altNorms) {
      const hit = normHeaders.find(h => h.norm === alt);
      if (hit) return hit.raw;
    }
    // fallback: includes
    for (const alt of altNorms) {
      const hit = normHeaders.find(h => h.norm.includes(alt) || alt.includes(h.norm));
      if (hit) return hit.raw;
    }
    return null;
  }

  function normHeader(s) {
    return String(s || '')
      .trim()
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function safeText(v) {
    const t = String(v ?? '').trim();
    return t;
  }

  function getCompanies(rows) {
    const set = new Set();
    for (const r of rows) {
      const e = safeText(r.empresa);
      if (e) set.add(e);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
  }

  function fillCompanyDropdown(companies) {
    el.companySelect.innerHTML = `<option value="">— Seleccionar —</option>` +
      companies.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  }

  // ====== Build Editable Report ======
  function buildReportForSelectedCompany() {
    hideAlert();

    if (!STATE.selectedCompany) {
      STATE.reportRows = [];
      renderTable();
      updateStats();
      updateEmailPreview();
      enableControls(false);
      return;
    }

    const filtered = STATE.normalizedRows.filter(r => r.empresa === STATE.selectedCompany);
    if (!filtered.length) {
      STATE.reportRows = [];
      setStatus('Sin filas', 'warn');
      showAlert(`No hay filas para la empresa: ${STATE.selectedCompany}`);
      renderTable();
      updateStats();
      updateEmailPreview();
      enableControls(true);
      return;
    }

    // Clonar para edición
    STATE.reportRows = filtered.map(r => ({
      codigo: r.codigo,
      inversionista: r.inversionista,
      fechas_pago: r.fechas_pago,
      moneda: r.moneda,
      devolucion_capital: r.devolucion_capital,
      interes_depositar: r.interes_depositar,
      mora: r.mora,
      retencion: r.retencion,
      total_depositar: r.total_depositar,
      __warnCurrency: !!r.__warnCurrency
    }));

    STATE.page = 1;
    renderTable();
    updateStats();
    updateEmailPreview();
    setStatus('Listo', 'good');
  }

  function renderTable() {
    // Head
    el.tableHead.innerHTML = `
      <tr>
        <th style="min-width:70px">Acción</th>
        ${REPORT_COLS.map(c => `<th>${escapeHtml(c.label)}</th>`).join('')}
      </tr>
    `;

    const totalRows = STATE.reportRows.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / STATE.pageSize));
    STATE.page = Math.min(STATE.page, totalPages);

    const start = (STATE.page - 1) * STATE.pageSize;
    const end = Math.min(start + STATE.pageSize, totalRows);
    const pageRows = STATE.reportRows.slice(start, end);

    el.pageInfo.textContent = totalRows ? `Página ${STATE.page} / ${totalPages} (filas ${start + 1}-${end})` : '—';

    el.prevPageBtn.disabled = !(STATE.page > 1);
    el.nextPageBtn.disabled = !(STATE.page < totalPages);

    // Body
    const rowsHtml = pageRows.map((row, idxOnPage) => {
      const globalIdx = start + idxOnPage;
      const warnPill = row.__warnCurrency ? `<span class="pill warn">Moneda default</span>` : `<span class="pill good">OK</span>`;

      return `
        <tr data-idx="${globalIdx}">
          <td>
            <div class="row-actions">
              <button class="btn btn-sm btn-del" type="button" data-action="delete">Eliminar</button>
              ${warnPill}
            </div>
          </td>
          ${REPORT_COLS.map(col => {
            const val = row[col.key] ?? '';
            const editable = col.editable ? 'contenteditable="true"' : '';
            const aria = `aria-label="${escapeHtml(col.label)}"`;
            const data = `data-col="${col.key}"`;
            return `
              <td>
                <div class="cell-edit" ${editable} ${aria} ${data}>${escapeHtml(String(val))}</div>
              </td>
            `;
          }).join('')}
        </tr>
      `;
    }).join('');

    el.tableBody.innerHTML = rowsHtml || `<tr><td colspan="${REPORT_COLS.length + 1}" class="muted" style="padding:14px">Sin datos</td></tr>`;

    // Delegación: delete + edit
    el.tableBody.querySelectorAll('.btn-del').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const tr = e.target.closest('tr');
        const idx = Number(tr?.getAttribute('data-idx'));
        if (!Number.isFinite(idx)) return;
        STATE.reportRows.splice(idx, 1);
        const totalPagesNew = Math.max(1, Math.ceil(STATE.reportRows.length / STATE.pageSize));
        if (STATE.page > totalPagesNew) STATE.page = totalPagesNew;
        renderTable();
        updateStats();
        updateEmailPreview();
      });
    });

    el.tableBody.querySelectorAll('.cell-edit[contenteditable="true"]').forEach(div => {
      div.addEventListener('blur', (e) => {
        const tr = e.target.closest('tr');
        const idx = Number(tr?.getAttribute('data-idx'));
        const col = e.target.getAttribute('data-col');
        if (!Number.isFinite(idx) || !col) return;

        let value = String(e.target.textContent ?? '').trim();

        // Normalizaciones por tipo
        const colDef = REPORT_COLS.find(c => c.key === col);
        if (colDef?.type === 'money') {
          value = toFixed2(parseMoney(value).value);
        } else if (colDef?.type === 'dateText') {
          value = formatAnyDate(value);
        } else if (col === 'moneda') {
          value = normalizeCurrencyText(value);
        }

        STATE.reportRows[idx][col] = value;

        // Re-evaluar advertencia moneda si editan interés
        if (col === 'interes_depositar' || col === 'moneda') {
          STATE.reportRows[idx].__warnCurrency = (normalizeCurrencyText(STATE.reportRows[idx].moneda) === 'PEN' && looksCurrencyUnknownHint(value));
        }

        // Re-render solo stats + email, sin re-render tabla completa (performance)
        updateStats();
        updateEmailPreview();
      });
    });
  }

  function buildAddRowInputs() {
    el.addRowGrid.innerHTML = '';
    for (const f of ADD_ROW_FIELDS) {
      const wrap = document.createElement('div');
      wrap.className = 'field';
      const label = document.createElement('label');
      label.textContent = f.label;

      const input = document.createElement('input');
      input.className = 'inp';
      input.placeholder = f.placeholder;
      input.setAttribute('data-add', f.key);

      wrap.appendChild(label);
      wrap.appendChild(input);
      el.addRowGrid.appendChild(wrap);
    }
  }

  function onAddRow() {
    if (!STATE.selectedCompany) return;

    const obj = {};
    for (const f of ADD_ROW_FIELDS) {
      const inp = el.addRowGrid.querySelector(`[data-add="${f.key}"]`);
      obj[f.key] = String(inp?.value ?? '').trim();
    }

    obj.moneda = normalizeCurrencyText(obj.moneda || 'PEN');

    // Money fields
    obj.devolucion_capital = toFixed2(parseMoney(obj.devolucion_capital).value);
    obj.interes_depositar = toFixed2(parseMoney(obj.interes_depositar).value);
    obj.mora = toFixed2(parseMoney(obj.mora).value);
    obj.retencion = toFixed2(parseMoney(obj.retencion).value);
    obj.total_depositar = toFixed2(parseMoney(obj.total_depositar).value);

    // Date
    obj.fechas_pago = formatAnyDate(obj.fechas_pago);

    obj.__warnCurrency = false;

    STATE.reportRows.unshift(obj);
    STATE.page = 1;
    renderTable();
    updateStats();
    updateEmailPreview();

    // limpiar inputs
    el.addRowGrid.querySelectorAll('.inp').forEach(i => i.value = '');
  }

  // ====== Exports ======
  function exportExcel() {
    if (!STATE.selectedCompany || !STATE.reportRows.length) return;

    const data = STATE.reportRows.map(r => ({
      'Código': r.codigo,
      'Inversionista': r.inversionista,
      'Fechas de pago': r.fechas_pago,
      'Moneda': r.moneda,
      'Dev. capital': r.devolucion_capital,
      'Interés': r.interes_depositar,
      'Mora': r.mora,
      'Retención': r.retencion,
      'Total a depositar': r.total_depositar
    }));

    const ws = XLSX.utils.json_to_sheet(data, { skipHeader: false });

    // Anchos simples
    ws['!cols'] = [
      { wch: 14 }, { wch: 28 }, { wch: 18 }, { wch: 8 },
      { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 16 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reporte');

    const { monthName, year } = getMonthYear();
    const safeCompany = sanitizeFilename(STATE.selectedCompany);
    const filename = `Reporte_Pagos_${safeCompany}_${monthName}_${year}.xlsx`;

    XLSX.writeFile(wb, filename);
  }

  function exportPdf() {
    if (!STATE.selectedCompany || !STATE.reportRows.length) return;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });

    const { monthName, year } = getMonthYear();
    const title = `Reporte de Pagos Mensuales – ${STATE.selectedCompany} – ${monthName} ${year}`;

    const marginX = 36;
    const marginTop = 56;
    const headerH = 44;

    const rows = STATE.reportRows.map(r => ([
      r.codigo,
      r.inversionista,
      r.fechas_pago,
      r.moneda,
      r.devolucion_capital,
      r.interes_depositar,
      r.mora,
      r.retencion,
      r.total_depositar
    ]));

    const head = [[
      'Código','Inversionista','Fechas de pago','Moneda',
      'Dev. capital','Interés','Mora','Retención','Total a depositar'
    ]];

    const drawHeaderFooter = (data) => {
      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();

      // Header bg line
      doc.setDrawColor(43, 108, 255);
      doc.setLineWidth(2);
      doc.line(marginX, 42, pageW - marginX, 42);

      // Logo
      if (STATE.logoDataUrl && !STATE.logoFailed) {
        try {
          doc.addImage(STATE.logoDataUrl, 'PNG', marginX, 14, 86, 28);
        } catch (_) { /* no-op */ }
      }

      // Title
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.text(title, marginX + 100, 34);

      // Footer
      const pageNumber = doc.internal.getNumberOfPages();
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(160);
      doc.text('Tasatop – Operaciones', marginX, pageH - 18);

      const label = `Página ${data.pageNumber} de ${pageNumber}`;
      doc.text(label, pageW - marginX, pageH - 18, { align: 'right' });
      doc.setTextColor(0);
    };

    doc.autoTable({
      head,
      body: rows,
      startY: marginTop + headerH,
      margin: { left: marginX, right: marginX, top: marginTop, bottom: 34 },
      styles: {
        font: 'helvetica',
        fontSize: 9,
        cellPadding: 6,
        overflow: 'linebreak',
        valign: 'top'
      },
      headStyles: {
        fillColor: [15, 32, 58],
        textColor: [232, 240, 255],
        fontStyle: 'bold'
      },
      alternateRowStyles: { fillColor: [245, 247, 255] },
      didDrawPage: (data) => drawHeaderFooter(data)
    });

    const safeCompany = sanitizeFilename(STATE.selectedCompany);
    doc.save(`Reporte_Pagos_${safeCompany}_${monthName}_${year}.pdf`);
  }

  async function copyEmail() {
    const text = el.emailText.value || '';
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      setStatus('Correo copiado', 'good');
      setTimeout(() => setStatus('Listo', 'good'), 900);
    } catch (e) {
      showAlert('No se pudo copiar. Selecciona el texto y copia manualmente.');
    }
  }

  // ====== Email ======
  function updateEmailPreview() {
    const empresa = STATE.selectedCompany || '{Empresa}';
    const { monthName, year } = getMonthYear();
    const resumen = getEmailSummary();

    const body =
`Buen día, equipo ${empresa},
Un gusto saludarlos. Les envío la información correspondiente a los abonos de ${monthName} ${year}.${resumen}

Por favor, su apoyo confirmando que se realizarán los pagos de acuerdo a lo programado y que los mismos sean subidos al sistema de Tasatop para que los inversionistas puedan visualizar los abonos. Asimismo, solicitamos subir los certificados de retenciones del mes anterior y remitir los PDTs de los meses pendientes.

Les recordamos que pagar en fecha evita la mora y las penalidades establecidas en los contratos. Es importante responder este correo confirmando que el pago se realizará en la fecha correspondiente; de lo contrario, se considerará como incumplimiento de los compromisos contractuales.

Saludos,
Operaciones – Tasatop`;

    el.emailText.value = body;
  }

  function getEmailSummary() {
    if (!STATE.reportRows.length) return '';
    const count = STATE.reportRows.length;

    // Total a depositar (suma numérica)
    let sum = 0;
    for (const r of STATE.reportRows) sum += parseMoney(r.total_depositar).value || 0;

    // Moneda “dominante”: si hay mezcla, mostrar MIX
    const set = new Set(STATE.reportRows.map(r => normalizeCurrencyText(r.moneda)));
    const moneda = set.size === 1 ? Array.from(set)[0] : 'MIX';

    const sumTxt = moneda === 'MIX' ? formatNumber(sum) : `${moneda} ${
    
    (sum)}`;
    return `\n\nResumen: ${count} registro(s) — Total a depositar: ${sumTxt}.`;
  }

  function getMonthYear() {
    const monthIdx = Number(el.monthSelect.value || 1);
    const year = Number(el.yearSelect.value || new Date().getFullYear());
    const monthNames = [
      'Enero','Febrero','Marzo','Abril','Mayo','Junio',
      'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
    ];
    return { monthIdx, monthName: monthNames[Math.max(0, monthIdx - 1)], year };
  }

  // ====== Stats ======
  function updateStats() {
    const n = STATE.reportRows.length;
    el.rowsCount.textContent = n ? String(n) : '—';

    let sum = 0;
    for (const r of STATE.reportRows) sum += parseMoney(r.total_depositar).value || 0;
    el.totalDeposit.textContent = n ? formatNumber(sum) : '—';

    const warns = STATE.reportRows.filter(r => r.__warnCurrency).length;
    el.warnCount.textContent = n ? String(warns) : '—';
  }

  // ====== Parsing Helpers ======
  function detectCurrency(v) {
    const s = String(v ?? '').toUpperCase();
    if (s.includes('S/') || s.includes('PEN') || s.includes('SOLES')) return 'PEN';
    if (s.includes('$') || s.includes('USD') || s.includes('DOLAR') || s.includes('DÓLAR')) return 'USD';

    // Si viene como número sin símbolo: default PEN pero advertir
    if (s.trim() === '' || s.trim() === '0' || s.trim() === '0.00') return 'PEN';
    return 'PEN_DEFAULTED';
  }

  function normalizeCurrencyText(v) {
    const s = String(v ?? '').trim().toUpperCase();
    if (s.includes('USD') || s.includes('$')) return 'USD';
    if (s.includes('PEN') || s.includes('S/')) return 'PEN';
    if (s === 'SOL' || s === 'SOLES') return 'PEN';
    return s === 'USD' ? 'USD' : 'PEN';
  }

  function looksCurrencyUnknownHint(v) {
    const s = String(v ?? '').toUpperCase();
    return !(s.includes('USD') || s.includes('$') || s.includes('PEN') || s.includes('S/'));
  }

  function parseMoney(v) {
    // Retorna { value:number } con 2 decimales, tolera "S/ 1,234.50", "1.234,50", "$ 50"
    if (typeof v === 'number' && Number.isFinite(v)) return { value: v };

    let s = String(v ?? '').trim();
    if (!s) return { value: 0 };

    // quitar moneda, espacios raros
    s = s.replace(/\s+/g, ' ');
    s = s.replace(/PEN|USD|S\/|\$/gi, '');
    s = s.trim();

    // dejar solo dígitos, , . y -
    s = s.replace(/[^\d,.\-]/g, '');

    if (!s || s === '-' || s === '.' || s === ',') return { value: 0 };

    // decidir separador decimal por el último . o ,
    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');

    let decimalSep = null;
    if (lastDot > lastComma) decimalSep = '.';
    else if (lastComma > lastDot) decimalSep = ',';

    let normalized = s;

    if (decimalSep) {
      const parts = s.split(decimalSep);
      const dec = parts.pop();
      const intPart = parts.join(decimalSep);
      normalized = intPart.replace(/[.,]/g, '') + '.' + dec;
    } else {
      normalized = s.replace(/[.,]/g, '');
    }

    const num = Number(normalized);
    return { value: Number.isFinite(num) ? num : 0 };
  }

  function toFixed2(v) {
 const num = typeof v === 'number' ? v : Number(v);
  if (!Number.isFinite(num)) return '';
  const rounded = Math.round(num * 100) / 100;

  // Formato visual con separador de miles y 2 decimales:
  // 12000.00 -> "12,000.00"
  // 1230.56  -> "1,230.56"
  return rounded.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  }

  function formatNumber(n) {
    const num = Number(n);
    if (!Number.isFinite(num)) return '0.00';
    return num.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatAnyDate(v) {
    if (v === null || v === undefined) return '';

    // XLSX puede entregar Date si cellDates=true
    if (v instanceof Date && !isNaN(v.getTime())) {
      return formatDateDMY(v);
    }

    // Excel serial number
    if (typeof v === 'number' && Number.isFinite(v)) {
      const d = excelDateToJSDate(v);
      return d ? formatDateDMY(d) : '';
    }

    const s = String(v).trim();
    if (!s) return '';

    // yyyy-mm-dd
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return isNaN(d.getTime()) ? s : formatDateDMY(d);
    }

    // dd/mm/yyyy o dd-mm-yyyy
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      let yy = Number(m[3]);
      if (yy < 100) yy += 2000;
      const d = new Date(yy, mm - 1, dd);
      return isNaN(d.getTime()) ? s : formatDateDMY(d);
    }

    // fallback: devolver texto tal cual
    return s;
  }

  function excelDateToJSDate(serial) {
    // Usa XLSX si está disponible
    try {
      if (window.XLSX?.SSF?.parse_date_code) {
        const o = window.XLSX.SSF.parse_date_code(serial);
        if (!o) return null;
        return new Date(o.y, o.m - 1, o.d);
      }
    } catch (_) { /* no-op */ }

    // Fallback aproximado
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const d = new Date(utcValue * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  function formatDateDMY(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yy = d.getFullYear();
    return `${dd}/${mm}/${yy}`;
  }

  function escapeHtml(str) {
    return String(str ?? '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  function sanitizeFilename(s) {
    return String(s ?? '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^\w\s-]/g, '')
      .trim()
      .replace(/\s+/g, '_')
      .slice(0, 80) || 'Empresa';
  }
})();
