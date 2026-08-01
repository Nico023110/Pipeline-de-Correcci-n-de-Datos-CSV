// Initialize Lucide Icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// App State
const state = {
    uploadedFiles: [],
    rawRows: [],
    cleanedRows: [],
    errorLog: [],
    pagination: {
        all: { page: 1, records: [] },
        corregidos: { page: 1, records: [] },
        log: { page: 1, records: [] },
        pageSize: 50
    }
};

function initApp() {
    // Navigation Tabs
    const navButtons = document.querySelectorAll('.nav-btn');
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.getAttribute('data-tab');
            navButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.add('hidden');
            });
            document.getElementById(`tab-${targetTab}`).classList.remove('hidden');
        });
    });

    // File Selection
    const fileInput = document.getElementById('file-input');
    fileInput.addEventListener('change', (e) => handleFilesSelect(Array.from(e.target.files)));

    const dropzone = document.getElementById('main-dropzone');
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
            handleFilesSelect(Array.from(e.dataTransfer.files));
        }
    });

    // Buttons
    document.getElementById('btn-ejecutar').addEventListener('click', runCorrectionPipeline);
    document.getElementById('btn-load-sample').addEventListener('click', loadSampleData);

    // Search Box Listeners
    setupSearch('search-pipeline', 'all');
    setupSearch('search-corregidos', 'corregidos');
    setupSearch('search-log', 'log');

    // Pagination Listeners
    setupPaginationControls('all', 'btn-prev-page', 'btn-next-page');
    setupPaginationControls('corregidos', 'btn-prev-corregidos', 'btn-next-corregidos');
    setupPaginationControls('log', 'btn-prev-log', 'btn-next-log');

    // Export Buttons
    document.getElementById('btn-export-csv').addEventListener('click', exportCleanedCSV);
    document.getElementById('btn-export-log-excel').addEventListener('click', exportLogExcel);
}

// Handle File Selection & Parsing
async function handleFilesSelect(files) {
    if (!files || !files.length) return;

    showLoader(`Leyendo y procesando ${files.length} archivo(s)...`);
    state.uploadedFiles = files;
    state.rawRows = [];

    const fileTagsList = document.getElementById('file-tags-list');
    fileTagsList.innerHTML = '';

    for (let file of files) {
        const span = document.createElement('span');
        span.className = 'file-tag';
        span.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        fileTagsList.appendChild(span);

        try {
            const rows = await parseFile(file);
            if (rows && rows.length) {
                state.rawRows = state.rawRows.concat(rows);
            }
        } catch (err) {
            console.warn(`Error al leer ${file.name}:`, err);
        }
    }

    document.getElementById('btn-ejecutar').disabled = state.rawRows.length === 0;
    hideLoader();
}

// File Parsing Supporting Excel & Delimited Text
async function parseFile(file) {
    const ext = file.name.toLowerCase();

    if (ext.endsWith('.xlsx')) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.worksheets[0];
            const rows = [];
            const headers = [];
            let headerParsed = false;

            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                const values = row.values;
                if (!headerParsed) {
                    for (let i = 1; i < values.length; i++) {
                        headers.push(values[i] ? values[i].toString().trim() : `Col_${i}`);
                    }
                    headerParsed = true;
                } else {
                    const rowObj = {};
                    for (let i = 1; i <= headers.length; i++) {
                        const header = headers[i - 1];
                        let val = values[i];
                        rowObj[header] = val !== null && val !== undefined ? val.toString().trim() : '';
                    }
                    rows.push(rowObj);
                }
            });
            return rows;
        } catch (err) {
            return parseCSVText(await file.text());
        }
    } else {
        const text = await file.text();
        return parseCSVText(text);
    }
}

// Multi-Delimiter CSV/TSV Parser
function parseCSVText(text) {
    if (!text || !text.trim()) return [];

    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];

    const firstLine = lines[0];
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const semiCount = (firstLine.match(/;/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const pipeCount = (firstLine.match(/\|/g) || []).length;

    let delimiter = ',';
    const maxCount = Math.max(tabCount, semiCount, commaCount, pipeCount);
    if (maxCount > 0) {
        if (maxCount === tabCount) delimiter = '\t';
        else if (maxCount === semiCount) delimiter = ';';
        else if (maxCount === pipeCount) delimiter = '|';
        else delimiter = ',';
    }

    const headers = firstLine.split(delimiter).map(h => h.replace(/^["']|["']$/g, '').trim());
    const result = [];

    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(delimiter);
        const row = {};
        headers.forEach((header, index) => {
            let val = values[index] !== undefined ? values[index] : '';
            row[header] = val.replace(/^["']|["']$/g, '').trim();
        });
        result.push(row);
    }
    return result;
}

// Correction Engine: Runs All Standard RIPS Correction Rules Automatically
function runCorrectionPipeline() {
    if (!state.rawRows.length) return;

    showLoader('Ejecutando Pipeline de Limpieza & Corrección RIPS...');
    setTimeout(() => {
        state.cleanedRows = [];
        state.errorLog = [];

        state.rawRows.forEach((rawRow, idx) => {
            const rowNum = idx + 1;
            const cleanedRow = { ...rawRow };
            let hasCorrection = false;

            const getVal = (keys) => {
                for (let k of keys) {
                    for (let key in rawRow) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return { key, val: String(rawRow[key]).trim() };
                    }
                }
                return { key: '', val: '' };
            };

            // Document Type & Document Number Normalization
            const docTypeObj = getVal(['tipo_documento', 'tipodoc', 'tipo_doc', 'td']);
            if (docTypeObj.val) {
                const normDocType = docTypeObj.val.toUpperCase().replace(/[^A-Z]/g, '');
                const validTypes = ['CC', 'TI', 'CE', 'PA', 'NV', 'CD', 'MS', 'RC'];
                const finalType = validTypes.includes(normDocType) ? normDocType : 'CC';
                if (docTypeObj.val !== finalType) {
                    cleanedRow[docTypeObj.key] = finalType;
                    hasCorrection = true;
                    state.errorLog.push({
                        fila: rowNum,
                        campo: docTypeObj.key,
                        original: docTypeObj.val,
                        corregido: finalType,
                        regla: 'Normalización de Tipo Documento RIPS'
                    });
                }
            }

            const docNumObj = getVal(['num_documento', 'documento', 'cedula', 'num_doc']);
            if (docNumObj.val) {
                const normDocNum = docNumObj.val.replace(/[^0-9]/g, '');
                if (docNumObj.val !== normDocNum && normDocNum.length > 0) {
                    cleanedRow[docNumObj.key] = normDocNum;
                    hasCorrection = true;
                    state.errorLog.push({
                        fila: rowNum,
                        campo: docNumObj.key,
                        original: docNumObj.val,
                        corregido: normDocNum,
                        regla: 'Remoción de caracteres especiales en Documento'
                    });
                }
            }

            // Date Standardization
            const fechaObj = getVal(['fecha_atencion', 'fecha', 'fechainicioatencion', 'fecha_servicio']);
            if (fechaObj.val) {
                const stdDate = standardizeDate(fechaObj.val);
                if (stdDate && fechaObj.val !== stdDate) {
                    cleanedRow[fechaObj.key] = stdDate;
                    hasCorrection = true;
                    state.errorLog.push({
                        fila: rowNum,
                        campo: fechaObj.key,
                        original: fechaObj.val,
                        corregido: stdDate,
                        regla: 'Estandarización de Fecha a formato ISO YYYY-MM-DD'
                    });
                }
            }

            // CUPS Sanitization
            const cupsObj = getVal(['cod_cups', 'cups', 'cod_procedimiento', 'actividad']);
            if (cupsObj.val) {
                const cleanCups = cupsObj.val.replace(/[^0-9A-Z]/gi, '').toUpperCase();
                if (cupsObj.val !== cleanCups) {
                    cleanedRow[cupsObj.key] = cleanCups;
                    hasCorrection = true;
                    state.errorLog.push({
                        fila: rowNum,
                        campo: cupsObj.key,
                        original: cupsObj.val,
                        corregido: cleanCups,
                        regla: 'Limpieza de código CUPS'
                    });
                }
            }

            // CIE-10 Normalization
            const cieObj = getVal(['cod_diagnostico', 'cie10', 'diagnostico_principal', 'diag_principal']);
            if (cieObj.val) {
                const cleanCie = cieObj.val.replace(/[^0-9A-Z]/gi, '').toUpperCase();
                if (cieObj.val !== cleanCie) {
                    cleanedRow[cieObj.key] = cleanCie;
                    hasCorrection = true;
                    state.errorLog.push({
                        fila: rowNum,
                        campo: cieObj.key,
                        original: cieObj.val,
                        corregido: cleanCie,
                        regla: 'Normalización de Diagnóstico CIE-10'
                    });
                }
            }

            cleanedRow._hasCorrection = hasCorrection;
            cleanedRow._id = rowNum;
            state.cleanedRows.push(cleanedRow);
        });

        updateUI();
        hideLoader();
    }, 600);
}

// Date Parser: Converts DD/MM/YYYY or YYYY/MM/DD to YYYY-MM-DD
function standardizeDate(str) {
    if (!str) return '';
    const clean = str.trim().replace(/\//g, '-');
    const parts = clean.split('-');

    if (parts.length === 3) {
        if (parts[0].length === 4) {
            return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
        } else if (parts[2].length === 4) {
            return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
        }
    }
    return str;
}

// Update UI & KPI Cards
function updateUI() {
    const total = state.cleanedRows.length;
    const corregidos = state.errorLog.length;
    const limpios = total - state.cleanedRows.filter(r => !r._hasCorrection).length;
    const tasa = total ? ((1 - (corregidos / (total * 4))) * 100).toFixed(1) : 100;

    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-corregidos').textContent = limpios.toLocaleString();
    document.getElementById('kpi-errores').textContent = corregidos.toLocaleString();
    document.getElementById('kpi-tasa').textContent = `${tasa}%`;

    document.getElementById('badge-count').textContent = `${total.toLocaleString()} registros`;
    document.getElementById('badge-corregidos').textContent = `${limpios.toLocaleString()} registros`;
    document.getElementById('badge-log').textContent = `${corregidos.toLocaleString()} correcciones`;

    document.getElementById('btn-export-csv').disabled = total === 0;
    document.getElementById('btn-export-log-excel').disabled = corregidos === 0;

    state.pagination.all.records = state.cleanedRows; state.pagination.all.page = 1;
    state.pagination.corregidos.records = state.cleanedRows.filter(r => r._hasCorrection); state.pagination.corregidos.page = 1;
    state.pagination.log.records = state.errorLog; state.pagination.log.page = 1;

    renderTabTable('all');
    renderTabTable('corregidos');
    renderTabTable('log');
}

// Render Tab Table
function renderTabTable(tabKey) {
    const tbodyId = tabKey === 'all' ? 'table-body' : `table-body-${tabKey}`;
    const pageInfoId = tabKey === 'all' ? 'page-info' : `page-info-${tabKey}`;
    const pageNumId = tabKey === 'all' ? 'current-page-num' : `page-num-${tabKey}`;
    const prevBtnId = tabKey === 'all' ? 'btn-prev-page' : `btn-prev-${tabKey}`;
    const nextBtnId = tabKey === 'all' ? 'btn-next-page' : `btn-next-${tabKey}`;

    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    const tabState = state.pagination[tabKey];
    const records = tabState.records;
    const pageSize = state.pagination.pageSize;
    const totalPages = Math.max(1, Math.ceil(records.length / pageSize));
    const currentPage = Math.min(tabState.page, totalPages);
    tabState.page = currentPage;

    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, records.length);
    const pageRecords = records.slice(startIdx, endIdx);

    if (!records.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${tabKey === 'log' ? 5 : 8}" class="empty-table-msg">
                    <i data-lucide="info"></i>
                    <p>No hay registros en esta vista.</p>
                </td>
            </tr>
        `;
        document.getElementById(pageInfoId).textContent = 'Mostrando 0 registros';
        document.getElementById(prevBtnId).disabled = true;
        document.getElementById(nextBtnId).disabled = true;
        lucide.createIcons();
        return;
    }

    if (tabKey === 'log') {
        pageRecords.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><code>#${row.fila}</code></td>
                <td><strong>${row.campo}</strong></td>
                <td><span style="color: #F87171; text-decoration: line-through">${row.original}</span></td>
                <td><span style="color: #34D399; font-weight: 600">${row.corregido}</span></td>
                <td><span class="badge" style="background: rgba(245, 158, 11, 0.15); color: #FBBF24">${row.regla}</span></td>
            `;
            tbody.appendChild(tr);
        });
    } else {
        pageRecords.forEach(row => {
            const tr = document.createElement('tr');
            const getField = (keys) => {
                for (let k of keys) {
                    for (let key in row) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return row[key];
                    }
                }
                return '-';
            };

            const docType = getField(['tipo_documento', 'tipodoc', 'td']) || 'CC';
            const docNum = getField(['num_documento', 'documento', 'cedula']);
            const nombre = getField(['nombre_afiliado', 'nombre', 'paciente']) || `AFILIADO #${row._id}`;
            const cups = getField(['cod_cups', 'cups', 'actividad']) || '890201';
            const cie10 = getField(['cod_diagnostico', 'cie10', 'diagnostico_principal']) || 'I10X';
            const fecha = getField(['fecha_atencion', 'fecha', 'fecha_servicio']) || '2026-07-15';

            const statusBadge = row._hasCorrection 
                ? '<span class="badge badge-success">Corregido & Limpio</span>'
                : '<span class="badge" style="background: rgba(99,102,241,0.15); color: #818CF8">Original Sin Errores</span>';

            tr.innerHTML = `
                <td><code>#${row._id}</code></td>
                <td><span class="badge" style="background: rgba(255,255,255,0.08); color: #FFF">${docType}</span></td>
                <td><code>${docNum}</code></td>
                <td><strong>${nombre}</strong></td>
                <td><code>${cups}</code></td>
                <td><code>${cie10}</code></td>
                <td>${fecha}</td>
                <td>${statusBadge}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    document.getElementById(pageInfoId).textContent = `Mostrando ${(startIdx + 1).toLocaleString()} a ${endIdx.toLocaleString()} de ${records.length.toLocaleString()} registros`;
    document.getElementById(pageNumId).textContent = `Página ${currentPage} de ${totalPages}`;
    document.getElementById(prevBtnId).disabled = currentPage === 1;
    document.getElementById(nextBtnId).disabled = currentPage === totalPages;
}

function setupPaginationControls(tabKey, prevBtnId, nextBtnId) {
    document.getElementById(prevBtnId).addEventListener('click', () => {
        if (state.pagination[tabKey].page > 1) {
            state.pagination[tabKey].page--;
            renderTabTable(tabKey);
        }
    });

    document.getElementById(nextBtnId).addEventListener('click', () => {
        const totalPages = Math.ceil(state.pagination[tabKey].records.length / state.pagination.pageSize);
        if (state.pagination[tabKey].page < totalPages) {
            state.pagination[tabKey].page++;
            renderTabTable(tabKey);
        }
    });
}

function setupSearch(inputId, tabKey) {
    document.getElementById(inputId).addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const source = tabKey === 'all' ? state.cleanedRows : 
                       tabKey === 'corregidos' ? state.cleanedRows.filter(r => r._hasCorrection) : state.errorLog;

        if (tabKey === 'log') {
            state.pagination.log.records = source.filter(item => 
                String(item.fila).includes(q) ||
                item.campo.toLowerCase().includes(q) ||
                item.original.toLowerCase().includes(q) ||
                item.corregido.toLowerCase().includes(q) ||
                item.regla.toLowerCase().includes(q)
            );
        } else {
            state.pagination[tabKey].records = source.filter(row => JSON.stringify(row).toLowerCase().includes(q));
        }

        state.pagination[tabKey].page = 1;
        renderTabTable(tabKey);
    });
}

// Demo Sample Data Loader
function loadSampleData() {
    showLoader('Cargando dataset de prueba de facturación RIPS...');
    setTimeout(() => {
        const sampleRows = [];
        const docTypesRaw = ['cc.', 'T.I', 'C.E.', 'CC', 'ti', 'PA'];
        const cupsRaw = ['890.201', '890201', '890-202', ' 890301 ', '890201'];
        const cie10Raw = ['i10x', 'I10.X', 'e119', 'E11.9', 'j00x'];
        const fechasRaw = ['15/07/2026', '2026/07/16', '17-07-2026', '2026-07-18', '19/07/2026'];

        for (let i = 1; i <= 2500; i++) {
            const isFlawed = i % 2 === 0;
            sampleRows.push({
                tipo_documento: isFlawed ? docTypesRaw[i % docTypesRaw.length] : 'CC',
                num_documento: isFlawed ? `1.144.0${i + 100}-A` : `11440${i + 100}`,
                nombre_afiliado: `PACIENTE RIPS DEMO #${i}`,
                cod_cups: isFlawed ? cupsRaw[i % cupsRaw.length] : '890201',
                cod_diagnostico: isFlawed ? cie10Raw[i % cie10Raw.length] : 'I10X',
                fecha_atencion: isFlawed ? fechasRaw[i % fechasRaw.length] : '2026-07-15'
            });
        }

        state.rawRows = sampleRows;
        document.getElementById('btn-ejecutar').disabled = false;
        runCorrectionPipeline();
    }, 600);
}

// Export Cleaned CSV
function exportCleanedCSV() {
    if (!state.cleanedRows.length) return;

    const exportRows = state.cleanedRows.map(row => {
        const clean = { ...row };
        delete clean._id;
        delete clean._hasCorrection;
        return clean;
    });

    const headers = Object.keys(exportRows[0]);
    let csvContent = headers.join(';') + '\n';

    exportRows.forEach(row => {
        const line = headers.map(h => `"${row[h] || ''}"`).join(';');
        csvContent += line + '\n';
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Facturacion_RIPS_Corregida_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
}

// Export Audit Error Log to Excel
async function exportLogExcel() {
    if (!state.errorLog.length) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Log Inconsistencias RIPS');

    sheet.columns = [
        { header: 'FILA #', key: 'fila', width: 12 },
        { header: 'CAMPO AFECTADO', key: 'campo', width: 25 },
        { header: 'VALOR ORIGINAL (CRUDO)', key: 'original', width: 30 },
        { header: 'VALOR CORREGIDO', key: 'corregido', width: 30 },
        { header: 'REGLA APLICADA', key: 'regla', width: 45 }
    ];

    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '10B981' }
        };
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
    });

    state.errorLog.forEach(item => sheet.addRow(item));

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Log_Inconsistencias_RIPS_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
