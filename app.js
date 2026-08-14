const COLUMNS = ['料號', '品名', '分類', '供應商', '倉庫', '現有庫存', '安全庫存', '單價', '狀態', '更新日期'];
const STORAGE_KEY = 'scm-control-room-records';
const LANGUAGE_KEY = 'scm-control-room-language';
const COLUMN_TRANSLATIONS = { '料號': 'partNumber', '品名': 'itemName', '分類': 'category', '供應商': 'supplier', '倉庫': 'warehouse', '現有庫存': 'currentStock', '安全庫存': 'safetyStock', '單價': 'unitPrice', '狀態': 'status', '更新日期': 'updatedDate' };
const translations = {
  'zh-Hant': {
    subtitle: '供應鏈庫存與採購風險監控', language: '語言', downloadTemplate: '下載 Excel 範本', importExcel: '匯入 Excel', exportData: '匯出資料', inventorySummary: '庫存摘要', managedItems: '管理品項', currentDataset: '目前資料集', inventoryValue: '庫存價值', stockTimesPrice: '現有庫存 x 單價', lowStockRisk: '低庫存風險', belowSafetyStock: '低於安全庫存', supplierCount: '供應商數', activeSources: '啟用供應來源', filters: '資料篩選', clear: '清除', keyword: '關鍵字', searchPlaceholder: '料號、品名、供應商', warehouse: '倉庫', allWarehouses: '全部倉庫', status: '狀態', allStatuses: '全部狀態', normal: '正常', watch: '注意', shortage: '缺料', riskOnly: '僅顯示低庫存風險', riskNote: '紅色標記代表庫存低於安全水位', inventoryList: '庫存清單', addItem: '新增品項', partNumber: '料號', itemName: '品名', category: '分類', supplier: '供應商', currentSafety: '現有 / 安全', currentStock: '現有庫存', safetyStock: '安全庫存', unitPrice: '單價', updatedDate: '更新日期', actions: '操作', emptyState: '尚無符合條件的資料。請匯入 Excel 或新增品項。', close: '關閉', cancel: '取消', save: '儲存資料', edit: '編輯', newItem: '新增品項', editItem: '編輯 {item}', importSuccess: '已匯入 {count} 筆資料，並儲存於此瀏覽器。', importFailed: '匯入失敗：{message}', missingColumns: '缺少欄位：{columns}', noValidRows: '找不到可匯入的有效資料列。', readFailed: '無法讀取此檔案。', xlsxUnavailable: 'Excel 功能載入失敗，請確認網路連線後重新整理頁面。', noExportRows: '目前沒有可匯出的資料。', added: '已新增品項。', updated: '已更新品項。', templateFile: 'SCM_Excel_範本.xlsx', exportFile: 'SCM_庫存資料.xlsx', sheetName: '庫存資料'
  },
  en: {
    subtitle: 'Supply chain inventory and procurement risk monitoring', language: 'Language', downloadTemplate: 'Download Excel template', importExcel: 'Import Excel', exportData: 'Export data', inventorySummary: 'Inventory summary', managedItems: 'Managed items', currentDataset: 'Current dataset', inventoryValue: 'Inventory value', stockTimesPrice: 'Current stock x unit price', lowStockRisk: 'Low-stock risk', belowSafetyStock: 'Below safety stock', supplierCount: 'Suppliers', activeSources: 'Active sources', filters: 'Filters', clear: 'Clear', keyword: 'Keyword', searchPlaceholder: 'Part number, item, supplier', warehouse: 'Warehouse', allWarehouses: 'All warehouses', status: 'Status', allStatuses: 'All statuses', normal: 'Normal', watch: 'Watch', shortage: 'Shortage', riskOnly: 'Low-stock risk only', riskNote: 'Red rows are below the safety stock level', inventoryList: 'Inventory list', addItem: 'Add item', partNumber: 'Part number', itemName: 'Item name', category: 'Category', supplier: 'Supplier', currentSafety: 'Current / safety', currentStock: 'Current stock', safetyStock: 'Safety stock', unitPrice: 'Unit price', updatedDate: 'Updated date', actions: 'Actions', emptyState: 'No records match the current filters. Import Excel or add an item.', close: 'Close', cancel: 'Cancel', save: 'Save record', edit: 'Edit', newItem: 'Add item', editItem: 'Edit {item}', importSuccess: 'Imported {count} records and saved them in this browser.', importFailed: 'Import failed: {message}', missingColumns: 'Missing columns: {columns}', noValidRows: 'No valid records were found to import.', readFailed: 'This file could not be read.', xlsxUnavailable: 'Excel features could not load. Check the network connection and refresh the page.', noExportRows: 'There are no records to export.', added: 'Item added.', updated: 'Item updated.', templateFile: 'SCM_Excel_Template.xlsx', exportFile: 'SCM_Inventory.xlsx', sheetName: 'Inventory'
  }
};

const demoRecords = [
  { '料號': 'RM-10021', '品名': '鋁合金外殼', '分類': '原物料', '供應商': '東亞金屬', '倉庫': '桃園 A', '現有庫存': 280, '安全庫存': 200, '單價': 145, '狀態': '正常', '更新日期': '2026-08-14' },
  { '料號': 'RM-10048', '品名': '主機板模組', '分類': '電子料', '供應商': '聯創科技', '倉庫': '新竹 B', '現有庫存': 64, '安全庫存': 120, '單價': 980, '狀態': '缺料', '更新日期': '2026-08-14' },
  { '料號': 'PK-20017', '品名': '防靜電包材', '分類': '包裝材', '供應商': '宏宇包裝', '倉庫': '桃園 A', '現有庫存': 850, '安全庫存': 600, '單價': 8.5, '狀態': '正常', '更新日期': '2026-08-13' },
  { '料號': 'RM-10064', '品名': '散熱片組', '分類': '原物料', '供應商': '精鑄工業', '倉庫': '台中 C', '現有庫存': 110, '安全庫存': 150, '單價': 76, '狀態': '注意', '更新日期': '2026-08-12' },
  { '料號': 'CP-30009', '品名': '控制器成品', '分類': '半成品', '供應商': '內製', '倉庫': '新竹 B', '現有庫存': 310, '安全庫存': 250, '單價': 1680, '狀態': '正常', '更新日期': '2026-08-14' }
];

const elements = {
  body: document.querySelector('#inventoryTableBody'),
  empty: document.querySelector('#emptyState'),
  search: document.querySelector('#searchInput'),
  warehouse: document.querySelector('#warehouseFilter'),
  status: document.querySelector('#statusFilter'),
  riskOnly: document.querySelector('#riskOnlyFilter'),
  notice: document.querySelector('#notice'),
  dialog: document.querySelector('#recordDialog'),
  form: document.querySelector('#recordForm'),
  dialogTitle: document.querySelector('#dialogTitle'),
  itemCount: document.querySelector('#itemCount'),
  inventoryValue: document.querySelector('#inventoryValue'),
  riskCount: document.querySelector('#riskCount'),
  supplierCount: document.querySelector('#supplierCount')
};

let records = loadRecords();
let editingIndex = null;
let language = localStorage.getItem(LANGUAGE_KEY) === 'en' ? 'en' : 'zh-Hant';

function t(key, variables = {}) {
  return Object.entries(variables).reduce((text, [name, value]) => text.replaceAll(`{${name}}`, value), translations[language][key] ?? key);
}

function columnLabel(column) {
  return t(COLUMN_TRANSLATIONS[column]);
}

function statusLabel(status) {
  return t(status === '缺料' ? 'shortage' : status === '注意' ? 'watch' : 'normal');
}

function updateLanguage() {
  document.documentElement.lang = language;
  document.title = language === 'en' ? 'SCM Control Room' : 'SCM 控制中心';
  document.querySelector('#languageSelect').value = language;
  document.querySelectorAll('[data-i18n]').forEach((element) => { element.textContent = t(element.dataset.i18n); });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((element) => { element.placeholder = t(element.dataset.i18nPlaceholder); });
  document.querySelectorAll('[data-i18n-aria]').forEach((element) => { element.setAttribute('aria-label', t(element.dataset.i18nAria)); });
  refreshWarehouseOptions();
  render();
}

function loadRecords() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return Array.isArray(saved) && saved.length ? saved : demoRecords;
  } catch {
    return demoRecords;
  }
}

function saveRecords() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

function toNumber(value) {
  const numericValue = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(numericValue) ? numericValue : 0;
}

function normalizeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? String(value ?? '').trim() : parsed.toISOString().slice(0, 10);
}

function normalizeRecord(row) {
  return {
    '料號': String(row['料號'] ?? '').trim(),
    '品名': String(row['品名'] ?? '').trim(),
    '分類': String(row['分類'] ?? '').trim(),
    '供應商': String(row['供應商'] ?? '').trim(),
    '倉庫': String(row['倉庫'] ?? '').trim(),
    '現有庫存': toNumber(row['現有庫存']),
    '安全庫存': toNumber(row['安全庫存']),
    '單價': toNumber(row['單價']),
    '狀態': ['正常', '注意', '缺料'].includes(row['狀態']) ? row['狀態'] : '正常',
    '更新日期': normalizeDate(row['更新日期'])
  };
}

function isAtRisk(record) {
  return record['現有庫存'] < record['安全庫存'];
}

function formatCurrency(value) {
  return new Intl.NumberFormat(language === 'en' ? 'en-US' : 'zh-TW', { style: 'currency', currency: 'TWD', maximumFractionDigits: 0 }).format(value);
}

function getFilteredRecords() {
  const query = elements.search.value.trim().toLocaleLowerCase();
  return records.filter((record) => {
    const matchQuery = !query || ['料號', '品名', '供應商'].some((column) => String(record[column]).toLocaleLowerCase().includes(query));
    return matchQuery && (!elements.warehouse.value || record['倉庫'] === elements.warehouse.value) && (!elements.status.value || record['狀態'] === elements.status.value) && (!elements.riskOnly.checked || isAtRisk(record));
  });
}

function statusClass(status) {
  return status === '缺料' ? 'status-shortage' : status === '注意' ? 'status-watch' : 'status-normal';
}

function refreshWarehouseOptions() {
  const selectedValue = elements.warehouse.value;
  const warehouses = [...new Set(records.map((record) => record['倉庫']).filter(Boolean))].sort();
  elements.warehouse.innerHTML = `<option value="">${t('allWarehouses')}</option>` + warehouses.map((warehouse) => `<option value="${escapeHtml(warehouse)}">${escapeHtml(warehouse)}</option>`).join('');
  elements.warehouse.value = warehouses.includes(selectedValue) ? selectedValue : '';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#039;', '"': '&quot;' })[character]);
}

function render() {
  const filtered = getFilteredRecords();
  const totalValue = records.reduce((sum, record) => sum + record['現有庫存'] * record['單價'], 0);
  const locale = language === 'en' ? 'en-US' : 'zh-TW';
  elements.itemCount.textContent = records.length.toLocaleString(locale);
  elements.inventoryValue.textContent = formatCurrency(totalValue);
  elements.riskCount.textContent = records.filter(isAtRisk).length.toLocaleString(locale);
  elements.supplierCount.textContent = new Set(records.map((record) => record['供應商']).filter(Boolean)).size.toLocaleString(locale);
  elements.body.innerHTML = filtered.map((record) => {
    const originalIndex = records.indexOf(record);
    const lowStock = isAtRisk(record);
    return `<tr class="${lowStock ? 'risk-row' : ''}">
      <td>${escapeHtml(record['料號'])}</td><td>${escapeHtml(record['品名'])}</td><td>${escapeHtml(record['分類'])}</td><td>${escapeHtml(record['供應商'])}</td><td>${escapeHtml(record['倉庫'])}</td>
      <td class="number stock-pair"><strong class="${lowStock ? '' : 'normal-stock'}">${record['現有庫存'].toLocaleString(locale)}</strong> / ${record['安全庫存'].toLocaleString(locale)}</td>
      <td class="number">${formatCurrency(record['單價'])}</td><td><span class="status ${statusClass(record['狀態'])}">${statusLabel(record['狀態'])}</span></td><td>${escapeHtml(record['更新日期'])}</td>
      <td><button class="row-action" type="button" data-edit-index="${originalIndex}">${t('edit')}</button></td></tr>`;
  }).join('');
  elements.empty.hidden = filtered.length > 0;
}

function showNotice(message, isError = false) {
  elements.notice.textContent = message;
  elements.notice.classList.toggle('error', isError);
  elements.notice.hidden = false;
  window.clearTimeout(showNotice.timeoutId);
  showNotice.timeoutId = window.setTimeout(() => { elements.notice.hidden = true; }, 5000);
}

function ensureXlsx() {
  if (window.XLSX) return true;
  showNotice(t('xlsxUnavailable'), true);
  return false;
}

function exportRecords(rows, fileName) {
  if (!ensureXlsx()) return;
  const headers = COLUMNS.map(columnLabel);
  const worksheet = XLSX.utils.json_to_sheet(rows.map((record) => Object.fromEntries(COLUMNS.map((column) => [columnLabel(column), record[column]]))), { header: headers });
  worksheet['!cols'] = COLUMNS.map((column) => ({ wch: Math.max(columnLabel(column).length + 2, ...rows.map((record) => String(record[column] ?? '').length + 2), 12) }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, t('sheetName'));
  XLSX.writeFile(workbook, fileName);
}

function importFile(file) {
  if (!ensureXlsx()) return;
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const headerToColumn = Object.fromEntries(COLUMNS.flatMap((column) => [[column, column], [translations.en[COLUMN_TRANSLATIONS[column]], column]]));
      const headers = new Set(Object.keys(rawRows[0] ?? {}).map((header) => headerToColumn[header]));
      const missing = COLUMNS.filter((column) => !headers.has(column));
      if (missing.length) throw new Error(t('missingColumns', { columns: missing.map(columnLabel).join(', ') }));
      const imported = rawRows.map((row) => normalizeRecord(Object.fromEntries(Object.entries(row).map(([header, value]) => [headerToColumn[header] ?? header, value])))).filter((record) => record['料號'] && record['品名']);
      if (!imported.length) throw new Error(t('noValidRows'));
      records = imported;
      saveRecords();
      refreshWarehouseOptions();
      render();
      showNotice(t('importSuccess', { count: imported.length }));
    } catch (error) {
      showNotice(t('importFailed', { message: error.message }), true);
    }
  };
  reader.onerror = () => showNotice(t('readFailed'), true);
  reader.readAsArrayBuffer(file);
}

function openDialog(index = null) {
  editingIndex = index;
  const record = index === null ? normalizeRecord({ '更新日期': new Date() }) : records[index];
  elements.dialogTitle.textContent = index === null ? t('newItem') : t('editItem', { item: record['料號'] });
  COLUMNS.forEach((column) => { elements.form.elements[column].value = record[column] ?? ''; });
  elements.dialog.showModal();
}

elements.search.addEventListener('input', render);
elements.warehouse.addEventListener('change', render);
elements.status.addEventListener('change', render);
elements.riskOnly.addEventListener('change', render);
document.querySelector('#clearFiltersButton').addEventListener('click', () => { elements.search.value = ''; elements.warehouse.value = ''; elements.status.value = ''; elements.riskOnly.checked = false; render(); });
document.querySelector('#languageSelect').addEventListener('change', (event) => { language = event.target.value; localStorage.setItem(LANGUAGE_KEY, language); updateLanguage(); });
document.querySelector('#downloadTemplateButton').addEventListener('click', () => exportRecords([normalizeRecord({ '料號': 'RM-00001', '品名': '範例品項', '分類': '原物料', '供應商': '範例供應商', '倉庫': '桃園 A', '現有庫存': 0, '安全庫存': 0, '單價': 0, '狀態': '正常', '更新日期': new Date() })], t('templateFile')));
document.querySelector('#exportButton').addEventListener('click', () => { const rows = getFilteredRecords(); if (!rows.length) return showNotice(t('noExportRows'), true); exportRecords(rows, t('exportFile')); });
document.querySelector('#fileInput').addEventListener('change', (event) => { if (event.target.files[0]) importFile(event.target.files[0]); event.target.value = ''; });
document.querySelector('#addRecordButton').addEventListener('click', () => openDialog());
document.querySelector('#closeDialogButton').addEventListener('click', () => elements.dialog.close());
document.querySelector('#cancelDialogButton').addEventListener('click', () => elements.dialog.close());
elements.body.addEventListener('click', (event) => { const index = event.target.dataset.editIndex; if (index !== undefined) openDialog(Number(index)); });
elements.form.addEventListener('submit', (event) => { event.preventDefault(); const record = normalizeRecord(Object.fromEntries(new FormData(elements.form))); const isNewRecord = editingIndex === null; if (isNewRecord) records.unshift(record); else records[editingIndex] = record; saveRecords(); refreshWarehouseOptions(); render(); elements.dialog.close(); showNotice(t(isNewRecord ? 'added' : 'updated')); });

updateLanguage();
