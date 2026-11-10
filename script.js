// ================= CONFIGURATION =================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz_mEnxHFDApvKuuzLT1-Qzi1hfHHyhZOmtfH0KLDLACvfS_y_V9dyIVM85SgWUrIMN/exec';
const ROOT_FOLDER_ID = '1SLd83WomxtEH3UPuzo3GpSBbUwiaWPHE';
const TENANT_SPREADSHEET_ID = '1eQUsnIRkXBqohI_6F-oV95bKZAn75lilbkx4KEqd4dw';

const PROPERTIES = ['ECO 1', 'ECO 2', 'GREEN', 'HOMEY', 'ADI', 'KALAYAAN', 'PLEASANT', 'DREAM', 'PENTHAUZ'];
const DEFAULT_PROPERTY = 'ECO 1';

function calculateConsumption(prevReading, presentReading) {
  return (Number(presentReading) || 0) - (Number(prevReading) || 0);
}

function getPresentReading(record) {
  const keys = new Set(['presentreading', 'currentreading', 'current', 'upon', 'meterpresentreading', 'meterpresent']);
  const values = Object.entries(record || {})
    .filter(([key]) => keys.has(key.replace(/[^a-z]/gi, '').toLowerCase()))
    .map(([, value]) => value)
    .filter(value => value !== undefined && value !== null && value !== '');
  const nonZeroValue = values.find(value => Number(value) !== 0);
  return Number(nonZeroValue !== undefined ? nonZeroValue : values[0]) || 0;
}

function getImportedReading(rowKeys, keys) {
  for (const key of keys) {
    if (rowKeys[key] !== undefined && rowKeys[key] !== null && rowKeys[key] !== '') {
      return Number(rowKeys[key]) || 0;
    }
  }
  return 0;
}

let appData = {};
let activeTab = null;
let workbookCache = null;
let currentEditingBill = null;
let isAddMode = false;
let sortField = null;
let sortDirection = 'asc';

const tabs = document.querySelectorAll('.office-tab');
const contents = document.querySelectorAll('.ribbon-content');

tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    tabs.forEach(t => t.classList.remove('active'));
    contents.forEach(c => c.classList.add('hidden'));
    tab.classList.add('active');
    document.getElementById(`ribbon-${tab.dataset.target}`).classList.remove('hidden');
  });
});

async function initDatabase() {
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = "Loading Data from Google Sheets...";

  try {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'load' })
    });
    const res = await response.json();

    if (res.success) {
      appData = res.data;
      for (const tabName in appData) {
        if (!Array.isArray(appData[tabName])) continue;
        appData[tabName].forEach(bill => {
          bill.pdfUrl = bill.pdfUrl || '';
          if (!bill.status) bill.status = 'Pending';
          if (!bill.dateSent) bill.dateSent = '';
          bill.balance = parseFloat(bill.balance) || 0;
          bill.penalty = parseFloat(bill.penalty) || 0;
          if (!bill.property || !PROPERTIES.includes(bill.property)) bill.property = tabName;
          
          bill.rate = parseFloat(bill.rate) || 0;
          bill.prevReading = parseFloat(bill.prevReading) || 0;
          bill.presentReading = getPresentReading(bill);
          bill.consumption = calculateConsumption(bill.prevReading, bill.presentReading);
          bill.total = bill.consumption * bill.rate;
        });
      }

      PROPERTIES.forEach(prop => {
        if (!appData[prop]) appData[prop] = [];
      });

      activeTab = PROPERTIES[0];
      renderTabs();
      renderSpreadsheet();
    } else {
      alert("Failed to load database: " + res.error);
    }
  } catch (error) {
    console.error(error);
    alert("Could not connect to the Google Apps Script.");
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
}

async function syncToDatabase(bill, tabName, actionType) {
  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: actionType, 
        tabName: tabName,
        tenantName: bill.tenantName,
        property: bill.property,
        unitNo: bill.unitNo,
        prevReading: bill.prevReading,
        presentReading: bill.presentReading,
        currentReading: bill.presentReading,
        current: bill.presentReading,
        upon: bill.presentReading,
        CURRENT: bill.presentReading,
        UPON: bill.presentReading,
        consumption: bill.consumption,
        rate: bill.rate,
        balance: bill.balance,
        penalty: bill.penalty || 0,
        total: bill.total,
        billingDate: bill.billingDate,
        status: bill.status,
        pdfUrl: bill.pdfUrl || '',
        dateSent: bill.dateSent || ''
      })
    });
    return await res.json();
  } catch (e) {
    console.error(e);
    return { success: false };
  }
}

function normalizeTenantValue(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLocaleLowerCase();
}

function normalizeHeader(value) {
  return String(value == null ? '' : value).replace(/[^a-z0-9]/gi, '').toLocaleLowerCase();
}

async function loadTenantRecords(property) {
  const queryUrl = `https://docs.google.com/spreadsheets/d/${TENANT_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(property)}`;
  const response = await fetch(queryUrl);
  if (!response.ok) {
    throw new Error(`Could not read the "${property}" tab from the main tenant spreadsheet.`);
  }

  const csvText = await response.text();
  const workbook = XLSX.read(csvText, { type: 'string' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet) throw new Error(`The "${property}" tab could not be parsed.`);

  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
  const headerRowIndex = rows.findIndex(row => {
    const headers = row.map(normalizeHeader);
    return headers.includes('name') && headers.includes('unit');
  });
  if (headerRowIndex === -1) {
    throw new Error(`The "${property}" tab must contain "Name" and "Unit" columns.`);
  }

  const headers = rows[headerRowIndex].map(normalizeHeader);
  const nameColumn = headers.indexOf('name');
  const unitColumn = headers.indexOf('unit');
  return rows.slice(headerRowIndex + 1)
    .map(row => ({
      tenantName: String(row[nameColumn] == null ? '' : row[nameColumn]).trim(),
      unitNo: String(row[unitColumn] == null ? '' : row[unitColumn]).trim()
    }))
    .filter(record => record.tenantName && record.unitNo);
}

function addToIndex(index, key, record) {
  if (!key) return;
  if (!index.has(key)) index.set(key, []);
  index.get(key).push(record);
}

async function syncTenantDetails() {
  const button = document.getElementById('syncTenantsBtn');
  button.disabled = true;
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = 'Loading main tenant spreadsheet...';

  let updatedCount = 0;
  let unchangedCount = 0;
  let unmatchedCount = 0;
  let ambiguousCount = 0;
  let failedTenant = '';

  try {
    const recordsByProperty = {};
    for (const property of PROPERTIES) {
      recordsByProperty[property] = await loadTenantRecords(property);
    }

    for (const property of PROPERTIES) {
      const sourceRecords = recordsByProperty[property];
      const sourceByPair = new Map();
      const sourceByUnit = new Map();
      const sourceByName = new Map();
      sourceRecords.forEach(record => {
        addToIndex(sourceByPair, `${normalizeTenantValue(record.tenantName)}|${normalizeTenantValue(record.unitNo)}`, record);
        addToIndex(sourceByUnit, normalizeTenantValue(record.unitNo), record);
        addToIndex(sourceByName, normalizeTenantValue(record.tenantName), record);
      });

      const bills = appData[property] || [];
      const localUnitCounts = new Map();
      const localNameCounts = new Map();
      bills.forEach(bill => {
        const unit = normalizeTenantValue(bill.unitNo);
        const name = normalizeTenantValue(bill.tenantName);
        localUnitCounts.set(unit, (localUnitCounts.get(unit) || 0) + 1);
        localNameCounts.set(name, (localNameCounts.get(name) || 0) + 1);
      });
      const matchedSources = new Set();

      for (let index = 0; index < bills.length; index++) {
        const bill = bills[index];
        const nameKey = normalizeTenantValue(bill.tenantName);
        const unitKey = normalizeTenantValue(bill.unitNo);
        const pairMatches = sourceByPair.get(`${nameKey}|${unitKey}`) || [];
        let matches = pairMatches;

        if (matches.length === 0 && unitKey && localUnitCounts.get(unitKey) === 1) {
          matches = sourceByUnit.get(unitKey) || [];
        }
        if (matches.length === 0 && nameKey && localNameCounts.get(nameKey) === 1) {
          matches = sourceByName.get(nameKey) || [];
        }

        if (matches.length === 0) {
          unmatchedCount++;
          continue;
        }
        if (matches.length !== 1 || matchedSources.has(matches[0])) {
          ambiguousCount++;
          continue;
        }

        const source = matches[0];
        matchedSources.add(source);
        if (bill.tenantName === source.tenantName && String(bill.unitNo || '').trim() === source.unitNo) {
          unchangedCount++;
          continue;
        }

        const updatedBill = { ...bill, tenantName: source.tenantName, unitNo: source.unitNo, property };
        const hasDuplicateName = bills.some((other, otherIndex) =>
          otherIndex !== index && normalizeTenantValue(other.tenantName) === normalizeTenantValue(updatedBill.tenantName)
        );
        if (hasDuplicateName) {
          ambiguousCount++;
          continue;
        }

        failedTenant = `${property}: ${bill.tenantName} (${bill.unitNo || 'no unit'})`;
        document.getElementById('loadingText').textContent = `Updating ${failedTenant}...`;
        if (bill.tenantName !== updatedBill.tenantName) {
          const deleteResponse = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'delete', tabName: property, tenantName: bill.tenantName })
          });
          const deleteResult = await deleteResponse.json();
          if (!deleteResult.success) {
            throw new Error(deleteResult.error || `Could not rename ${failedTenant}.`);
          }

          const saveResult = await syncToDatabase(updatedBill, property, 'save');
          if (!saveResult.success) {
            const rollbackResult = await syncToDatabase(bill, property, 'save');
            if (!rollbackResult.success) {
              throw new Error(`Updating ${failedTenant} failed, and its original row could not be restored. Reload and verify the spreadsheet.`);
            }
            throw new Error(saveResult.error || `Could not save the tenant name for ${failedTenant}; the old row was restored.`);
          }
        } else {
          const saveResult = await syncToDatabase(updatedBill, property, 'update');
          if (!saveResult.success) throw new Error(saveResult.error || `Could not update ${failedTenant}.`);
        }

        bills[index] = updatedBill;
        updatedCount++;
        failedTenant = '';
      }
    }

    renderSpreadsheet();
    alert(`Tenant sync complete.\n\nUpdated: ${updatedCount}\nAlready up to date: ${unchangedCount}\nNo match: ${unmatchedCount}\nSkipped as ambiguous: ${ambiguousCount}`);
  } catch (error) {
    console.error(error);
    const progress = updatedCount ? ` ${updatedCount} record(s) were updated before the error.` : '';
    alert(`Tenant sync stopped${failedTenant ? ` while updating ${failedTenant}` : ''}: ${error.message}${progress}`);
  } finally {
    button.disabled = false;
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
}

document.getElementById('syncTenantsBtn').addEventListener('click', syncTenantDetails);

function formatPeso(amount) {
  return '₱' + (Number(amount) || 0).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDate(dateObj) {
  const yyyy = dateObj.getFullYear();
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function toDateInputValue(val) {
  if (!val) return '';
  if (val instanceof Date) return isNaN(val) ? '' : formatDate(val);
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d) ? '' : formatDate(d);
  }
  const s = String(val).trim();
  if (!s) return '';
  if (s.includes('T')) return s.split('T')[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d) ? '' : formatDate(d);
}

function formatDisplayDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}

function getMonthFolderName(billingDate) {
  let d = new Date(billingDate);
  if (isNaN(d.getTime())) d = new Date();
  return d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
}

function renderTabs() {
  const tabsContainer = document.getElementById('tabsContainer');
  tabsContainer.innerHTML = '';

  PROPERTIES.forEach(tabName => {
    const isActive = tabName === activeTab;
    const tabBtn = document.createElement('button');
    tabBtn.className = `px-4 py-2 rounded-t-lg font-bold text-sm transition-colors border-b-2 whitespace-nowrap ${isActive ? 'bg-white border-blue-600 text-blue-700 shadow-sm' : 'bg-slate-200 border-transparent text-slate-600 hover:bg-slate-300'}`;
    tabBtn.textContent = tabName;
    
    tabBtn.addEventListener('click', () => {
      activeTab = tabName;
      document.getElementById('searchInput').value = ''; 
      document.getElementById('statusFilter').value = 'ALL';
      renderTabs();
      renderSpreadsheet();
    });

    tabsContainer.appendChild(tabBtn);
  });
}

function renderSpreadsheet() {
  const spreadsheetBody = document.getElementById('spreadsheetBody');
  spreadsheetBody.innerHTML = '';
  document.getElementById('consumptionSum').textContent = '0 cu.m';
  document.getElementById('totalSum').textContent = formatPeso(0);
  document.querySelectorAll('#spreadsheetTable th[data-sort-field]').forEach(header => {
    const indicator = header.querySelector('.sort-indicator');
    const isActive = header.dataset.sortField === sortField;
    header.setAttribute('aria-sort', isActive ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none');
    if (indicator) indicator.textContent = isActive ? (sortDirection === 'asc' ? ' ▲' : ' ▼') : ' ↕';
  });
  
  let pending = 0, generated = 0, sent = 0;

  if (!activeTab || !appData[activeTab] || appData[activeTab].length === 0) {
    spreadsheetBody.innerHTML = `<tr><td colspan="15" class="px-4 py-12 text-center text-slate-400">No data in "${activeTab}" tab.</td></tr>`;
    document.getElementById('countPending').textContent = `0 Pending`;
    document.getElementById('countGenerated').textContent = `0 Generated`;
    document.getElementById('countSent').textContent = `0 Sent`;
    return;
  }

  const bills = appData[activeTab];
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const statusValue = document.getElementById('statusFilter').value;

  bills.forEach(b => {
    if (b.status === 'Pending') pending++;
    else if (b.status === 'Generated') generated++;
    else if (b.status === 'Sent') sent++;
  });

  const filteredBills = bills.filter(b => {
    const matchesSearch = (b.tenantName || '').toLowerCase().includes(searchTerm) 
      || (b.unitNo || '').toLowerCase().includes(searchTerm);
    const matchesStatus = (statusValue === 'ALL') || (b.status === statusValue);
    return matchesSearch && matchesStatus;
  });

  if (sortField) {
    const numericFields = ['prevReading', 'presentReading', 'consumption', 'rate', 'balance', 'penalty', 'total'];
    filteredBills.sort((a, b) => {
      const left = a[sortField];
      const right = b[sortField];
      const leftEmpty = left == null || left === '';
      const rightEmpty = right == null || right === '';
      if (leftEmpty || rightEmpty) return leftEmpty === rightEmpty ? 0 : (leftEmpty ? 1 : -1);

      let comparison;
      if (numericFields.includes(sortField)) {
        comparison = (Number(left) || 0) - (Number(right) || 0);
      } else if (sortField === 'billingDate' || sortField === 'dateSent') {
        comparison = toDateInputValue(left).localeCompare(toDateInputValue(right));
      } else {
        comparison = String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: 'base' });
      }
      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }

  if (filteredBills.length === 0) {
    spreadsheetBody.innerHTML = `<tr><td colspan="15" class="px-4 py-12 text-center text-slate-400">No bills found matching your search.</td></tr>`;
  } else {
    filteredBills.forEach((bill) => {
      let badgeClass = 'bg-yellow-100 text-yellow-800';
      if (bill.status === 'Generated') badgeClass = 'bg-blue-100 text-blue-800';
      if (bill.status === 'Sent') badgeClass = 'bg-green-100 text-green-800';

      const hasPdf = bill.pdfUrl && String(bill.pdfUrl).trim() !== '';
      let pdfCell = hasPdf ? `<a href="${bill.pdfUrl}" target="_blank" rel="noopener" class="text-blue-600 hover:underline font-bold text-xs flex justify-center items-center gap-1">View PDF</a>` : `<span class="text-slate-300 text-xs">—</span>`;

      const isDisabled = bill.status !== 'Sent';
      const disabledAttr = isDisabled ? 'disabled' : '';
      
      let dateSentHtml = '';
      if (bill.dateSent && bill.dateSent !== '') {
        try {
          const cleanDateStr = bill.dateSent.split('T')[0];
          dateSentHtml = `<input type="date" class="table-date-input" value="${cleanDateStr}" ${disabledAttr} onchange="updateDateSent('${bill.tenantName.replace(/'/g, "\\'")}', this.value)">`;
        } catch(e) {
          dateSentHtml = `<input type="date" class="table-date-input" ${disabledAttr} onchange="updateDateSent('${bill.tenantName.replace(/'/g, "\\'")}', this.value)">`;
        }
      } else {
        dateSentHtml = `<input type="date" class="table-date-input" ${disabledAttr} onchange="updateDateSent('${bill.tenantName.replace(/'/g, "\\'")}', this.value)">`;
      }

      const safeName = bill.tenantName.replace(/'/g, "\\'");

      let actionBtn = '';
      if (bill.status === 'Pending') {
        actionBtn = `
          <div class="flex justify-end items-center gap-1">
            <button onclick="generateSingle('${safeName}')" class="bg-slate-800 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-900 transition-colors" title="Generate PDF">Generate</button>
            <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit bill">Edit</button>
            <button onclick="deleteBill('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete bill">Del</button>
          </div>`;
      } else if (bill.status === 'Generated') {
        actionBtn = `
          <div class="flex justify-end items-center gap-1 flex-wrap">
            <button onclick="regenerateSingle('${safeName}')" class="bg-slate-500 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-600 transition-colors" title="Re-generate">Regenerate</button>
            <button onclick="markAsSent('${safeName}')" class="bg-green-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-green-700 transition-colors">Sent</button>
            <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit bill">Edit</button>
            <button onclick="deleteBill('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete bill">Del</button>
          </div>`;
      } else {
        actionBtn = `
          <div class="flex justify-end items-center gap-1">
            <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit bill">Edit</button>
            <button onclick="deleteBill('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete bill">Del</button>
          </div>`;
      }

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-50 transition-colors";

      tr.innerHTML = `
        <td class="px-4 py-3 font-medium text-slate-800">${bill.tenantName}</td>
        <td class="px-4 py-3 text-slate-600">${bill.property || activeTab}</td>
        <td class="px-4 py-3 text-slate-600">${bill.unitNo || ''}</td>
        <td class="px-4 py-3 text-slate-600">${formatDisplayDate(bill.billingDate)}</td>
        <td class="px-4 py-3 text-slate-600">${bill.prevReading || 0}</td>
        <td class="px-4 py-3 text-slate-600">${bill.presentReading || 0}</td>
        <td class="px-4 py-3 text-slate-600">${bill.consumption || 0} cu.m</td>
        <td class="px-4 py-3 text-slate-600">${formatPeso(bill.rate)}</td>
        <td class="px-4 py-3 text-slate-600">${formatPeso(bill.balance)}</td>
        <td class="px-4 py-3 text-slate-600">${formatPeso(bill.penalty)}</td>
        <td class="px-4 py-3 text-slate-600 font-medium">${formatPeso(bill.total)}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 text-[10px] uppercase tracking-wide rounded-full font-bold ${badgeClass}">${bill.status}</span></td>
        <td class="px-4 py-3 text-center">${pdfCell}</td>
        <td class="px-4 py-3 text-right">${actionBtn}</td>
        <td class="px-4 py-3 text-center">${dateSentHtml}</td>
      `;
      const billIndex = appData[activeTab].indexOf(bill);
      tr.dataset.billIndex = billIndex;
      const cells = tr.querySelectorAll('td');
      const editorFields = [
        ['tenantName', bill.tenantName || '', 'text'],
        ['property', bill.property || activeTab, 'select'],
        ['unitNo', bill.unitNo || '', 'text'],
        ['billingDate', toDateInputValue(bill.billingDate), 'date'],
        ['prevReading', bill.prevReading || 0, 'number'],
        ['presentReading', bill.presentReading || 0, 'number'],
        ['consumption', bill.consumption || 0, 'number'],
        ['rate', bill.rate || 0, 'number'],
        ['balance', bill.balance || 0, 'number'],
        ['penalty', bill.penalty || 0, 'number']
      ];
      editorFields.forEach(([field, value, type], index) => {
        const editor = document.createElement(type === 'select' ? 'select' : 'input');
        editor.className = 'spreadsheet-editor';
        editor.dataset.field = field;
        editor.setAttribute('aria-label', field);
        if (type === 'select') {
          const options = field === 'property' ? PROPERTIES : ['Pending', 'Generated', 'Sent'];
          options.forEach(optionValue => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue;
            editor.appendChild(option);
          });
          editor.value = value;
        } else {
          editor.type = type;
          editor.value = value;
          if (type === 'number') editor.step = '0.01';
          if (field === 'consumption') {
            editor.readOnly = true;
            editor.title = 'Calculated from present reading minus previous reading';
          }
        }
        cells[index].replaceChildren(editor);
        cells[index].classList.add('spreadsheet-cell');
      });

      const statusEditor = document.createElement('select');
      statusEditor.className = 'spreadsheet-editor';
      statusEditor.dataset.field = 'status';
      statusEditor.setAttribute('aria-label', 'status');
      ['Pending', 'Generated', 'Sent'].forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        statusEditor.appendChild(option);
      });
      statusEditor.value = bill.status || 'Pending';
      cells[11].replaceChildren(statusEditor);
      cells[11].classList.add('spreadsheet-cell');

      const dateSentEditor = document.createElement('input');
      dateSentEditor.type = 'date';
      dateSentEditor.className = 'spreadsheet-editor';
      dateSentEditor.dataset.field = 'dateSent';
      dateSentEditor.setAttribute('aria-label', 'dateSent');
      dateSentEditor.value = toDateInputValue(bill.dateSent);
      cells[14].replaceChildren(dateSentEditor);
      cells[14].classList.add('spreadsheet-cell');
      spreadsheetBody.appendChild(tr);
    });
  }

  const visibleConsumption = filteredBills.reduce((sum, bill) => sum + (Number(bill.consumption) || 0), 0);
  const visibleTotal = filteredBills.reduce((sum, bill) => sum + (Number(bill.total) || 0), 0);
  document.getElementById('consumptionSum').textContent = `${visibleConsumption.toLocaleString('en-PH', { maximumFractionDigits: 2 })} cu.m`;
  document.getElementById('totalSum').textContent = formatPeso(visibleTotal);
  document.getElementById('countPending').textContent = `${pending} Pending`;
  document.getElementById('countGenerated').textContent = `${generated} Generated`;
  document.getElementById('countSent').textContent = `${sent} Sent`;
}

document.querySelector('#spreadsheetTable thead').addEventListener('click', event => {
  const header = event.target.closest('th[data-sort-field]');
  if (!header) return;

  if (sortField === header.dataset.sortField) {
    sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
  } else {
    sortField = header.dataset.sortField;
    sortDirection = 'asc';
  }
  renderSpreadsheet();
});

document.addEventListener('wheel', event => {
  const input = event.target.closest('input[type="number"]');
  if (input && document.activeElement === input) input.blur();
}, { capture: true, passive: true });

document.getElementById('spreadsheetBody').addEventListener('change', async event => {
  const editor = event.target.closest('.spreadsheet-editor');
  if (!editor) return;

  const row = editor.closest('tr');
  const originalTab = activeTab;
  const billIndex = Number(row.dataset.billIndex);
  const originalBill = appData[originalTab] && appData[originalTab][billIndex];
  if (!originalBill) return;

  const oldBill = { ...originalBill };
  const updatedBill = { ...oldBill };
  const field = editor.dataset.field;
  let value = editor.value;

  if (field === 'tenantName') {
    value = value.trim();
    if (!value) {
      alert('Tenant name is required.');
      renderSpreadsheet();
      return;
    }
  } else if (['prevReading', 'presentReading', 'consumption', 'rate', 'balance', 'penalty'].includes(field)) {
    value = Number(value);
    if (!Number.isFinite(value)) {
      alert('Enter a valid number.');
      renderSpreadsheet();
      return;
    }
  }

  updatedBill[field] = value;
  if (field === 'prevReading' || field === 'presentReading') {
    updatedBill.consumption = calculateConsumption(updatedBill.prevReading, updatedBill.presentReading);
  }
  updatedBill.property = updatedBill.property || originalTab;
  updatedBill.total = (Number(updatedBill.consumption) || 0) * (Number(updatedBill.rate) || 0);

  const targetTab = field === 'property' ? value : originalTab;
  const identityChanged = targetTab !== originalTab || updatedBill.tenantName !== oldBill.tenantName;
  const targetBills = appData[targetTab] || [];
  const isDuplicate = targetBills.some((bill, index) =>
    bill.tenantName === updatedBill.tenantName && (targetTab !== originalTab || index !== billIndex)
  );
  if (isDuplicate) {
    alert(`A bill with this tenant name already exists in "${targetTab}".`);
    renderSpreadsheet();
    return;
  }

  editor.disabled = true;
  try {
    if (identityChanged) {
      const deleteResponse = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'delete', tabName: originalTab, tenantName: oldBill.tenantName })
      });
      const deleteResult = await deleteResponse.json();
      if (!deleteResult.success) throw new Error(deleteResult.error || 'Could not move or rename the bill.');

      const saveResult = await syncToDatabase(updatedBill, targetTab, 'save');
      if (!saveResult.success) {
        const rollbackResult = await syncToDatabase(oldBill, originalTab, 'save');
        if (!rollbackResult.success) {
          throw new Error('Saving failed and the original row could not be restored. Reload and check Google Sheets.');
        }
        throw new Error(saveResult.error || 'Could not save the edit. The original row was restored.');
      }

      appData[originalTab].splice(billIndex, 1);
      if (!appData[targetTab]) appData[targetTab] = [];
      appData[targetTab].push(updatedBill);
      activeTab = targetTab;
      renderTabs();
    } else {
      const saveResult = await syncToDatabase(updatedBill, originalTab, 'update');
      if (!saveResult.success) throw new Error(saveResult.error || 'Could not save the edit.');
      appData[originalTab][billIndex] = updatedBill;
    }
    renderSpreadsheet();
  } catch (error) {
    console.error(error);
    alert(`Could not save this cell: ${error.message}`);
    renderSpreadsheet();
  }
});

document.getElementById('spreadsheetBody').addEventListener('keydown', event => {
  if (event.key === 'Enter' && event.target.matches('.spreadsheet-editor:not(select)')) {
    event.preventDefault();
    event.target.blur();
  } else if (event.key === 'Escape' && event.target.matches('.spreadsheet-editor')) {
    renderSpreadsheet();
  }
});

window.openAddModal = function() {
  if (!activeTab) return alert("Select a tab first.");
  isAddMode = true;
  currentEditingBill = null;
  document.getElementById('editTenantName').value = '';
  document.getElementById('editTenantName').readOnly = false;
  document.getElementById('editProperty').value = activeTab || DEFAULT_PROPERTY;
  document.getElementById('editUnitNo').value = '';
  document.getElementById('editPrevReading').value = '0';
  document.getElementById('editPresentReading').value = '0';
  document.getElementById('editConsumption').value = calculateConsumption(0, 0);
  document.getElementById('editRate').value = '0';
  document.getElementById('editBalance').value = '0';
  document.getElementById('editPenalty').value = '0';
  document.getElementById('editBillingDate').value = new Date().toISOString().split('T')[0];
  document.getElementById('editStatus').value = 'Pending';
  document.getElementById('editDateSent').value = '';
  document.getElementById('editBillModal').querySelector('h3').lastChild.textContent = ' Add Bill';
  document.getElementById('editBillModal').classList.remove('hidden');
  document.getElementById('editBillModal').classList.add('flex');
};

window.openEditModal = function(tenantName) {
  if (!activeTab) return;
  const bill = appData[activeTab].find(b => b.tenantName === tenantName);
  if (!bill) return alert("Bill not found.");
  isAddMode = false;
  currentEditingBill = bill;

  document.getElementById('editTenantName').value = bill.tenantName || '';
  document.getElementById('editTenantName').readOnly = false;
  document.getElementById('editProperty').value = bill.property || activeTab;
  document.getElementById('editUnitNo').value = bill.unitNo || '';
  document.getElementById('editPrevReading').value = bill.prevReading || 0;
  document.getElementById('editPresentReading').value = bill.presentReading || 0;
  document.getElementById('editConsumption').value = calculateConsumption(bill.prevReading, bill.presentReading);
  document.getElementById('editRate').value = bill.rate || 0;
  document.getElementById('editBalance').value = bill.balance || 0;
  document.getElementById('editPenalty').value = bill.penalty || 0;
  document.getElementById('editBillingDate').value = toDateInputValue(bill.billingDate);
  document.getElementById('editStatus').value = bill.status || 'Pending';
  document.getElementById('editDateSent').value = toDateInputValue(bill.dateSent);

  const modal = document.getElementById('editBillModal');
  modal.querySelector('h3').lastChild.textContent = ' Edit Bill';
  modal.classList.remove('hidden');
  modal.classList.add('flex');
};

window.closeEditModal = function() {
  const modal = document.getElementById('editBillModal');
  modal.classList.add('hidden');
  modal.classList.remove('flex');
  currentEditingBill = null;
  isAddMode = false;
};

window.saveBillEdit = async function() {
  const name = document.getElementById('editTenantName').value.trim();
  if (!name) return alert("Tenant name is required.");

  const property = document.getElementById('editProperty').value;
  if (!PROPERTIES.includes(property)) return alert("Please select a valid property.");

  const unitNo = document.getElementById('editUnitNo').value.trim();
  const prevReading = Number(document.getElementById('editPrevReading').value) || 0;
  const presentReading = Number(document.getElementById('editPresentReading').value) || 0;
  const consumption = calculateConsumption(prevReading, presentReading);
  const rate = Number(document.getElementById('editRate').value) || 0;
  const balance = Number(document.getElementById('editBalance').value) || 0;
  const penalty = Number(document.getElementById('editPenalty').value) || 0;
  const billingDate = document.getElementById('editBillingDate').value;
  const status = document.getElementById('editStatus').value;
  const dateSent = document.getElementById('editDateSent').value || '';

  const total = consumption * rate;

  const targetTab = property;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = isAddMode ? 'Creating bill...' : 'Saving changes...';

  try {
    if (isAddMode) {
      if (!appData[targetTab]) appData[targetTab] = [];
      if (appData[targetTab].some(b => b.tenantName === name)) {
        throw new Error(`A bill with this tenant name already exists in "${targetTab}".`);
      }
      const newBill = {
        tenantName: name,
        property, unitNo, prevReading, presentReading, consumption, rate, balance, penalty, total,
        billingDate, status, pdfUrl: '', dateSent
      };
      appData[targetTab].push(newBill);
      await syncToDatabase(newBill, targetTab, 'save');
      activeTab = targetTab;
    } else {
      if (!currentEditingBill) throw new Error("No bill selected.");
      const originalName = currentEditingBill.tenantName;
      const originalTab = activeTab;
      const propertyChanged = property !== originalTab;
      
      if (!propertyChanged && name !== originalName && appData[originalTab].some(b => b.tenantName === name)) {
        throw new Error(`Another bill with this tenant name already exists in "${originalTab}".`);
      }

      const updated = {
        ...currentEditingBill,
        tenantName: name,
        property, unitNo, prevReading, presentReading, consumption, rate, balance, penalty, total,
        billingDate, status, dateSent
      };

      if (propertyChanged) {
        await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({ action: 'delete', tabName: originalTab, tenantName: originalName })
        });
        const oldIdx = appData[originalTab].findIndex(b => b.tenantName === originalName);
        if (oldIdx !== -1) appData[originalTab].splice(oldIdx, 1);

        if (!appData[targetTab]) appData[targetTab] = [];
        if (appData[targetTab].some(b => b.tenantName === name)) {
          throw new Error(`A bill with this tenant name already exists in "${targetTab}".`);
        }
        appData[targetTab].push(updated);
        await syncToDatabase(updated, targetTab, 'save');
        activeTab = targetTab;
      } else {
        if (name !== originalName) {
          await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'delete', tabName: originalTab, tenantName: originalName })
          });
          const idx = appData[originalTab].findIndex(b => b.tenantName === originalName);
          if (idx !== -1) appData[originalTab][idx] = updated;
          await syncToDatabase(updated, targetTab, 'save');
        } else {
          const idx = appData[originalTab].findIndex(b => b.tenantName === originalName);
          if (idx !== -1) appData[originalTab][idx] = updated;
          await syncToDatabase(updated, targetTab, 'update');
        }
      }
    }

    closeEditModal();
    renderTabs();
    renderSpreadsheet();
  } catch (err) {
    console.error(err);
    alert("Error: " + err.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

window.deleteBill = async function(tenantName) {
  if (!activeTab) return;
  const idx = appData[activeTab].findIndex(b => b.tenantName === tenantName);
  if (idx === -1) return alert("Bill not found.");
  const bill = appData[activeTab][idx];

  const confirmDelete = confirm(`Delete bill for "${tenantName}" (Unit ${bill.unitNo}) from "${activeTab}" tab? This will remove the row from Google Sheets and cannot be undone.`);
  if (!confirmDelete) return;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Deleting ${tenantName}...`;

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'delete', tabName: activeTab, tenantName })
    });
    const resData = await res.json();
    if (resData.success) {
      appData[activeTab].splice(idx, 1);
      renderSpreadsheet();
    } else {
      alert("Failed to delete: " + (resData.error || "Unknown error"));
    }
  } catch (e) {
    console.error(e);
    alert("Error deleting bill: " + e.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

document.getElementById('searchInput').addEventListener('input', renderSpreadsheet);
document.getElementById('statusFilter').addEventListener('change', renderSpreadsheet);

window.updateDateSent = async function(tenantName, newDateStr) {
  const billIndex = appData[activeTab].findIndex(b => b.tenantName === tenantName);
  if (billIndex === -1) return;
  const bill = appData[activeTab][billIndex];

  bill.dateSent = newDateStr;
  await syncToDatabase(bill, activeTab, 'update');
};

async function generatePDFDoc(bill) {
  const safeSetText = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  try {
    const subtotal = Number(bill.consumption || 0) * Number(bill.rate || 0);
    const finalTotal = subtotal + Number(bill.balance || 0) + Number(bill.penalty || 0);

    safeSetText('outDate', formatDisplayDate(bill.billingDate) || '');
    safeSetText('outTenant', bill.tenantName || '');
    safeSetText('outProperty', bill.property || activeTab);
    safeSetText('outUnitNo', bill.unitNo || '');
    safeSetText('outPrevReading', bill.prevReading || 0);
    safeSetText('outPresentReading', bill.presentReading || 0);
    safeSetText('outConsumption', bill.consumption || 0);
    safeSetText('outRate', formatPeso(bill.rate || 0));
    safeSetText('outBalance', formatPeso(bill.balance || 0));
    safeSetText('outPenalty', formatPeso(bill.penalty || 0));
    safeSetText('outSubtotalBottom', formatPeso(subtotal));
    safeSetText('outTotal', formatPeso(finalTotal));
  } catch (err) {
    throw new Error("Failed to populate template: " + err.message);
  }

  try {
    if (typeof html2canvas === 'undefined' || typeof window.jspdf === 'undefined') {
       throw new Error("Required libraries not loaded.");
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const billPage = document.getElementById('bill-page');

    const canvas = await html2canvas(billPage, { scale: 2, useCORS: true, windowWidth: 800 });
    doc.addImage(canvas.toDataURL('image/jpeg', 0.98), 'JPEG', 0, 0, 210, 297);

    return doc;
  } catch (err) {
    throw new Error("Canvas/PDF rendering failed: " + err.message);
  }
}

async function uploadPdfToDrive(bill, doc) {
  const fileName = `Water Bill - ${bill.tenantName}.pdf`;
  const monthName = getMonthFolderName(bill.billingDate);
  const propertyName = bill.property || activeTab;
  const base64Data = doc.output('datauristring').split(',')[1];

  const uploadRes = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'uploadPdf',
      fileName: fileName,
      base64: base64Data,
      rootFolderId: ROOT_FOLDER_ID,
      propertyName: propertyName,
      monthName: monthName,
      tenantName: bill.tenantName
    })
  });

  let uploadData;
  try {
    uploadData = await uploadRes.json();
  } catch (parseError) {
    throw new Error("Invalid response from server. Check Apps Script deployment.");
  }

  const isSuccess = uploadData && (uploadData.success === true || (uploadData.data && uploadData.data.success === true));
  const fileUrl = uploadData && (uploadData.fileUrl || (uploadData.data && uploadData.data.fileUrl));
  
  if (!isSuccess || !fileUrl) {
    const detail = uploadData ? (uploadData.error || uploadData.message || JSON.stringify(uploadData)) : 'Empty response';
    throw new Error("Upload failed: " + detail);
  }

  return fileUrl;
}

window.generateSingle = async function(tenantName) {
  const billIndex = appData[activeTab].findIndex(b => b.tenantName === tenantName);
  if (billIndex === -1) return;
  const bill = appData[activeTab][billIndex];
  
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Generating PDF for ${bill.tenantName}...`;

  try {
    const doc = await generatePDFDoc(bill);
    document.getElementById('loadingText').textContent = `Uploading to Drive...`;
    bill.pdfUrl = await uploadPdfToDrive(bill, doc);

    bill.status = 'Generated';
    renderSpreadsheet(); 
    await syncToDatabase(bill, activeTab, 'update');
  } catch (error) {
    console.error(error);
    alert("Error generating PDF: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

window.regenerateSingle = async function(tenantName) {
  const billIndex = appData[activeTab].findIndex(b => b.tenantName === tenantName);
  if (billIndex === -1) return;
  const bill = appData[activeTab][billIndex];
  
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Regenerating PDF for ${bill.tenantName}...`;

  try {
    const doc = await generatePDFDoc(bill);
    document.getElementById('loadingText').textContent = `Re-uploading to Drive...`;
    bill.pdfUrl = await uploadPdfToDrive(bill, doc);
    
    renderSpreadsheet();
    await syncToDatabase(bill, activeTab, 'update');
  } catch (error) {
    console.error(error);
    alert("Error regenerating PDF: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

document.getElementById('batchGenerateBtn').addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");
  
  const pendingBills = appData[activeTab].filter(b => b.status === 'Pending');
  
  if (pendingBills.length === 0) {
    return alert("There are no Pending bills in this tab to generate.");
  }

  const confirmGenerate = confirm(`Are you sure you want to generate PDFs for all ${pendingBills.length} Pending bills in "${activeTab}"?`);
  if (!confirmGenerate) return;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');

  try {
    for (let i = 0; i < pendingBills.length; i++) {
      const bill = pendingBills[i];
      document.getElementById('loadingText').textContent = `Generating (${i + 1}/${pendingBills.length}): ${bill.tenantName}...`;
      
      const doc = await generatePDFDoc(bill);
      document.getElementById('loadingText').textContent = `Uploading (${i + 1}/${pendingBills.length}): ${bill.tenantName}...`;
      bill.pdfUrl = await uploadPdfToDrive(bill, doc);

      bill.status = 'Generated';
      renderSpreadsheet(); 
      await syncToDatabase(bill, activeTab, 'update');
    }
    
    alert(`Successfully generated and uploaded ${pendingBills.length} bills.`);
  } catch (error) {
    console.error(error);
    alert("An error occurred during batch generation: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

document.getElementById('batchDownloadZipBtn').addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");

  const generatedBills = appData[activeTab].filter(b => b.pdfUrl && b.pdfUrl.trim() !== '');
  
  if (generatedBills.length === 0) {
    return alert("There are no generated PDFs in this tab to download. Please generate them first.");
  }

  const confirmZip = confirm(`Are you sure you want to download a ZIP file containing ${generatedBills.length} generated PDFs?`);
  if (!confirmZip) return;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Preparing ZIP...`;

  try {
    const zip = new JSZip();
    let count = 0;

    for (let i = 0; i < generatedBills.length; i++) {
      const bill = generatedBills[i];
      
      try {
        document.getElementById('loadingText').textContent = `Downloading (${i + 1}/${generatedBills.length}): ${bill.tenantName}...`;
        
        const response = await fetch(bill.pdfUrl);
        if (!response.ok) throw new Error("Network response was not ok");
        
        const blob = await response.blob();
        const fileName = `Water Bill - ${bill.tenantName}.pdf`;
        const propertyName = bill.property || activeTab;
        const monthName = getMonthFolderName(bill.billingDate);
        const tenantFolder = bill.tenantName || "Unknown Tenant";
        const folderPath = `${propertyName}/${monthName}/${tenantFolder}/`;
        
        zip.file(folderPath + fileName, blob);
        count++;
      } catch (err) {
        console.warn(`Could not download PDF for ${bill.tenantName}:`, err);
      }
    }

    if (count === 0) throw new Error("Could not download any PDF files.");

    document.getElementById('loadingText').textContent = `Zipping files...`;
    const content = await zip.generateAsync({ type: "blob" });
    
    const zipFileName = `Water Bills - ${activeTab} - ${new Date().toISOString().split('T')[0]}.zip`;
    
    const a = document.createElement("a");
    a.href = URL.createObjectURL(content);
    a.download = zipFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);

  } catch (error) {
    console.error(error);
    alert("Error creating ZIP: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

window.markAsSent = async function(tenantName) {
  const billIndex = appData[activeTab].findIndex(b => b.tenantName === tenantName);
  if (billIndex === -1) return;
  const bill = appData[activeTab][billIndex];
  bill.status = 'Sent';
  renderSpreadsheet();
  await syncToDatabase(bill, activeTab, 'update');
};

document.getElementById('clearCurrentTabBtn').addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");
  
  const confirmClear = confirm(`Are you sure you want to clear ALL data in the "${activeTab}" tab? This will wipe Google Sheets data.`);
  
  if (confirmClear) {
    document.getElementById('loadingOverlay').classList.remove('hidden');
    document.getElementById('loadingOverlay').classList.add('flex');
    document.getElementById('loadingText').textContent = `Clearing tab ${activeTab}...`;

    try {
      const res = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'clearTab', tabName: activeTab })
      });
      
      const resData = await res.json();
      
      if (resData.success) {
        appData[activeTab] = [];
        renderSpreadsheet();
        alert(`Tab "${activeTab}" cleared.`);
      } else {
        alert("Failed to clear: " + (resData.error || "Unknown"));
      }
    } catch (e) {
      alert("Error clearing tab.");
    } finally {
      document.getElementById('loadingOverlay').classList.add('hidden');
      document.getElementById('loadingOverlay').classList.remove('flex');
    }
  }
});

document.getElementById('addBillFormBtn').addEventListener('click', async () => {
  const tenantName = document.getElementById('inputTenant').value;
  if (!tenantName) return alert("Enter Tenant Name");

  const property = document.getElementById('inputProperty').value;
  if (!PROPERTIES.includes(property)) return alert("Please select a valid property.");

  const targetTab = property;
  const unitNo = document.getElementById('inputUnitNo').value;
  const prevReading = Number(document.getElementById('inputPrevReading').value) || 0;
  const presentReading = Number(document.getElementById('inputPresentReading').value) || 0;
  const consumption = calculateConsumption(prevReading, presentReading);
  const rate = Number(document.getElementById('inputRate').value) || 0;
  const balance = Number(document.getElementById('inputBalance').value) || 0;
  const penalty = Number(document.getElementById('inputPenalty').value) || 0;
  const billingDate = document.getElementById('inputBillingDate').value || new Date().toISOString().split('T')[0];

  const total = consumption * rate;

  if (!appData[targetTab]) appData[targetTab] = [];
  if (appData[targetTab].some(b => b.tenantName === tenantName)) {
    return alert(`A bill with this tenant name already exists in "${targetTab}".`);
  }

  const newBill = {
    tenantName: tenantName,
    property: property,
    unitNo: unitNo,
    prevReading: prevReading,
    presentReading: presentReading,
    consumption: consumption,
    rate: rate,
    balance: balance,
    penalty: penalty,
    total: total,
    billingDate: billingDate,
    status: 'Pending',
    pdfUrl: '',
    dateSent: ''
  };

  appData[targetTab].push(newBill);
  document.getElementById('inputTenant').value = '';

  activeTab = targetTab;
  renderTabs();
  renderSpreadsheet();

  await syncToDatabase(newBill, targetTab, 'save');
});

[
  ['inputPrevReading', 'inputPresentReading', 'inputConsumption'],
  ['editPrevReading', 'editPresentReading', 'editConsumption']
].forEach(([previousId, presentId, consumptionId]) => {
  const updateConsumption = () => {
    const previous = document.getElementById(previousId).value;
    const present = document.getElementById(presentId).value;
    document.getElementById(consumptionId).value = calculateConsumption(previous, present);
  };
  document.getElementById(previousId).addEventListener('input', updateConsumption);
  document.getElementById(presentId).addEventListener('input', updateConsumption);
});

// ================= EXCEL IMPORT =================
document.getElementById('excelInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = "Reading Excel File...";

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      workbookCache = XLSX.read(data, { type: 'array', cellDates: true });
      
      const sheetSelect = document.getElementById('sheetSelect');
      sheetSelect.innerHTML = '';
      workbookCache.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelect.appendChild(option);
      });
      document.getElementById('sheetSelectionArea').classList.remove('hidden');
    } catch (error) {
      alert("Error reading file.");
    } finally {
      document.getElementById('loadingOverlay').classList.add('hidden');
      document.getElementById('loadingOverlay').classList.remove('flex');
      document.getElementById('excelInput').value = ''; 
    }
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById('confirmLoadSheetBtn').addEventListener('click', async () => {
  if (!workbookCache) return;
  const selectedSheetName = document.getElementById('sheetSelect').value;
  const selectedProperty = document.getElementById('importPropertySelect').value;
  
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  
  try {
    const worksheet = workbookCache.Sheets[selectedSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { raw: true });
    if (json.length === 0) return alert("Sheet empty.");

    let importedCount = 0;

    for (let i = 0; i < json.length; i++) {
      document.getElementById('loadingText').textContent = `Uploading Data... (${i + 1}/${json.length})`;
      
      const row = json[i];
      const rowKeys = Object.keys(row).reduce((acc, key) => {
        acc[key.trim().toUpperCase()] = row[key];
        return acc;
      }, {});

      const rawUnit = (rowKeys['UNIT'] || '').toString().trim();
      if (!rawUnit) continue; 

      const prevReading = Number(rowKeys['PREVIOUS'] || 0);
      const presentReading = getImportedReading(rowKeys, ['CURRENT', 'CURRENT READING', 'PRESENT', 'PRESENT READING', 'METER PRESENT', 'METER PRESENT READING', 'UPON']);
      const consumption = calculateConsumption(prevReading, presentReading);
      const rate = Number(rowKeys['PER CUBIC'] || 0);
      const penalty = Number(rowKeys['PENALTY'] || 0);
      const total = consumption * rate;

      const tenantName = rawUnit;

      if (!appData[selectedProperty]) appData[selectedProperty] = [];
      if (appData[selectedProperty].some(b => b.tenantName === tenantName)) {
        console.warn(`Skipping duplicate unit "${tenantName}" in "${selectedProperty}"`);
        continue;
      }

      const newBill = {
        tenantName: tenantName,
        property: selectedProperty,
        unitNo: rawUnit,
        prevReading: prevReading,
        presentReading: presentReading,
        consumption: consumption,
        rate: rate,
        balance: 0, 
        penalty: penalty,
        total: total,
        billingDate: new Date().toISOString().split('T')[0], 
        status: 'Pending',
        pdfUrl: '',
        dateSent: ''
      };

      appData[selectedProperty].push(newBill);
      await syncToDatabase(newBill, selectedProperty, 'save');
      importedCount++;
    }

    activeTab = selectedProperty;
    document.getElementById('searchInput').value = ''; 
    document.getElementById('statusFilter').value = 'ALL'; 
    renderTabs(); 
    renderSpreadsheet();
    document.getElementById('sheetSelectionArea').classList.add('hidden');
    workbookCache = null;
    alert(`Successfully imported ${importedCount} rows to ${selectedProperty}.`);
  } catch (error) {
    console.error(error);
    alert("Error loading sheet.");
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

// ================= IMPORT VIA LINK (RELIABLE METHOD) =================
document.getElementById('importLinkBtn').addEventListener('click', async () => {
  const url = prompt("Enter the Google Sheets Link (Must be shared as 'Anyone with the link'):");
  if (!url) return;

  // Extract Spreadsheet ID and GID
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) return alert("Invalid Google Sheets URL. Please ensure it looks like 'https://docs.google.com/spreadsheets/d/YOUR_ID/...' ");
  
  const spreadsheetId = match[1];
  
  // Try to extract GID
  let gid = '0'; // Default to first sheet
  const gidMatch = url.match(/[#&]gid=([0-9]+)/);
  if (gidMatch) {
    gid = gidMatch[1];
  }

  // Ask for property
  const propertyPrompt = `Assign imported data to which property?\n\nOptions: ${PROPERTIES.join(', ')}`;
  const selectedProperty = prompt(propertyPrompt, DEFAULT_PROPERTY);
  
  if (!selectedProperty || !PROPERTIES.includes(selectedProperty.trim())) {
    return alert("Invalid property selected. Import cancelled.");
  }

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = "Fetching data from Google Sheets...";

  try {
    // Use the Google Visualization API (Bypasses CORS securely)
    const queryUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
    
    const response = await fetch(queryUrl);
    if (!response.ok) throw new Error("Failed to fetch. Ensure the sheet is shared as 'Anyone with the link can view'.");
    
    const csvText = await response.text();
    
    // Parse CSV using XLSX
    const workbook = XLSX.read(csvText, { type: 'string' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { raw: true });

    if (json.length === 0) throw new Error("The imported sheet is empty.");

    let importedCount = 0;
    const targetProp = selectedProperty.trim();

    for (let i = 0; i < json.length; i++) {
      document.getElementById('loadingText').textContent = `Importing row ${i + 1} of ${json.length}...`;
      
      const row = json[i];
      const rowKeys = Object.keys(row).reduce((acc, key) => {
        acc[key.trim().toUpperCase()] = row[key];
        return acc;
      }, {});

      const rawUnit = (rowKeys['UNIT'] || '').toString().trim();
      if (!rawUnit) continue;

      const prevReading = Number(rowKeys['PREVIOUS'] || 0);
      const presentReading = getImportedReading(rowKeys, ['CURRENT', 'CURRENT READING', 'PRESENT', 'PRESENT READING', 'METER PRESENT', 'METER PRESENT READING', 'UPON']);
      const consumption = calculateConsumption(prevReading, presentReading);
      const rate = Number(rowKeys['PER CUBIC'] || 0);
      const penalty = Number(rowKeys['PENALTY'] || 0);
      const total = consumption * rate;

      const tenantName = rawUnit;

      if (!appData[targetProp]) appData[targetProp] = [];
      if (appData[targetProp].some(b => b.tenantName === tenantName)) {
        continue; // Skip duplicates
      }

      const newBill = {
        tenantName: tenantName,
        property: targetProp,
        unitNo: rawUnit,
        prevReading: prevReading,
        presentReading: presentReading,
        consumption: consumption,
        rate: rate,
        balance: 0,
        penalty: penalty,
        total: total,
        billingDate: new Date().toISOString().split('T')[0],
        status: 'Pending',
        pdfUrl: '',
        dateSent: ''
      };

      appData[targetProp].push(newBill);
      await syncToDatabase(newBill, targetProp, 'save');  
      importedCount++;
    }

    activeTab = targetProp;
    document.getElementById('searchInput').value = ''; 
    document.getElementById('statusFilter').value = 'ALL'; 
    renderTabs(); 
    renderSpreadsheet();
    
    alert(`Successfully imported ${importedCount} rows from Google Sheets to ${targetProp}.`);

  } catch (error) {
    console.error(error);
    alert("Import Error: " + error.message + "\n\nPlease make sure the Google Sheet is shared with 'Anyone with the link can view'.");
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

document.getElementById('exportDriveBtn').addEventListener('click', () => {
  alert("Export to Drive functionality would use the same Apps Script endpoint. Ensure ROOT_FOLDER_ID is set correctly.");
});

const TAB_EXPORT_HEADERS = [
  'Tenant Name', 'Property', 'Unit No', 'Billing Date', 'Previous', 'Current',
  'Consumption (cu.m)', 'Rate', 'Balance', 'Penalty', 'Total', 'Status',
  'PDF URL', 'Date Sent'
];

function getActiveTabBills() {
  if (!activeTab || !Array.isArray(appData[activeTab])) {
    alert('There is no active property tab to export.');
    return null;
  }
  return appData[activeTab];
}

function getTabExportRows(bills) {
  const rows = bills.map(bill => [
    bill.tenantName || '',
    bill.property || activeTab,
    bill.unitNo || '',
    toDateInputValue(bill.billingDate),
    Number(bill.prevReading) || 0,
    Number(bill.presentReading) || 0,
    Number(bill.consumption) || 0,
    Number(bill.rate) || 0,
    Number(bill.balance) || 0,
    Number(bill.penalty) || 0,
    Number(bill.total) || 0,
    bill.status || '',
    bill.pdfUrl || '',
    toDateInputValue(bill.dateSent)
  ]);
  const totalConsumption = bills.reduce((sum, bill) => sum + (Number(bill.consumption) || 0), 0);
  const totalAmount = bills.reduce((sum, bill) => sum + (Number(bill.total) || 0), 0);
  const totalsRow = Array(TAB_EXPORT_HEADERS.length).fill('');
  totalsRow[0] = 'TOTAL (ALL BILLS)';
  totalsRow[6] = totalConsumption;
  totalsRow[10] = totalAmount;
  return [TAB_EXPORT_HEADERS, ...rows, totalsRow];
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function printActiveTab() {
  const bills = getActiveTabBills();
  if (!bills) return;

  const rows = bills.map(bill => {
    const values = [
      bill.tenantName || '',
      bill.property || activeTab,
      bill.unitNo || '',
      formatDisplayDate(bill.billingDate),
      Number(bill.prevReading) || 0,
      Number(bill.presentReading) || 0,
      `${Number(bill.consumption) || 0} cu.m`,
      formatPeso(bill.rate),
      formatPeso(bill.balance),
      formatPeso(bill.penalty),
      formatPeso(bill.total),
      bill.status || '',
      bill.pdfUrl || '',
      formatDisplayDate(bill.dateSent)
    ];
    return `<tr>${values.map(value => `<td>${escapeHtml(value)}</td>`).join('')}</tr>`;
  }).join('');
  const totalConsumption = bills.reduce((sum, bill) => sum + (Number(bill.consumption) || 0), 0);
  const totalAmount = bills.reduce((sum, bill) => sum + (Number(bill.total) || 0), 0);
  const totals = Array(TAB_EXPORT_HEADERS.length).fill('');
  totals[0] = 'TOTAL (ALL BILLS)';
  totals[6] = `${totalConsumption} cu.m`;
  totals[10] = formatPeso(totalAmount);

  document.getElementById('printReport').innerHTML = `
    <h1>${escapeHtml(activeTab)} - Water Bill Records</h1>
    <p>${bills.length} records | Printed ${escapeHtml(new Date().toLocaleString())}</p>
    <table>
      <thead><tr>${TAB_EXPORT_HEADERS.map(header => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr>${totals.map(value => `<th>${escapeHtml(value)}</th>`).join('')}</tr></tfoot>
    </table>`;
  window.print();
}

document.getElementById('printTabBtn').addEventListener('click', printActiveTab);
document.getElementById('downloadTabBtn').addEventListener('click', () => {
  const bills = getActiveTabBills();
  if (!bills) return;
  if (typeof XLSX === 'undefined') {
    alert('Excel export is unavailable because the spreadsheet library did not load.');
    return;
  }

  const worksheet = XLSX.utils.aoa_to_sheet(getTabExportRows(bills));
  worksheet['!cols'] = [
    { wch: 24 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
    { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 },
    { wch: 36 }, { wch: 14 }
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Water Bills');
  const safeTabName = activeTab.replace(/[<>:"/\\|?*]/g, '_');
  XLSX.writeFile(workbook, `${safeTabName}-water-bills.xlsx`);
});

window.onload = initDatabase;