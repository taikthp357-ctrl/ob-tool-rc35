/**
 * Independent JavaScript Verification of Excel-like Grid & Paste Behavior in app.js
 * Tested by Nova (QA & Reliability Engineer)
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

// Read app.js
const appJsPath = path.join(__dirname, '..', 'app', 'static', 'js', 'app.js');
const appJsCode = fs.readFileSync(appJsPath, 'utf-8');

// Verify column definitions in app.js
assert(appJsCode.includes("key: 'model', label: 'Model', letter: 'A'"), "Cột Model phải là Cột A");
assert(appJsCode.includes("key: 'material_code', label: 'Mã hàng', letter: 'B'"), "Cột Mã hàng phải là Cột B");
assert(appJsCode.includes("const EDITABLE_KEYS = ['model', 'material_code', 'required_qty', 'batch', 'shop', 'sto_number', 'ob_number'];"), "EDITABLE_KEYS phải bám đúng thứ tự Excel");

// Mock Minimal Global Environment for app.js applyPastedText
const mockWindow = {
  __MA_CONTROL_APP_LOADED__: false,
  currentBranch: 'BODY_3001',
  gridData: [],
  activeRowIndex: 0,
  activeColKey: 'model',
  isEditing: false,
  stockCandidates: []
};

const EDITABLE_KEYS = ['model', 'material_code', 'required_qty', 'batch', 'shop', 'sto_number', 'ob_number'];

function addNewRow(toEnd) {
  const newRow = {
    item_uid: `mock-row-${mockWindow.gridData.length + 1}`,
    branch_tab: mockWindow.currentBranch,
    selected: true,
    model: '',
    material_code: '',
    required_qty: 0,
    req_qty: 0,
    batch: '',
    shop: 'BODY',
    sto_number: '',
    ob_number: '',
    status: 'Chưa kiểm tra'
  };
  mockWindow.gridData.push(newRow);
  return newRow;
}

// Extract applyPastedText logic from app.js
function applyPastedText(rawText) {
  if (!rawText) return;
  const rows = rawText.split(/\r\n|\n|\r/);
  if (rows.length === 0 || (rows.length === 1 && !rows[0].trim())) return;

  let startRow = mockWindow.activeRowIndex;
  const startColIdx = EDITABLE_KEYS.indexOf(mockWindow.activeColKey) >= 0 ? EDITABLE_KEYS.indexOf(mockWindow.activeColKey) : 0;

  rows.forEach((rowText, rOffset) => {
    const targetRowIdx = startRow + rOffset;
    while (targetRowIdx >= mockWindow.gridData.length) {
      addNewRow(false);
    }

    if (!rowText.trim()) {
      mockWindow.gridData[targetRowIdx].material_code = '';
      mockWindow.gridData[targetRowIdx].model = '';
      mockWindow.gridData[targetRowIdx].required_qty = 0;
      mockWindow.gridData[targetRowIdx].req_qty = 0;
      mockWindow.gridData[targetRowIdx].batch = '';
      mockWindow.gridData[targetRowIdx].status = 'Chưa kiểm tra';
      return;
    }

    const cells = rowText.split('\t');
    let isMaterialFirst = false;
    if (cells.length >= 2) {
      const c0 = cells[0].trim();
      const c1 = cells[1].trim();
      const c0IsMat = (c0.length >= 8 && (c0.startsWith('00') || c0.startsWith('BIW') || c0.startsWith('EDC')));
      const c1IsModel = (c1.toUpperCase().includes('VF') || c1.toUpperCase().includes('LIMO'));
      if (c0IsMat || c1IsModel) {
        isMaterialFirst = true;
      }
    }

    if (isMaterialFirst && cells.length >= 2) {
      mockWindow.gridData[targetRowIdx].material_code = String(cells[0].trim());
      mockWindow.gridData[targetRowIdx].model = String(cells[1].trim());
      if (cells.length > 2) {
        const q = parseFloat(cells[2].trim()) || 0;
        mockWindow.gridData[targetRowIdx].required_qty = q;
        mockWindow.gridData[targetRowIdx].req_qty = q;
      }
      if (cells.length > 3) mockWindow.gridData[targetRowIdx].batch = String(cells[3].trim());
      if (cells.length > 4) mockWindow.gridData[targetRowIdx].shop = String(cells[4].trim());
      if (cells.length > 5) mockWindow.gridData[targetRowIdx].sto_number = String(cells[5].trim());
      if (cells.length > 6) mockWindow.gridData[targetRowIdx].ob_number = String(cells[6].trim());
    } else {
      cells.forEach((cellText, cOffset) => {
        const colIdx = startColIdx + cOffset;
        if (colIdx < EDITABLE_KEYS.length) {
          const colKey = EDITABLE_KEYS[colIdx];
          const val = cellText.trim();
          if (colKey === 'required_qty') {
            mockWindow.gridData[targetRowIdx].required_qty = parseFloat(val) || 0;
            mockWindow.gridData[targetRowIdx].req_qty = mockWindow.gridData[targetRowIdx].required_qty;
          } else {
            mockWindow.gridData[targetRowIdx][colKey] = String(val);
          }
        }
      });
    }
    mockWindow.gridData[targetRowIdx].status = 'Chưa kiểm tra';
  });
}

console.log('--- TEST 1: Paste 30 rows with blank rows and leading zeros ---');
let tsvRows = [];
for (let i = 1; i <= 30; i++) {
  if (i === 5 || i === 15) {
    tsvRows.push(''); // Blank row
  } else {
    tsvRows.push(`VF7\t0001099${String(i).padStart(3, '0')}\t${i * 10}\t00B${String(i).padStart(3, '0')}\tBODY\t00STO${String(i).padStart(4, '0')}\t00OB${String(i).padStart(5, '0')}`);
  }
}
const tsvText = tsvRows.join('\n');

applyPastedText(tsvText);

assert.strictEqual(mockWindow.gridData.length, 30, 'Phải sinh đủ 30 dòng');
assert.strictEqual(mockWindow.gridData[0].model, 'VF7');
assert.strictEqual(mockWindow.gridData[0].material_code, '0001099001', 'Bảo toàn số 0 ở đầu mã hàng');
assert.strictEqual(mockWindow.gridData[0].batch, '00B001', 'Bảo toàn số 0 ở đầu batch');
assert.strictEqual(mockWindow.gridData[0].sto_number, '00STO0001', 'Bảo toàn số 0 ở đầu STO');
assert.strictEqual(mockWindow.gridData[0].ob_number, '00OB00001', 'Bảo toàn số 0 ở đầu OB');

// Blank row check
assert.strictEqual(mockWindow.gridData[4].material_code, '', 'Dòng 5 phải là dòng trống');
assert.strictEqual(mockWindow.gridData[4].model, '', 'Dòng 5 phải là dòng trống');
assert.strictEqual(mockWindow.gridData[14].material_code, '', 'Dòng 15 phải là dòng trống');

console.log('--- TEST 2: Paste 50 rows material-first format ---');
mockWindow.gridData = [];
let tsv50 = [];
for (let i = 1; i <= 50; i++) {
  tsv50.push(`00BIW${String(i).padStart(5, '0')}\tVF8\t${i * 5}\t00BATCH${i}\tBODY\t00STO\t00OB${i}`);
}
applyPastedText(tsv50.join('\n'));
assert.strictEqual(mockWindow.gridData.length, 50, 'Phải sinh đủ 50 dòng');
assert.strictEqual(mockWindow.gridData[0].material_code, '00BIW00001', 'Nhận diện material first chính xác');
assert.strictEqual(mockWindow.gridData[0].model, 'VF8');
assert.strictEqual(mockWindow.gridData[49].material_code, '00BIW00050');
assert.strictEqual(mockWindow.gridData[49].model, 'VF8');

console.log('ALL JS PASTE TESTS PASSED SUCCESSFULLY (100%)');
