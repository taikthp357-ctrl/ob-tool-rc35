// MA CONTROL — JavaScript Excel Grid & 5-Step Workflow Engine (app.js)
// Hỗ trợ Paste Ctrl+V đa ô/dòng giữ số 0 đầu, Di chuyển Tab/Enter, Inline editing,
// Cổng kiểm soát CF thủ công (POST /api/pick-orders/{id}/verify-cf), và tạo BBBG V2.2.14.

if (!window.__MA_CONTROL_APP_LOADED__) {
  window.__MA_CONTROL_APP_LOADED__ = true;

  window.currentBranch = 'BODY_3001';
  window.gridData = [];
  window.activeRowIndex = 0;
  window.activeColKey = 'material_code';
  window.isEditing = false;
  window.stockCandidates = [];

  const COLUMNS = [
    { key: 'model', label: 'Model', letter: 'A', editable: true, width: '110px' },
    { key: 'material_code', label: 'Mã hàng', letter: 'B', editable: true, width: '160px' },
    { key: 'required_qty', label: 'SL yêu cầu', letter: 'C', editable: true, width: '100px' },
    { key: 'batch', label: 'Batch', letter: 'D', editable: true, width: '110px' },
    { key: 'shop', label: 'Shop', letter: 'E', editable: true, width: '90px' },
    { key: 'sto_number', label: 'STO', letter: 'F', editable: true, width: '120px' },
    { key: 'ob_number', label: 'OB', letter: 'G', editable: true, width: '120px' },
    { key: 'status', label: 'Trạng thái', letter: 'H', editable: false, width: '140px' }
  ];

  const EDITABLE_KEYS = ['model', 'material_code', 'required_qty', 'batch', 'shop', 'sto_number', 'ob_number'];

  document.addEventListener('DOMContentLoaded', () => {
    loadGridData(window.currentBranch);
    loadPickOrders();
    setupEventListeners();
  });

  window.loadGridData = async function(branch) {
    try {
      const res = await fetch(`/api/ob/items?branch_tab=${branch}`);
      if (res.ok) {
        const items = await res.json();
        window.gridData = items.map(it => ({
          ...it,
          required_qty: it.required_qty !== undefined ? it.required_qty : (it.req_qty || 0),
          selected: it.selected !== undefined ? !!it.selected : true
        }));
      } else {
        window.gridData = [];
      }
    } catch (err) {
      console.error('Error fetching grid data:', err);
      window.gridData = [];
    }
    if (window.gridData.length === 0) {
      addNewRow(false);
      addNewRow(false);
      addNewRow(false);
    }
    renderGrid();
  };

  window.getFilteredGridData = function() {
    const q = (document.getElementById('input-search')?.value || '').trim().toLowerCase();
    const modelFilter = document.getElementById('select-model')?.value || '';
    const statusFilter = document.getElementById('select-status')?.value || '';

    return window.gridData.map((row, idx) => ({ row, idx })).filter(({ row }) => {
      if (q) {
        const mat = String(row.material_code || '').toLowerCase();
        const ob = String(row.ob_number || '').toLowerCase();
        const sto = String(row.sto_number || '').toLowerCase();
        if (!mat.includes(q) && !ob.includes(q) && !sto.includes(q)) return false;
      }
      if (modelFilter && !String(row.model || '').toUpperCase().includes(modelFilter.toUpperCase())) {
        return false;
      }
      if (statusFilter) {
        const norm = normalizeStatusKey(row.status);
        if (norm !== statusFilter) return false;
      }
      return true;
    });
  };

  function normalizeStatusKey(status) {
    const s = String(status || '').trim();
    if (s === 'READY' || s === 'Sẵn sàng điền' || s === 'READY_FOR_PICK' || s === 'Đủ điều kiện tạo pick' || s === 'ELIGIBLE_FOR_PICK') return 'READY_FOR_PICK';
    if (s === 'INSUFFICIENT' || s === 'Thiếu SL OB' || s === 'LOCKED_SHORTAGE' || s === 'LOCKED' || s === 'SHORTAGE' || s === 'Đang khóa / Thiếu tồn' || s === 'LOCKED_OR_INSUFFICIENT') return 'LOCKED_OR_INSUFFICIENT';
    if (s === 'WAITING_PICK' || s === 'Chờ pick / CF' || s === 'WAITING_CF' || s === 'Chờ CF' || s === 'PENDING_CF') return 'PENDING_CF';
    if (s === 'PARTIAL_CF' || s === 'CF một phần') return 'PARTIAL_CF';
    if (s === 'CONFIRMED' || s === 'Đã CF' || s === 'CF_COMPLETED' || s === 'USER_CONFIRMED' || s === 'Đã CF đầy đủ' || s === 'CONFIRMED_CF' || s === 'Đã xác nhận CF đầy đủ') return 'CONFIRMED_CF';
    if (s === 'BBBG_READY' || s === 'Đủ điều kiện BBBG' || s === 'READY_FOR_BBBG' || s === 'Đã lập BBBG' || s === 'HANDED_OVER' || s === 'ELIGIBLE_BBBG' || s === 'Đủ điều kiện tạo/đối soát BBBG') return 'ELIGIBLE_BBBG';
    if (s === 'INVALID' || s === 'CANCELLED' || s === 'INVALID_CANCELLED' || s === 'Không hợp lệ / Đã hủy' || s === 'Không hợp lệ' || s === 'INVALID_OR_CANCELLED') return 'INVALID_OR_CANCELLED';
    return 'UNCHECKED';
  }

  window.renderGrid = function() {
    const tbody = document.getElementById('grid-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const visibleRows = window.getFilteredGridData();
    visibleRows.forEach(({ row, idx: rIdx }, displayIdx) => {
      const tr = document.createElement('tr');
      tr.dataset.rowIndex = rIdx;

      const tdNum = document.createElement('td');
      tdNum.className = 'row-num';
      tdNum.textContent = displayIdx + 1;
      tr.appendChild(tdNum);

      COLUMNS.forEach(col => {
        const td = document.createElement('td');
        td.dataset.colKey = col.key;
        td.dataset.rowIndex = rIdx;

        if (rIdx === window.activeRowIndex && col.key === window.activeColKey) {
          td.classList.add('cell-active');
        }

        if (col.key === 'status') {
          td.innerHTML = renderStatusChip(row.status);
        } else {
          td.textContent = row[col.key] !== undefined && row[col.key] !== null ? row[col.key] : '';
        }

        td.addEventListener('click', () => selectCell(rIdx, col.key));
        td.addEventListener('dblclick', () => {
          if (col.editable) startInlineEdit(td, rIdx, col.key);
        });

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    updateSummary();
  };

  function renderStatusChip(status) {
    const s = String(status || '').trim();
    if (s === 'Sẵn sàng điền') {
      return `<span class="chip chip-ready">Sẵn sàng điền</span>`;
    }
    if (s === 'Thiếu SL OB') {
      return `<span class="chip chip-insufficient">⚠ Thiếu SL OB</span>`;
    }
    if (s === 'Chờ pick / CF') {
      return `<span class="chip chip-waiting-pick">Chờ pick / CF</span>`;
    }
    const norm = normalizeStatusKey(status);
    switch (norm) {
      case 'READY_FOR_PICK':
        return `<span class="chip chip-ready">Đủ điều kiện tạo pick</span>`;
      case 'LOCKED_OR_INSUFFICIENT':
        return `<span class="chip chip-insufficient">⚠ Đang khóa / Thiếu tồn</span>`;
      case 'PENDING_CF':
        return `<span class="chip chip-waiting-pick">Chờ CF</span>`;
      case 'PARTIAL_CF':
        return `<span class="chip chip-waiting-pick" style="background:#fef3c7;color:#92400e;border:1px solid #fde68a;">CF một phần</span>`;
      case 'CONFIRMED_CF':
        return `<span class="chip chip-ready" style="background:#dcfce7;color:#15803d;border:1px solid #86efac;">Đã xác nhận CF</span>`;
      case 'ELIGIBLE_BBBG':
        return `<span class="chip chip-ready" style="background:#e0e7ff;color:#4338ca;border:1px solid #c7d2fe;">Đủ điều kiện BBBG</span>`;
      case 'INVALID_OR_CANCELLED':
        return `<span class="chip chip-insufficient">Không hợp lệ / Đã hủy</span>`;
      case 'UNCHECKED':
      default:
        return `<span class="chip chip-pending">Chưa kiểm tra</span>`;
    }
  }

  window.selectCell = function(rIdx, colKey) {
    if (window.isEditing) return;
    window.activeRowIndex = Math.max(0, Math.min(rIdx, window.gridData.length - 1));
    window.activeColKey = colKey;

    document.querySelectorAll('.cell-active').forEach(el => el.classList.remove('cell-active'));
    const target = document.querySelector(`td[data-row-index="${window.activeRowIndex}"][data-col-key="${window.activeColKey}"]`);
    if (target) {
      target.classList.add('cell-active');
      target.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  };

  function startInlineEdit(td, rIdx, colKey) {
    if (window.isEditing) return;
    window.isEditing = true;

    const currentVal = window.gridData[rIdx][colKey] || '';
    td.classList.add('cell-edit-mode');
    td.innerHTML = '';

    const input = document.createElement('input');
    input.type = colKey === 'required_qty' ? 'number' : 'text';
    input.value = currentVal;
    td.appendChild(input);
    input.focus();
    input.select();

    const finishEdit = () => {
      if (!window.isEditing) return;
      const val = input.value.trim();
      if (colKey === 'required_qty') {
        window.gridData[rIdx].required_qty = parseFloat(val) || 0;
        window.gridData[rIdx].req_qty = window.gridData[rIdx].required_qty;
      } else {
        // Preserve leading zeroes as string!
        window.gridData[rIdx][colKey] = String(val);
      }
      window.gridData[rIdx].status = 'Chưa kiểm tra';
      window.isEditing = false;
      td.classList.remove('cell-edit-mode');
      renderGrid();
    };

    input.addEventListener('blur', finishEdit);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        finishEdit();
        selectCell(rIdx + 1, colKey);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        finishEdit();
        navigateNextCell(rIdx, colKey, e.shiftKey);
      } else if (e.key === 'Escape') {
        window.isEditing = false;
        td.classList.remove('cell-edit-mode');
        renderGrid();
      }
    });
  }

  function navigateNextCell(rIdx, colKey, isShift) {
    const editableCols = EDITABLE_KEYS;
    let idx = editableCols.indexOf(colKey);
    if (!isShift) {
      if (idx < editableCols.length - 1) {
        selectCell(rIdx, editableCols[idx + 1]);
      } else if (rIdx < window.gridData.length - 1) {
        selectCell(rIdx + 1, editableCols[0]);
      }
    } else {
      if (idx > 0) {
        selectCell(rIdx, editableCols[idx - 1]);
      } else if (rIdx > 0) {
        selectCell(rIdx - 1, editableCols[editableCols.length - 1]);
      }
    }
  }

  window.applyPastedText = function(rawText) {
    if (!rawText) return;
    const rows = rawText.split(/\r\n|\n|\r/);
    if (rows.length === 0 || (rows.length === 1 && !rows[0].trim())) return;

    let startRow = window.activeRowIndex;
    const startColIdx = EDITABLE_KEYS.indexOf(window.activeColKey) >= 0 ? EDITABLE_KEYS.indexOf(window.activeColKey) : 0;

    rows.forEach((rowText, rOffset) => {
      const targetRowIdx = startRow + rOffset;
      while (targetRowIdx >= window.gridData.length) {
        addNewRow(false);
      }

      if (!rowText.trim()) {
        window.gridData[targetRowIdx].material_code = '';
        window.gridData[targetRowIdx].model = '';
        window.gridData[targetRowIdx].required_qty = 0;
        window.gridData[targetRowIdx].req_qty = 0;
        window.gridData[targetRowIdx].batch = '';
        window.gridData[targetRowIdx].status = 'Chưa kiểm tra';
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
        window.gridData[targetRowIdx].material_code = String(cells[0].trim());
        window.gridData[targetRowIdx].model = String(cells[1].trim());
        if (cells.length > 2) {
          const q = parseFloat(cells[2].trim()) || 0;
          window.gridData[targetRowIdx].required_qty = q;
          window.gridData[targetRowIdx].req_qty = q;
        }
        if (cells.length > 3) window.gridData[targetRowIdx].batch = String(cells[3].trim());
        if (cells.length > 4) window.gridData[targetRowIdx].shop = String(cells[4].trim());
        if (cells.length > 5) window.gridData[targetRowIdx].sto_number = String(cells[5].trim());
        if (cells.length > 6) window.gridData[targetRowIdx].ob_number = String(cells[6].trim());
      } else {
      cells.forEach((cellText, cOffset) => {
        const colIdx = startColIdx + cOffset;
        if (colIdx < EDITABLE_KEYS.length) {
          const colKey = EDITABLE_KEYS[colIdx];
          const val = cellText.trim();
          if (colKey === 'required_qty') {
            window.gridData[targetRowIdx].required_qty = parseFloat(val) || 0;
            window.gridData[targetRowIdx].req_qty = window.gridData[targetRowIdx].required_qty;
          } else {
            // Keep exact text with leading zeros
            window.gridData[targetRowIdx][colKey] = String(val);
          }
        }
      });
      }
      window.gridData[targetRowIdx].status = 'Chưa kiểm tra';
    });

    renderGrid();
    selectCell(startRow, window.activeColKey);
  };

  function handlePaste(e) {
    if (window.isEditing || document.activeElement?.tagName === 'TEXTAREA' || document.activeElement?.tagName === 'INPUT') {
      return;
    }
    const clipboardData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipboardData) return;
    e.preventDefault();
    window.applyPastedText(clipboardData);
  }

  window.addNewRow = function(reRender = true) {
    const newRow = {
      id: `row-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      branch_tab: window.currentBranch,
      model: '',
      material_code: '',
      required_qty: 0,
      req_qty: 0,
      batch: '',
      shop: window.currentBranch.includes('BODY') ? 'BODY' : 'EDC',
      sto_number: '',
      ob_number: '',
      status: 'Chưa kiểm tra',
      selected: true
    };
    window.gridData.push(newRow);
    if (reRender) {
      renderGrid();
      selectCell(window.gridData.length - 1, 'model');
    }
  };

  window.deleteActiveRow = function() {
    if (window.gridData.length <= 1) {
      window.gridData[0] = {
        id: `row-${Date.now()}`,
        branch_tab: window.currentBranch,
        model: '', material_code: '', required_qty: 0, req_qty: 0,
        batch: '', shop: window.currentBranch.includes('BODY') ? 'BODY' : 'EDC',
        sto_number: '', ob_number: '', status: 'Chưa kiểm tra', selected: true
      };
    } else {
      window.gridData.splice(window.activeRowIndex, 1);
      if (window.activeRowIndex >= window.gridData.length) {
        window.activeRowIndex = window.gridData.length - 1;
      }
    }
    renderGrid();
  };

  function updateSummary() {
    const validRows = window.gridData.filter(r => String(r.material_code || '').trim().length > 0);
    const totalQty = validRows.reduce((acc, r) => acc + (parseFloat(r.required_qty || r.req_qty) || 0), 0);
    const summaryEl = document.getElementById('selection-summary');
    if (summaryEl) {
      summaryEl.textContent = `Số mã: ${validRows.length} • Tổng SL: ${totalQty.toLocaleString()} PCS`;
    }
  }

  window.loadPickOrders = async function() {
    try {
      const res = await fetch('/api/pick-orders');
      if (res.ok) {
        const orders = await res.json();
        renderPickOrders(orders);
      }
    } catch (err) {
      console.error('Error fetching pick orders:', err);
    }
  };

  function renderPickOrders(orders) {
    const tbody = document.getElementById('pick-orders-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const waitingCount = orders.filter(po => po.status === 'Chờ người dùng CF' || po.status === 'WAITING_USER_CF').length;
    const badgeEl = document.getElementById('badge-cf-count');
    if (badgeEl) badgeEl.textContent = waitingCount;

    if (orders.length === 0) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#94a3b8;padding:16px;">Chưa có đơn pick nào được tạo</td></tr>`;
      return;
    }

    orders.forEach((po, idx) => {
      const tr = document.createElement('tr');
      const code = po.pick_order_code || po.pick_order_no;
      const isWaiting = po.status === 'Chờ người dùng CF' || po.status === 'WAITING_USER_CF';
      const isConfirmed = po.status === 'Đã CF' || po.status === 'USER_CONFIRMED';
      const statusBadge = isWaiting
        ? `<span class="badge-cf-waiting">Chờ người dùng CF</span>`
        : `<span class="badge-cf-confirmed">${po.status === 'HANDED_OVER' ? 'Đã lập BBBG' : 'Đã CF'}</span>`;

      const actionButtons = isWaiting
        ? `<button class="btn btn-primary" style="padding:4px 10px;font-size:12px;" onclick="openCfModal('${code}')">Kiểm tra CF</button>`
        : (isConfirmed
            ? `<button class="btn btn-success" style="padding:4px 10px;font-size:12px;" onclick="openBbbgModal('${code}')">📄 Lập BBBG</button>`
            : `<span style="font-size:12px;color:#059669;font-weight:600;">✓ Hoàn tất</span>`);

      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><strong>${po.slip_no || code}</strong></td>
        <td>${po.ob_summary || '--'}</td>
        <td>${po.material_summary || '--'}</td>
        <td>${(po.total_qty || 0).toLocaleString()}</td>
        <td>${po.location_summary || 'KHO BODY'}</td>
        <td><span class="badge" style="background:#f1f5f9;color:#475569;font-size:11px;padding:3px 6px;border-radius:4px;border:1px solid #cbd5e1;">Do người làm</span></td>
        <td>${statusBadge}</td>
        <td>${actionButtons}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // CF Modal
  window.openCfModal = async function(pickOrderNo) {
    const modal = document.getElementById('cf-modal');
    document.getElementById('modal-pick-no').textContent = pickOrderNo;
    modal.dataset.pickOrderNo = pickOrderNo;
    modal.style.display = 'flex';
    let panel = document.getElementById('cf-actual-lines');
    if (!panel) {
      panel = document.createElement('div');
      panel.id = 'cf-actual-lines';
      modal.querySelector('.modal-body').appendChild(panel);
    }
    panel.textContent = 'Đang tải các dòng cần kiểm tra...';
    modal.dataset.cfLoaded = '';
    try {
      const response = await fetch(`/api/pick-orders/${encodeURIComponent(pickOrderNo)}`);
      if (!response.ok) throw new Error('Không tải được đơn pick');
      const order = await response.json();
      if (modal.dataset.pickOrderNo !== pickOrderNo) return;
      panel.textContent = '';
      const operatorLabel = document.createElement('label');
      operatorLabel.textContent = 'Người kiểm tra: ';
      const operator = document.createElement('input');
      operator.id = 'cf-operator';
      operator.type = 'text';
      operator.required = true;
      operatorLabel.appendChild(operator);
      panel.appendChild(operatorLabel);
      (order.items || []).forEach(line => {
        const label = document.createElement('label');
        label.style.display = 'block';
        label.textContent = `Dòng ${line.id} · ${line.material_code} · Batch ${line.batch || ''} · OB ${line.ob_number || ''} · Yêu cầu ${line.qty} — thực tế: `;
        const input = document.createElement('input');
        input.type = 'number';
        input.step = 'any';
        input.min = '0';
        input.required = true;
        input.className = 'cf-actual-qty';
        input.dataset.lineId = String(line.id);
        input.dataset.requested = String(line.qty);
        label.appendChild(input);
        panel.appendChild(label);
      });
      modal.dataset.cfLoaded = (order.items || []).length ? 'yes' : '';
    } catch (error) {
      panel.textContent = error.message;
    }
  };

  window.closeCfModal = function() {
    document.getElementById('cf-modal').style.display = 'none';
  };

  window.confirmPickOrder = async function() {
    const modal = document.getElementById('cf-modal');
    const pickOrderNo = modal.dataset.pickOrderNo;
    if (!pickOrderNo) return;
    if (modal.dataset.cfLoaded !== 'yes' || modal.dataset.cfSubmitting === 'yes') return;
    const operator = document.getElementById('cf-operator').value.trim();
    if (!operator) { alert('Nhập tên người đã kiểm tra thực tế.'); return; }
    const actualQuantities = {};
    for (const input of modal.querySelectorAll('.cf-actual-qty')) {
      const qty = Number(input.value);
      if (!input.value.trim() || !Number.isFinite(qty) || qty <= 0 || qty !== Number(input.dataset.requested)) {
        alert(`Dòng ${input.dataset.lineId}: nhập số lượng thực tế lớn hơn 0 và khớp yêu cầu trước khi CF.`);
        input.focus();
        return;
      }
      actualQuantities[input.dataset.lineId] = qty;
    }
    modal.dataset.cfSubmitting = 'yes';

    try {
      const res = await fetch(`/api/pick-orders/${pickOrderNo}/verify-cf`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ verified_by: operator, actual_quantities: actualQuantities })
      });
      if (res.ok) {
        alert(`Đã xác nhận CF thành công cho đơn pick ${pickOrderNo}! Đơn đã chuyển sang trạng thái 'Đã CF' và sẵn sàng tạo BBBG.`);
        closeCfModal();
        loadPickOrders();
      } else {
        const err = await res.json();
        alert(`Lỗi xác nhận CF: ${err.detail || 'Không thể xác nhận'}`);
      }
    } catch (err) {
      alert(`Lỗi kết nối máy chủ: ${err}`);
    } finally {
      modal.dataset.cfSubmitting = '';
    }
  };

  // BBBG Modal
  window.openBbbgModal = function(pickOrderNo) {
    const modal = document.getElementById('bbbg-modal');
    document.getElementById('bbbg-pick-no').value = pickOrderNo;
    modal.style.display = 'flex';
  };

  window.closeBbbgModal = function() {
    document.getElementById('bbbg-modal').style.display = 'none';
  };

  window.submitCreateBbbg = async function() {
    const pickNo = document.getElementById('bbbg-pick-no').value;
    const sender = document.getElementById('bbbg-sender').value;
    const receiver = document.getElementById('bbbg-receiver').value;
    const bin = document.getElementById('bbbg-bin').value;

    try {
      const res = await fetch('/api/bbbg/create', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pick_order_code: pickNo,
          sender_name: sender,
          receiver_name: receiver,
          storage_bin: bin
        })
      });
      if (res.ok) {
        const data = await res.json();
        alert(`Đã tạo Biên bản bàn giao: ${data.bbbg_no}\nModel: ${data.model}\n${data.gos_status === 'ATTACHED' ? 'Đã xác nhận đính kèm SAP.' : 'Chưa xác nhận đính kèm SAP; kiểm tra file và trạng thái trước khi bàn giao.'}`);
        let exportPanel = document.getElementById('bbbg-export-result');
        if (!exportPanel) {
          exportPanel = document.createElement('div');
          exportPanel.id = 'bbbg-export-result';
          document.querySelector('.pick-orders-section')?.appendChild(exportPanel);
        }
        exportPanel.textContent = `${data.is_draft ? 'BBBG bản nháp — cần đối chiếu trước khi dùng: ' : 'BBBG: '}${data.bbbg_no || ''} `;
        if (data.download_url && String(data.download_url).startsWith('/api/')) {
          const link = document.createElement('a');
          link.href = data.download_url;
          link.textContent = 'Tải file BBBG';
          link.setAttribute('download', '');
          exportPanel.appendChild(link);
        }
        closeBbbgModal();
        loadPickOrders();
      } else {
        const err = await res.json();
        alert(`Lỗi tạo BBBG: ${err.detail || 'Không thể tạo BBBG'}`);
      }
    } catch (err) {
      alert(`Lỗi kết nối: ${err}`);
    }
  };

  // Stock Modal
  window.openStockModal = async function() {
    const modal = document.getElementById('stock-modal');
    const container = document.getElementById('stock-modal-list');
    modal.style.display = 'flex';
    container.innerHTML = 'Đang tải tồn kho...';

    try {
      const zone = window.currentBranch === 'EDC_4007' ? 'plant=4007&sloc=4004' : 'plant=3001&sloc=1003';
      const res = await fetch('/api/stock/items?' + zone);
      const items = res.ok ? await res.json() : [];
      window.stockCandidates = items;
      if (items.length === 0) {
        container.innerHTML = `<p style="padding:12px;color:#64748b;">Chưa có snapshot tồn kho. Bạn có thể dùng mẫu có sẵn hoặc dán trực tiếp.</p>`;
        return;
      }
      let html = `<table class="pick-table"><thead><tr><th>Chọn</th><th>Plant/SLoc</th><th>Mã hàng</th><th>Batch</th><th>Tồn khả dụng</th><th>Cảnh báo</th></tr></thead><tbody>`;
      items.forEach((it, idx) => {
        html += `<tr>
          <td><input type="checkbox" class="chk-stock-cand" data-idx="${idx}" checked></td>
          <td>${it.plant} / ${it.sloc}</td>
          <td><strong>${it.material_code}</strong></td>
          <td>${it.batch}</td>
          <td>${it.unrestricted_qty} ${it.unit}</td>
          <td>${it.pull_line_warning || '<span style="color:#059669;">Bình thường</span>'}</td>
        </tr>`;
      });
      html += `</tbody></table>`;
      container.innerHTML = html;
    } catch (err) {
      container.innerHTML = `Lỗi tải tồn: ${err}`;
    }
  };

  window.closeStockModal = function() {
    document.getElementById('stock-modal').style.display = 'none';
  };

  window.importSelectedStock = function() {
    const checks = document.querySelectorAll('.chk-stock-cand:checked');
    checks.forEach(chk => {
      const it = window.stockCandidates[parseInt(chk.dataset.idx, 10)];
      if (it) {
        window.gridData.push({
          id: `stock-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          branch_tab: window.currentBranch,
          material_code: it.material_code,
          model: it.model || '',
          required_qty: it.unrestricted_qty,
          req_qty: it.unrestricted_qty,
          batch: it.batch,
          shop: window.currentBranch.includes('BODY') ? 'BODY' : 'EDC',
          sto_number: '',
          ob_number: '',
          status: 'Chưa kiểm tra',
          selected: true
        });
      }
    });
    closeStockModal();
    renderGrid();
  };

  // Paste Modal
  window.openPasteModal = function() {
    document.getElementById('paste-textarea').value = '';
    document.getElementById('paste-modal').style.display = 'flex';
    document.getElementById('paste-textarea').focus();
  };

  window.closePasteModal = function() {
    document.getElementById('paste-modal').style.display = 'none';
  };

  window.applyManualPaste = function() {
    const text = document.getElementById('paste-textarea').value;
    closePasteModal();
    window.applyPastedText(text);
  };

  window.switchView = function(viewName) {
    const targetName = viewName || 'grid';
    document.querySelectorAll('.view-section').forEach(el => {
      el.style.display = 'none';
    });
    const targetEl = document.getElementById(`view-${targetName}`);
    if (targetEl) {
      targetEl.style.display = 'block';
    } else {
      const gridEl = document.getElementById('view-grid');
      if (gridEl) gridEl.style.display = 'block';
    }

    document.querySelectorAll('.nav-item').forEach(item => {
      if (item.dataset.view === targetName) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    if (targetName === 'dashboard') {
      window.loadDashboardView();
    } else if (targetName === 'stock') {
      window.loadStockView();
    } else if (targetName === 'grid') {
      window.loadGridData(window.currentBranch);
      window.loadPickOrders();
    } else if (targetName === 'waiting_cf') {
      window.loadWaitingCfView();
    } else if (targetName === 'bbbg') {
      window.loadBbbgView();
    } else if (targetName === 'history') {
      window.loadHistoryView();
    } else if (targetName === 'config') {
      window.loadConfigView();
    }
  };

  // --- Navigation View: Dashboard ---
  window.loadDashboardView = async function() {
    const container = document.getElementById('dashboard-summary-cards');
    if (!container) return;
    container.innerHTML = 'Đang tải dữ liệu tổng quan...';
    try {
      const [stockRes, obRes, pickRes, bbbgRes] = await Promise.all([
        fetch('/api/stock/summary').then(r => r.ok ? r.json() : []),
        fetch(`/api/ob/items?branch_tab=${window.currentBranch}`).then(r => r.ok ? r.json() : []),
        fetch('/api/pick-orders').then(r => r.ok ? r.json() : []),
        fetch('/api/bbbg/records').then(r => r.ok ? r.json() : [])
      ]);
      const totalStockQty = (stockRes || []).reduce((acc, it) => acc + (parseFloat(it.total_qty) || 0), 0);
      const waitingPickCount = (pickRes || []).filter(p => p.status === 'Chờ người dùng CF' || p.status === 'PARTIAL_CF' || p.status === 'WAITING_USER_CF').length;
      const confirmedCount = (pickRes || []).filter(p => p.status === 'Đã CF' || p.status === 'USER_CONFIRMED').length;
      const bbbgCount = (bbbgRes || []).length;

      container.innerHTML = `
        <div style="background:#f1f5f9;padding:16px;border-radius:8px;border:1px solid #cbd5e1;">
          <div style="font-size:12px;color:#64748b;font-weight:600;">TỔNG TỒN KHẢ DỤNG</div>
          <div style="font-size:24px;font-weight:700;color:#0f172a;margin-top:4px;">${totalStockQty.toLocaleString()} <span style="font-size:14px;font-weight:400;">PCS</span></div>
          <div style="font-size:12px;color:#475569;margin-top:4px;">${stockRes.length} vùng kho (Plant 3001 & 4007)</div>
        </div>
        <div style="background:#f1f5f9;padding:16px;border-radius:8px;border:1px solid #cbd5e1;">
          <div style="font-size:12px;color:#64748b;font-weight:600;">DÒNG OB TRÊN LƯỚI</div>
          <div style="font-size:24px;font-weight:700;color:#2563eb;margin-top:4px;">${(obRes || []).length} <span style="font-size:14px;font-weight:400;">dòng</span></div>
          <div style="font-size:12px;color:#475569;margin-top:4px;">Nhánh đang chọn: ${window.currentBranch}</div>
        </div>
        <div style="background:#fef3c7;padding:16px;border-radius:8px;border:1px solid #fde68a;">
          <div style="font-size:12px;color:#92400e;font-weight:600;">ĐƠN CHỜ / DỞ DANG CF</div>
          <div style="font-size:24px;font-weight:700;color:#b45309;margin-top:4px;">${waitingPickCount} <span style="font-size:14px;font-weight:400;">đơn</span></div>
          <div style="font-size:12px;color:#92400e;margin-top:4px;">Cần kiểm đếm thực tế trước khi BBBG</div>
        </div>
        <div style="background:#dcfce7;padding:16px;border-radius:8px;border:1px solid #86efac;">
          <div style="font-size:12px;color:#166534;font-weight:600;">BBBG ĐÃ LẬP (V2.2.14)</div>
          <div style="font-size:24px;font-weight:700;color:#15803d;margin-top:4px;">${bbbgCount} <span style="font-size:14px;font-weight:400;">phiếu</span></div>
          <div style="font-size:12px;color:#166534;margin-top:4px;">${confirmedCount} đơn pick đã CF đủ sẵn sàng</div>
        </div>
      `;
    } catch (err) {
      container.innerHTML = `Lỗi tải dữ liệu tổng quan: ${err}`;
    }
  };

  // --- Navigation View: Stock ---
  window.loadStockView = async function() {
    await checkStockSapHealth();
    await loadVisualStockMatrix();
    await loadTwoStockTables();
    await loadLowStockSuggestions();
  };

  async function checkStockSapHealth() {
    const badge = document.getElementById('stock-sap-status-badge');
    if (!badge) return;
    try {
      const res = await fetch('/api/health');
      if (res.ok) {
        const data = await res.json();
        if (data.sap_connection === 'CONNECTED') {
          badge.style.background = '#ecfdf5';
          badge.style.color = '#047857';
          badge.style.borderColor = '#10b981';
          badge.textContent = '● SAP: Đã kết nối GUI';
        } else {
          badge.style.background = '#fef2f2';
          badge.style.color = '#b91c1c';
          badge.style.borderColor = '#fca5a5';
          badge.textContent = '● SAP: Chưa kết nối';
        }
      }
    } catch (_) {
      badge.textContent = '● SAP: Chưa kiểm tra được';
    }
  }

  window.downloadSapMb52Direct = async function() {
    const fb = document.getElementById('sap-download-feedback');
    const btn = document.getElementById('btn-sap-download-mb52');
    if (btn) btn.disabled = true;
    if (fb) {
      fb.style.display = 'block';
      fb.style.background = '#eff6ff';
      fb.style.color = '#1e40af';
      fb.style.border = '1px solid #bfdbfe';
      fb.textContent = 'Đang gửi lệnh tải tồn trực tiếp từ SAP MB52...';
    }
    try {
      const res = await fetch('/api/stock/sap/download-mb52', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ layout: '/KHOMA' })
      });
      const data = await res.json();
      if (!res.ok || data.status === 'SAP_NOT_CONNECTED' || data.status === 'ERROR') {
        if (fb) {
          fb.style.background = '#fef2f2';
          fb.style.color = '#b91c1c';
          fb.style.border = '1px solid #fca5a5';
          fb.textContent = `⚠️ ${data.message || data.detail || 'Không thể tải tồn từ SAP'}`;
        }
      } else {
        if (fb) {
          fb.style.background = '#f0fdf4';
          fb.style.color = '#166534';
          fb.style.border = '1px solid #86efac';
          fb.textContent = `✓ ${data.message || 'Tải tồn MB52 thành công!'}`;
        }
        await loadStockView();
      }
    } catch (err) {
      if (fb) {
        fb.style.background = '#fef2f2';
        fb.style.color = '#b91c1c';
        fb.style.border = '1px solid #fca5a5';
        fb.textContent = `Lỗi tải tồn SAP: ${err.message}`;
      }
    } finally {
      if (btn) btn.disabled = false;
    }
  };

  window.uploadMb52File = async function(event) {
    const file = event.target.files[0];
    if (!file) return;
    const fb = document.getElementById('sap-download-feedback');
    if (fb) {
      fb.style.display = 'block';
      fb.style.background = '#eff6ff';
      fb.style.color = '#1e40af';
      fb.style.border = '1px solid #bfdbfe';
      fb.textContent = `Đang nạp file MB52 dự phòng: ${file.name}...`;
    }
    try {
      const res = await fetch(`/api/stock/import-file?filename=${encodeURIComponent(file.name)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: file
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Lỗi nạp file MB52');
      if (fb) {
        fb.style.background = '#f0fdf4';
        fb.style.color = '#166534';
        fb.style.border = '1px solid #86efac';
        fb.textContent = `✓ Nạp file MB52 thành công: ${data.total_rows} dòng, Tổng SL: ${(data.total_qty || 0).toLocaleString()} PCS. Snapshot mới đã thay thế tồn scope hiện hành.`;
      }
      await loadStockView();
    } catch (err) {
      if (fb) {
        fb.style.background = '#fef2f2';
        fb.style.color = '#b91c1c';
        fb.style.border = '1px solid #fca5a5';
        fb.textContent = `Lỗi: ${err.message}`;
      }
    } finally {
      event.target.value = '';
    }
  };

  window.loadVisualStockMatrix = async function() {
    const thead = document.getElementById('visual-matrix-thead');
    const tbody = document.getElementById('visual-matrix-tbody');
    const tfoot = document.getElementById('visual-matrix-tfoot');
    if (!tbody) return;

    try {
      const res = await fetch('/api/stock/visual-matrix');
      if (!res.ok) throw new Error('Không tải được ma trận tồn');
      const data = await res.json();

      const titleEl = document.getElementById('matrix-title-display');
      const timeEl = document.getElementById('matrix-time-display');
      const dateEl = document.getElementById('matrix-date-display');
      if (titleEl) titleEl.textContent = data.title || '(LINE + KHO) BODY';
      if (timeEl) timeEl.textContent = data.updated_time || '--';
      if (dateEl) dateEl.textContent = data.updated_date || '--';

      const cols = data.columns || [];

      // Render Header
      let headHtml = `<tr>
        <th style="width:40px;background:#f8fafc;border:1px solid #cbd5e1;text-align:center;">Stt</th>
        <th style="width:90px;background:#f8fafc;border:1px solid #cbd5e1;text-align:center;">Xe</th>`;
      cols.forEach(c => {
        const color = c.is_alert ? '#b91c1c' : '#0f172a';
        headHtml += `<th style="background:#f8fafc;border:1px solid #cbd5e1;text-align:center;color:${color};font-weight:700;">${c.label}</th>`;
      });
      headHtml += `<th style="width:90px;background:#f8fafc;border:1px solid #cbd5e1;text-align:center;font-weight:700;">Tổng SL mã</th></tr>`;
      thead.innerHTML = headHtml;

      // Render Body
      let bodyHtml = '';
      (data.rows || []).forEach(r => {
        bodyHtml += `<tr>
          <td style="border:1px solid #cbd5e1;text-align:center;font-weight:600;">${r.stt}</td>
          <td style="border:1px solid #cbd5e1;text-align:center;font-weight:700;">${r.model}</td>`;
        cols.forEach(c => {
          const count = r.counts[c.key] || 0;
          const color = c.is_alert ? '#dc2626' : '#0f172a';
          if (count > 0) {
            bodyHtml += `<td style="border:1px solid #cbd5e1;text-align:center;font-weight:600;">
              <a href="#" onclick="openMatrixDrilldown('${r.model}', '${c.key}'); return false;" style="color:${color};text-decoration:underline;">${count}</a>
            </td>`;
          } else {
            bodyHtml += `<td style="border:1px solid #cbd5e1;text-align:center;color:#94a3b8;"></td>`;
          }
        });
        bodyHtml += `<td style="border:1px solid #cbd5e1;text-align:center;font-weight:700;background:#f8fafc;">
          <a href="#" onclick="openMatrixDrilldown('${r.model}', ''); return false;" style="color:#0f172a;text-decoration:underline;">${r.total_materials}</a>
        </td></tr>`;
      });
      tbody.innerHTML = bodyHtml;

      // Render Footer
      const totals = data.totals || {};
      const counts = totals.counts || {};
      const pcts = totals.percentages || {};

      let footHtml = `<tr style="background:#f1f5f9;font-weight:700;">
        <td colspan="2" style="border:1px solid #cbd5e1;text-align:center;">Tổng SL mã</td>`;
      cols.forEach(c => {
        const cVal = counts[c.key] || 0;
        const color = c.is_alert ? '#b91c1c' : '#0f172a';
        footHtml += `<td style="border:1px solid #cbd5e1;text-align:center;color:${color};">${cVal > 0 ? cVal : ''}</td>`;
      });
      footHtml += `<td style="border:1px solid #cbd5e1;text-align:center;">${totals.grand_total || 0}</td></tr>`;

      footHtml += `<tr style="background:#f8fafc;font-weight:700;">
        <td colspan="2" style="border:1px solid #cbd5e1;text-align:center;">Tỷ lệ</td>`;
      cols.forEach(c => {
        const pct = pcts[c.key] || '0%';
        const color = c.is_alert ? '#b91c1c' : '#0f172a';
        footHtml += `<td style="border:1px solid #cbd5e1;text-align:center;color:${color};">${pct}</td>`;
      });
      footHtml += `<td style="border:1px solid #cbd5e1;text-align:center;">100%</td></tr>`;
      tfoot.innerHTML = footHtml;
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="12" style="text-align:center;color:#ef4444;padding:16px;">Lỗi tải ma trận tồn: ${err.message}</td></tr>`;
    }
  };

  window.openMatrixDrilldown = async function(model, bucket) {
    const modal = document.getElementById('matrix-drilldown-modal');
    const title = document.getElementById('modal-matrix-filter-title');
    const tbody = document.getElementById('matrix-drilldown-body');
    if (!modal || !tbody) return;

    modal.style.display = 'flex';
    title.textContent = `Model: ${model || 'Tất cả'} • Dải tồn: ${bucket || 'Tất cả'}`;
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:16px;">Đang tải danh sách mã...</td></tr>';

    try {
      const params = new URLSearchParams();
      if (model) params.append('model', model);
      if (bucket) params.append('bucket', bucket);
      const res = await fetch(`/api/stock/visual-matrix/materials?${params.toString()}`);
      if (!res.ok) throw new Error('Không tải được danh sách mã');
      const items = await res.json();

      if (items.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:16px;color:#94a3b8;">Không có mã hàng nào trong ô này</td></tr>';
        return;
      }
      tbody.innerHTML = items.map((it, idx) => `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td><strong>${it.material_code}</strong></td>
          <td>${it.model || '--'}</td>
          <td style="text-align:center;">${it.us_rate || 1}</td>
          <td style="text-align:right;">${(it.line_3003 || 0).toLocaleString()}</td>
          <td style="text-align:right;">${(it.kho_1003 || 0).toLocaleString()}</td>
          <td style="text-align:right;font-weight:600;">${(it.tong_body || 0).toLocaleString()}</td>
          <td style="text-align:center;font-weight:700;color:${it.doh <= 200 ? '#dc2626' : '#059669'};">${it.doh}</td>
          <td style="text-align:center;"><span class="badge" style="background:#f1f5f9;color:#334155;font-size:11px;">${it.bucket}</span></td>
        </tr>
      `).join('');
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#ef4444;padding:16px;">Lỗi: ${err.message}</td></tr>`;
    }
  };

  window.closeMatrixDrilldownModal = function() {
    const modal = document.getElementById('matrix-drilldown-modal');
    if (modal) modal.style.display = 'none';
  };

  window.twoStockTablesData = { priority_table: [], total_stock_table: [] };

  window.loadTwoStockTables = async function() {
    const plant = document.getElementById('stock-filter-plant-sub')?.value || '3001';
    try {
      const res = await fetch(`/api/stock/two-tables?plant=${plant}`);
      if (!res.ok) throw new Error('Không tải được bảng tồn');
      window.twoStockTablesData = await res.json();
      renderTwoStockTables();
    } catch (err) {
      console.error('Error loading two stock tables:', err);
    }
  };

  function renderTwoStockTables() {
    const q = (document.getElementById('stock-filter-material-sub')?.value || '').trim().toLowerCase();
    const pBody = document.getElementById('stock-priority-body');
    const tBody = document.getElementById('stock-total-body');

    // Priority Table
    if (pBody) {
      const pItems = (window.twoStockTablesData.priority_table || []).filter(it => !q || it.material_code.toLowerCase().includes(q));
      if (pItems.length === 0) {
        pBody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#94a3b8;padding:16px;">Không có mã ưu tiên xuất</td></tr>';
      } else {
        pBody.innerHTML = pItems.map((it, idx) => `
          <tr>
            <td style="text-align:center;">${idx + 1}</td>
            <td><strong>${it.material_code}</strong></td>
            <td>${it.model || '--'}</td>
            <td style="text-align:center;">${it.us_rate || 1}</td>
            <td style="text-align:right;">${(it.line_3003 || 0).toLocaleString()}</td>
            <td style="text-align:right;">${(it.kho_1003 || 0).toLocaleString()}</td>
            <td style="text-align:right;font-weight:600;">${(it.tong_body || 0).toLocaleString()}</td>
            <td style="text-align:center;font-weight:700;color:${it.doh <= 200 ? '#dc2626' : '#059669'};">${it.doh}</td>
            <td style="color:#b91c1c;font-size:12px;font-weight:600;">🔥 ${it.reason || 'Ưu tiên xuất'}</td>
          </tr>
        `).join('');
      }
    }

    // Total Stock Table
    if (tBody) {
      const tItems = (window.twoStockTablesData.total_stock_table || []).filter(it => !q || it.material_code.toLowerCase().includes(q));
      if (tItems.length === 0) {
        tBody.innerHTML = '<tr><td colspan="11" style="text-align:center;color:#94a3b8;padding:16px;">Không có dữ liệu tồn chi tiết</td></tr>';
      } else {
        tBody.innerHTML = tItems.map((it, idx) => `
          <tr>
            <td style="text-align:center;">${idx + 1}</td>
            <td><strong>${it.material_code}</strong></td>
            <td>${it.model || '--'}</td>
            <td style="text-align:center;">${it.us_rate || 1}</td>
            <td style="text-align:right;">${(it.line_3003 || 0).toLocaleString()}</td>
            <td style="text-align:right;">${(it.kho_1003 || 0).toLocaleString()}</td>
            <td style="text-align:right;">${(it.line_3000 || 0).toLocaleString()}</td>
            <td style="text-align:right;">${(it.kho_3093 || 0).toLocaleString()}</td>
            <td style="text-align:right;">${(it.in_transit_3001 || 0).toLocaleString()}</td>
            <td style="text-align:right;font-weight:600;">${(it.tong_body || 0).toLocaleString()}</td>
            <td style="text-align:center;font-weight:700;">${it.doh}</td>
          </tr>
        `).join('');
      }
    }
  }

  window.switchStockSubTable = function(tab) {
    const cPri = document.getElementById('container-priority-table');
    const cTot = document.getElementById('container-total-table');
    const bPri = document.getElementById('btn-tab-priority-table');
    const bTot = document.getElementById('btn-tab-total-table');

    if (tab === 'priority') {
      if (cPri) cPri.style.display = 'block';
      if (cTot) cTot.style.display = 'none';
      if (bPri) { bPri.className = 'btn btn-primary'; }
      if (bTot) { bTot.className = 'btn btn-outline'; }
    } else {
      if (cPri) cPri.style.display = 'none';
      if (cTot) cTot.style.display = 'block';
      if (bPri) { bPri.className = 'btn btn-outline'; }
      if (bTot) { bTot.className = 'btn btn-primary'; }
    }
  };

  window.filterStockSubTables = function() {
    renderTwoStockTables();
  };

  window.lowStockSuggestionsData = [];

  window.loadLowStockSuggestions = async function() {
    const tbody = document.getElementById('suggestions-body');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:16px;">Đang tính toán lượng bổ sung từ DOH...</td></tr>';

    try {
      const res = await fetch('/api/stock/low-stock/suggestions?plant=3001&shop=BODY');
      if (!res.ok) throw new Error('Không tải được gợi ý tồn thấp');
      const items = await res.json();
      window.lowStockSuggestionsData = items;

      if (items.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#059669;padding:16px;">✓ Tồn kho đạt định mức an toàn; không có mã nào dưới ngưỡng DOH</td></tr>';
        return;
      }

      tbody.innerHTML = items.map((it, idx) => {
        const batchOpts = (it.batches || []).map(b => `<span style="display:inline-block;margin:2px 4px;padding:2px 6px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:4px;font-size:11px;">Lô ${b.batch} (Khả dụng: ${b.available_qty})</span>`).join('');
        return `
          <tr data-sugg-idx="${idx}">
            <td style="text-align:center;"><input type="checkbox" class="chk-sugg-row" data-idx="${idx}" checked></td>
            <td><strong>${it.material_code}</strong></td>
            <td>${it.model || '--'}</td>
            <td style="text-align:center;">${it.us_rate || 1}</td>
            <td style="text-align:right;">${it.tong_body}</td>
            <td style="text-align:center;font-weight:700;color:#dc2626;">${it.current_doh}</td>
            <td style="text-align:center;">${it.threshold_doh}</td>
            <td>
              <input type="number" step="any" min="1" id="sugg-qty-${idx}" value="${it.supplement_needed}" style="width:90px;padding:4px 6px;border:1px solid #cbd5e1;border-radius:4px;font-weight:600;" />
            </td>
            <td style="font-size:12px;">
              ${batchOpts || '<em style="color:#94a3b8;">Không có lô khả dụng</em>'}
            </td>
          </tr>
        `;
      }).join('');
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#ef4444;padding:16px;">Lỗi: ${err.message}</td></tr>`;
    }
  };

  window.toggleAllSuggestions = function(checked) {
    document.querySelectorAll('.chk-sugg-row').forEach(cb => { cb.checked = checked; });
  };

  window.transferSelectedSuggestionsToOB = async function() {
    const checkedBoxes = document.querySelectorAll('.chk-sugg-row:checked');
    if (checkedBoxes.length === 0) {
      alert('Vui lòng chọn ít nhất 1 mã gợi ý tồn thấp để chuyển sang Bảng điền OB!');
      return;
    }

    const itemsToTransfer = [];
    checkedBoxes.forEach(cb => {
      const idx = parseInt(cb.dataset.idx, 10);
      const sugg = window.lowStockSuggestionsData[idx];
      if (sugg) {
        const qtyInput = document.getElementById(`sugg-qty-${idx}`);
        const userQty = qtyInput ? parseFloat(qtyInput.value) : sugg.supplement_needed;
        const firstBatch = (sugg.batches && sugg.batches.length > 0) ? sugg.batches[0].batch : '';
        itemsToTransfer.push({
          material_code: sugg.material_code,
          model: sugg.model,
          required_qty: userQty,
          qty: userQty,
          batch: firstBatch,
          shop: sugg.shop || 'BODY'
        });
      }
    });

    try {
      const res = await fetch('/api/stock/low-stock/transfer-to-ob', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          branch_tab: window.currentBranch,
          items: itemsToTransfer
        })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Lỗi chuyển gợi ý sang OB');
      alert(`Đã chuyển thành công ${data.transferred_count} mã gợi ý tồn thấp sang Bảng điền OB!`);
      switchView('grid');
    } catch (err) {
      alert(`Lỗi: ${err.message}`);
    }
  };
  window.importSelectedStockFromView = function() {
    const checkboxes = document.querySelectorAll('.chk-stock-view-item:checked');
    if (checkboxes.length === 0) {
      alert('Vui lòng chọn ít nhất 1 mã tồn kho để nạp vào lưới!');
      return;
    }
    checkboxes.forEach(chk => {
      const idx = parseInt(chk.dataset.idx, 10);
      const it = window.stockViewItems[idx];
      if (it) {
        window.gridData.push({
          id: `stock-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          branch_tab: window.currentBranch,
          material_code: it.material_code,
          model: it.model || '',
          required_qty: it.unrestricted_qty,
          req_qty: it.unrestricted_qty,
          batch: it.batch,
          shop: window.currentBranch.includes('BODY') ? 'BODY' : 'EDC',
          sto_number: '',
          ob_number: '',
          status: 'Chưa kiểm tra',
          selected: true
        });
      }
    });
    alert(`Đã nạp ${checkboxes.length} mã vật tư từ tồn kho vào lưới OB!`);
    switchView('grid');
  };

  // --- Navigation View: Waiting CF ---
  window.loadWaitingCfView = async function() {
    const tbody = document.getElementById('waiting-cf-body');
    const histBody = document.getElementById('waiting-cf-history-body');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:16px;">Đang tải danh sách chờ CF...</td></tr>';

    const ob = (document.getElementById('cf-filter-ob')?.value || '').trim();
    const pick = (document.getElementById('cf-filter-pick')?.value || '').trim();
    const mat = (document.getElementById('cf-filter-material')?.value || '').trim();
    const batch = (document.getElementById('cf-filter-batch')?.value || '').trim();
    const status = document.getElementById('cf-filter-status')?.value || '';

    try {
      const params = new URLSearchParams();
      if (ob) params.append('ob_number', ob);
      if (pick) params.append('pick_order_code', pick);
      if (mat) params.append('material_code', mat);
      if (batch) params.append('batch', batch);
      if (status) params.append('status', status);

      const [lines, history] = await Promise.all([
        fetch(`/api/waiting-cf?${params.toString()}`).then(r => r.ok ? r.json() : []),
        fetch('/api/waiting-cf/history?limit=50').then(r => r.ok ? r.json() : [])
      ]);

      window.waitingCfLines = lines;
      const badgeEl = document.getElementById('badge-cf-count');
      if (badgeEl) badgeEl.textContent = lines.length;

      if (lines.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" style="text-align:center;color:#059669;padding:16px;">✓ Không có dòng nào đang chờ CF (hoặc tất cả các đơn đã hoàn tất CF 100% và rời danh sách)</td></tr>';
      } else {
        tbody.innerHTML = lines.map((l, idx) => {
          const isPartial = l.cf_status === 'PARTIAL_CF' || (l.cf_actual_qty > 0 && l.cf_actual_qty < l.req_qty);
          const statusBadge = isPartial
            ? `<span class="chip chip-waiting-pick" style="background:#fef3c7;color:#92400e;border:1px solid #fde68a;">CF một phần</span>`
            : `<span class="chip chip-waiting-pick">Chờ CF</span>`;
          return `
            <tr data-line-id="${l.id}">
              <td><strong>${l.pick_order_code}</strong></td>
              <td>${l.ob_number || '—'}</td>
              <td><strong>${l.material_code}</strong></td>
              <td>${l.batch || '—'}</td>
              <td>${l.req_qty}</td>
              <td>
                <input type="number" step="any" min="0" max="${l.req_qty}" class="search-input cf-input-actual" id="cf-act-${l.id}" value="${l.cf_actual_qty > 0 ? l.cf_actual_qty : ''}" placeholder="${l.req_qty}" style="width:90px;" oninput="onActualQtyChange(${l.id}, ${l.req_qty})">
              </td>
              <td><span id="cf-miss-${l.id}" style="color:${l.missing_qty > 0 ? '#b45309' : '#059669'};font-weight:600;">${l.missing_qty}</span></td>
              <td>
                <input type="text" class="search-input cf-input-reason" id="cf-reason-${l.id}" value="${l.shortage_reason || ''}" placeholder="Lý do nếu thiếu..." style="width:140px;">
              </td>
              <td>${statusBadge}</td>
              <td style="font-size:12px;color:#475569;">${l.cf_source_label || 'Thực tế'}</td>
              <td style="font-size:12px;color:#64748b;">${l.verified_by ? `${l.verified_by}<br><small>${l.verified_at ? l.verified_at.slice(0, 16) : ''}</small>` : 'Chưa kiểm'}</td>
              <td>
                <button class="btn btn-outline" style="padding:4px 8px;font-size:12px;" onclick="saveWaitingCfLine(${l.id})">Lưu dòng</button>
              </td>
            </tr>
          `;
        }).join('');
      }

      if (histBody) {
        if (history.length === 0) {
          histBody.innerHTML = '<tr><td colspan="11" style="text-align:center;color:#94a3b8;padding:16px;">Chưa có lịch sử kiểm tra CF</td></tr>';
        } else {
          histBody.innerHTML = history.map(h => `
            <tr>
              <td>${h.verified_at || h.created_at || ''}</td>
              <td><strong>${h.verified_by}</strong></td>
              <td>${h.ob_number || '—'}</td>
              <td>${h.material_code || '—'}</td>
              <td>${h.batch || '—'}</td>
              <td>${h.req_qty}</td>
              <td>${h.actual_qty}</td>
              <td>${h.missing_qty > 0 ? `<span style="color:#d97706;font-weight:600;">${h.missing_qty}</span>` : '0'}</td>
              <td>${h.reason || '—'}</td>
              <td>${renderStatusChip(h.status)}</td>
              <td style="font-size:12px;color:#475569;">${h.evidence || h.cf_source || ''}</td>
            </tr>
          `).join('');
        }
      }
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="12" style="text-align:center;color:#ef4444;padding:16px;">Lỗi tải dữ liệu Chờ CF: ${err}</td></tr>`;
    }
  };

  window.onActualQtyChange = function(lineId, reqQty) {
    const actInput = document.getElementById(`cf-act-${lineId}`);
    const missSpan = document.getElementById(`cf-miss-${lineId}`);
    if (!actInput || !missSpan) return;
    const val = parseFloat(actInput.value) || 0;
    const missing = Math.max(0, reqQty - val);
    missSpan.textContent = missing;
    missSpan.style.color = missing > 0 ? '#b45309' : '#059669';
  };

  window.saveWaitingCfLine = async function(lineId) {
    const operatorInput = document.getElementById('cf-view-operator');
    const evidenceInput = document.getElementById('cf-view-evidence');
    const actInput = document.getElementById(`cf-act-${lineId}`);
    const reasonInput = document.getElementById(`cf-reason-${lineId}`);

    const operator = (operatorInput?.value || '').trim();
    if (!operator) {
      alert('Vui lòng nhập tên người kiểm tra thực tế ở ô "Người kiểm thực tế"!');
      if (operatorInput) operatorInput.focus();
      return;
    }
    const actVal = actInput?.value !== '' ? parseFloat(actInput.value) : null;
    if (actVal === null || isNaN(actVal) || actVal < 0) {
      alert('Vui lòng nhập số lượng thực tế hợp lệ (>= 0)!');
      if (actInput) actInput.focus();
      return;
    }
    const reason = (reasonInput?.value || '').trim();
    const evidence = (evidenceInput?.value || '').trim();

    try {
      const res = await fetch('/api/waiting-cf/save-line', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          line_id: lineId,
          actual_qty: actVal,
          verified_by: operator,
          reason: reason,
          evidence: evidence
        })
      });
      const data = await res.json();
      if (res.ok) {
        if (data.bbbg_unlocked) {
          alert(`Đã hoàn tất CF 100% cho đơn pick ${data.pick_order_code}! Đơn đã đủ điều kiện tạo BBBG và rời danh sách Chờ CF.`);
        } else {
          alert(`Đã lưu tiến độ dòng ${lineId} (Trạng thái: ${data.cf_status}, Còn thiếu: ${data.missing_qty}). Đơn pick giữ trạng thái PARTIAL_CF (BBBG vẫn khóa).`);
        }
        loadWaitingCfView();
        loadPickOrders();
      } else {
        alert(`Lỗi lưu CF: ${data.detail || 'Không thể lưu'}`);
      }
    } catch (err) {
      alert(`Lỗi kết nối: ${err}`);
    }
  };

  window.saveAllWaitingCfLines = async function() {
    const operator = (document.getElementById('cf-view-operator')?.value || '').trim();
    if (!operator) {
      alert('Vui lòng nhập tên người kiểm tra thực tế!');
      document.getElementById('cf-view-operator')?.focus();
      return;
    }
    const evidence = (document.getElementById('cf-view-evidence')?.value || '').trim();
    const linesToSave = [];
    (window.waitingCfLines || []).forEach(l => {
      const actInput = document.getElementById(`cf-act-${l.id}`);
      const reasonInput = document.getElementById(`cf-reason-${l.id}`);
      if (actInput && actInput.value !== '') {
        const actVal = parseFloat(actInput.value);
        if (!isNaN(actVal) && actVal >= 0) {
          linesToSave.push({
            line_id: l.id,
            actual_qty: actVal,
            reason: (reasonInput?.value || '').trim(),
            evidence: evidence
          });
        }
      }
    });
    if (linesToSave.length === 0) {
      alert('Chưa có dòng nào được nhập SL thực tế để lưu!');
      return;
    }
    try {
      const res = await fetch('/api/waiting-cf/bulk-save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ verified_by: operator, lines: linesToSave, notes: evidence })
      });
      const data = await res.json();
      if (res.ok) {
        alert(`Đã lưu thành công ${data.saved_count || 0} dòng kiểm tra CF!`);
        loadWaitingCfView();
        loadPickOrders();
      } else {
        alert(`Lỗi lưu tiến độ: ${data.detail || 'Không thể lưu'}`);
      }
    } catch (err) {
      alert(`Lỗi kết nối: ${err}`);
    }
  };

  // --- Navigation View: BBBG ---
  window.loadBbbgView = async function() {
    const tbody = document.getElementById('bbbg-records-body');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:16px;">Đang tải danh sách BBBG...</td></tr>';
    try {
      const records = await fetch('/api/bbbg/records').then(r => r.ok ? r.json() : []);
      if (records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:#94a3b8;padding:16px;">Chưa có bản ghi BBBG nào được tạo</td></tr>';
        return;
      }
      tbody.innerHTML = records.map(b => `
        <tr>
          <td><strong>${b.bbbg_no || b.slip_no || b.bbbg_uuid}</strong></td>
          <td>${b.pick_order_no || b.pick_order_id || '—'}</td>
          <td>${b.shop || ''} · ${b.model || ''}</td>
          <td>${(b.total_qty || 0).toLocaleString()}</td>
          <td>${b.sender_name || '—'}</td>
          <td>${b.receiver_name || '—'}</td>
          <td>${b.storage_bin || '—'}</td>
          <td><span class="chip chip-ready">${b.status || 'CREATED'}</span></td>
          <td style="font-size:12px;color:#d97706;">
            SAP live write: Chưa hỗ trợ<br>
            GOS/Email/Print/UNC: Chưa hỗ trợ
          </td>
          <td>
            <a href="/api/bbbg/${b.id}/download" download class="btn btn-secondary" style="padding:4px 8px;font-size:12px;text-decoration:none;">📥 Tải Excel</a>
          </td>
        </tr>
      `).join('');
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;color:#ef4444;padding:16px;">Lỗi tải BBBG: ${err}</td></tr>`;
    }
  };

  // --- Navigation View: History ---
  window.loadHistoryView = async function() {
    const tbody = document.getElementById('history-events-body');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:16px;">Đang tải lịch sử sự kiện...</td></tr>';
    try {
      const events = await fetch('/api/jobs/events?limit=100').then(r => r.ok ? r.json() : []);
      window.historyEvents = events;
      renderHistoryTable();
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:#ef4444;padding:16px;">Lỗi tải lịch sử: ${err}</td></tr>`;
    }
  };

  window.renderHistoryTable = function() {
    const tbody = document.getElementById('history-events-body');
    if (!tbody) return;
    const q = (document.getElementById('history-filter-q')?.value || '').trim().toLowerCase();
    const events = (window.historyEvents || []).filter(ev => {
      if (!q) return true;
      const text = [ev.job_id, ev.business_key, ev.job_type, ev.step_name, ev.event_type, ev.target_ref, ev.status, ev.evidence].join(' ').toLowerCase();
      return text.includes(q);
    });
    if (events.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#94a3b8;padding:16px;">Không có sự kiện nào</td></tr>';
      return;
    }
    tbody.innerHTML = events.map(ev => `
      <tr>
        <td>${ev.id}</td>
        <td style="font-size:12px;">${ev.created_at || ''}</td>
        <td><strong>${ev.job_type || ''}</strong></td>
        <td>${ev.step_name || ev.event_type || ''}</td>
        <td>${ev.target_ref || ev.business_key || ''}</td>
        <td><span class="chip ${ev.status === 'SUCCESS' ? 'chip-ready' : 'chip-pending'}">${ev.status || ''}</span></td>
        <td style="font-size:12px;color:#334155;max-width:300px;word-break:break-word;">${ev.evidence || ev.sap_status_msg || '—'}</td>
      </tr>
    `).join('');
  };

  // --- Navigation View: Config ---
  window.loadConfigView = async function() {
    const box = document.getElementById('config-settings-box');
    const mapBody = document.getElementById('config-mappings-body');
    try {
      const [settings, mappings] = await Promise.all([
        fetch('/api/settings').then(r => r.ok ? r.json() : {}),
        fetch('/api/model-mappings').then(r => r.ok ? r.json() : [])
      ]);
      if (box) {
        box.innerHTML = `
          <pre style="background:#f8fafc;padding:12px;border:1px solid #cbd5e1;border-radius:6px;font-family:monospace;font-size:12px;max-height:160px;overflow-y:auto;">${JSON.stringify(settings, null, 2)}</pre>
        `;
      }
      if (mapBody) {
        mapBody.innerHTML = (mappings || []).map(m => `
          <tr>
            <td><strong>${m.material_code}</strong></td>
            <td>${m.model_code}</td>
            <td>${m.shop_code || '—'}</td>
            <td>${m.note || '—'}</td>
          </tr>
        `).join('');
      }
    } catch (err) {
      console.error('Error loading config:', err);
    }
  };

  window.saveModelMappingConfig = async function() {
    const mat = (document.getElementById('cfg-map-mat')?.value || '').trim();
    const model = (document.getElementById('cfg-map-model')?.value || '').trim();
    const shop = (document.getElementById('cfg-map-shop')?.value || '').trim();
    if (!mat || !model) {
      alert('Nhập mã hàng và model code!');
      return;
    }
    try {
      const res = await fetch('/api/model-mappings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ material_code: mat, model_code: model, shop_code: shop, note: 'Nhập từ giao diện Pilot' })
      });
      if (res.ok) {
        alert(`Đã lưu mapping cho mã hàng ${mat}!`);
        document.getElementById('cfg-map-mat').value = '';
        document.getElementById('cfg-map-model').value = '';
        loadConfigView();
      } else {
        alert('Lỗi lưu mapping');
      }
    } catch (err) {
      alert(`Lỗi kết nối: ${err}`);
    }
  };

  function setupEventListeners() {
    window.addEventListener('paste', handlePaste);

    window.addEventListener('keydown', (e) => {
      if (window.isEditing || document.activeElement?.tagName === 'INPUT' || document.activeElement?.tagName === 'TEXTAREA') return;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        selectCell(window.activeRowIndex + 1, window.activeColKey);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        selectCell(window.activeRowIndex - 1, window.activeColKey);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        navigateNextCell(window.activeRowIndex, window.activeColKey, e.shiftKey);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        selectCell(window.activeRowIndex + 1, window.activeColKey);
      }
    });

    const chkAll = document.getElementById('chk-all');
    if (chkAll) {
      chkAll.addEventListener('change', () => {
        window.gridData.forEach(r => r.selected = chkAll.checked);
        renderGrid();
      });
    }

    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        window.currentBranch = btn.dataset.branch;
        if (window.currentBranch === 'EDC_4007') {
          alert('Nhánh 4007 đang được hoãn theo chỉ đạo người dùng; không tự gộp vào BODY.');
        }
        loadGridData(window.currentBranch);
      });
    });

    document.getElementById('input-search')?.addEventListener('input', () => renderGrid());
    document.getElementById('select-model')?.addEventListener('change', () => renderGrid());
    document.getElementById('select-status')?.addEventListener('change', () => renderGrid());

    document.getElementById('btn-get-stock')?.addEventListener('click', () => openStockModal());
    document.getElementById('btn-paste-excel')?.addEventListener('click', () => openPasteModal());
    document.getElementById('btn-add-row')?.addEventListener('click', () => addNewRow());
    document.getElementById('btn-del-row')?.addEventListener('click', () => deleteActiveRow());
    document.getElementById('btn-check')?.addEventListener('click', runCheckAction);
    document.getElementById('btn-fill-ob')?.addEventListener('click', runFillObAction);
    document.getElementById('btn-create-pick')?.addEventListener('click', runCreatePickAction);
  }

  async function runCheckAction() {
    const selected = window.gridData.filter(r => String(r.material_code || '').trim().length > 0);
    if (selected.length === 0) {
      alert('Vui lòng nhập hoặc dán ít nhất 1 mã hàng để kiểm tra!');
      return;
    }
    if (window.obCheckRunning) return;
    window.obCheckRunning = true;
    const branch = window.currentBranch;
    const submitted = window.gridData.filter(row => String(row.material_code || '').trim()).map(row => ({
      ...row,
      item_uid: row.item_uid || String(row.id),
      selected: 1
    }));
    const scopedKey = row => {
      const key = row.item_uid || String(row.id);
      return key.startsWith(`${branch}:`) ? key : `${branch}:${key}`;
    };
    const selectedKeys = new Set(selected.map(scopedKey));
    try {
      const save = await fetch('/api/ob/batch-save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ branch_tab: branch, items: submitted })
      });
      if (!save.ok) throw new Error('Không lưu được các dòng; chưa kiểm tra tồn kho.');
      const fresh = await fetch(`/api/ob/items?branch_tab=${encodeURIComponent(branch)}`);
      if (!fresh.ok) throw new Error('Không tải được ID dòng đã lưu; chưa kiểm tra tồn kho.');
      const rows = await fresh.json();
      const itemIds = rows.filter(row => selectedKeys.has(scopedKey(row)))
        .map(row => Number(row.id)).filter(Number.isInteger);
      if (!itemIds.length || itemIds.length !== selected.length) throw new Error('Không đối chiếu đủ các dòng đã chọn; hãy kiểm tra lại dữ liệu.');
      const body = branch.includes('BODY');
      const check = await fetch('/api/ob/check', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ branch_tab: branch, item_ids: itemIds,
          source_plant: body ? '3001' : '4007', source_sloc: body ? '1003' : '4004' })
      });
      const result = await check.json();
      if (!check.ok) throw new Error(result.detail || 'Không kiểm tra được tồn kho.');
      if (window.currentBranch === branch) {
        await loadGridData(branch);
        window.gridData.forEach(row => { row.selected = itemIds.includes(Number(row.id)); });
        renderGrid();
      }
      alert(`Đã kiểm tra ${result.checked_count || 0} dòng với dữ liệu tồn kho. Xem trạng thái từng dòng trước khi tiếp tục.`);
    } catch (error) {
      selected.forEach(row => { row.status = 'Chưa kiểm tra'; });
      renderGrid();
      alert(`Kiểm tra chưa hoàn tất: ${error.message}`);
    } finally {
      window.obCheckRunning = false;
    }
  }

  async function runFillObAction() {
    const selected = window.gridData.filter(r => String(r.material_code || '').trim().length > 0 && normalizeStatusKey(r.status) === 'READY');
    if (selected.length === 0) {
      alert('Chỉ các dòng ở trạng thái "Sẵn sàng điền" mới có thể Điền OB!');
      return;
    }
    if (selected.some(row => !String(row.ob_number || '').trim())) {
      alert('Nhập số OB thật vào cột OB và đối chiếu trong SAP trước khi tiếp tục. Phần mềm không tự tạo số OB.');
      return;
    }
    if (!confirm('Anh/chị đã đối chiếu các số OB này trong SAP và xác nhận chúng tồn tại, đúng mã hàng và số lượng?')) return;
    selected.forEach(row => { row.status = 'Chờ pick / CF'; });
    renderGrid();
  }

  async function runCreatePickAction() {
    const readyItems = window.gridData.filter(r => String(r.material_code || '').trim().length > 0 && r.ob_number && normalizeStatusKey(r.status) === 'WAITING_PICK');
    if (readyItems.length === 0) {
      alert('Chỉ các dòng đã có OB và ở trạng thái "Chờ pick / CF" mới được Ra đơn pick!');
      return;
    }

    if (window.pickCreateRunning) return;
    window.pickCreateRunning = true;
    try {
      const res = await fetch('/api/pick-orders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          branch_tab: window.currentBranch,
          shop: window.currentBranch.includes('BODY') ? 'BODY' : 'EDC',
          model: readyItems[0].model || 'VF8',
          items: readyItems
        })
      });
      if (res.ok) {
        const data = await res.json();
        alert(`Đã sinh Đơn Pick thành công: ${data.pick_order_code}\nTrạng thái: "Chờ người dùng CF".\nTuyệt đối không tự động CF ngầm!`);
        loadPickOrders();
        renderGrid();
      } else {
        const error = await res.json();
        alert(`Lỗi tạo đơn pick: ${error.detail || 'Không thể tạo đơn'}`);
      }
    } catch (err) {
      alert(`Lỗi tạo đơn pick: ${err}`);
    } finally {
      window.pickCreateRunning = false;
    }
  }
}
