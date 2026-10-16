// File-based SAP pilot: every import and reconciliation is explicitly operator-triggered.
(() => {
  const byId = id => document.getElementById(id);
  async function responseData(response) {
    const data = await response.json();
    if (!response.ok) throw new Error(typeof data.detail === 'string' ? data.detail : JSON.stringify(data.detail));
    return data;
  }
  async function health() {
    const label = byId('sap-status-label');
    if (!label) return;
    try {
      const data = await responseData(await fetch('/api/health'));
      label.textContent = data.sap_connection === 'CONNECTED'
        ? 'Đọc được phiên SAP • nhập file để đối soát'
        : 'SAP chưa kết nối • chế độ nhập file';
      label.closest('.status-badge')?.classList.toggle('sap-connected', data.sap_connection === 'CONNECTED');
    } catch (_) { label.textContent = 'Chưa lấy được trạng thái SAP'; }
  }
  async function orders() {
    const select = byId('pilot-pick');
    const previous = select.value;
    const values = await responseData(await fetch('/api/pick-orders'));
    select.replaceChildren();
    const blank = new Option('Chọn đơn pick cần đối soát', '');
    select.add(blank);
    for (const order of values) {
      const option = new Option(`${order.pick_order_code || order.pick_order_no} — ${order.status_label || order.status}`, String(order.id));
      option.dataset.code = order.pick_order_code || order.pick_order_no;
      select.add(option);
    }
    if ([...select.options].some(o => o.value === previous)) select.value = previous;
  }
  async function action(button, callback) {
    button.disabled = true;
    const status = byId('pilot-source-result');
    status.textContent = 'Đang xử lý…';
    try { status.textContent = await callback(); }
    catch (error) { status.textContent = `Chưa hoàn tất: ${error.message}`; }
    finally { button.disabled = false; }
  }
  function selectedPick() {
    const id = byId('pilot-pick').value;
    if (!id) throw new Error('Chọn đơn pick trước');
    return id;
  }
  async function upload(kind) {
    const file = byId(`pilot-file-${kind}`).files[0];
    if (!file) throw new Error(`Chọn file ${kind.toUpperCase()}`);
    let url;
    if (kind === 'mb52') {
      url = `/api/stock/import-file?${new URLSearchParams({filename: file.name})}`;
    } else {
      const pick = selectedPick();
      const captured = byId('pilot-captured-at').value;
      if (!captured) throw new Error('Nhập thời điểm xuất báo cáo từ SAP');
      if (!byId('pilot-complete').checked) throw new Error('Xác nhận đã xuất đầy đủ OB của đơn pick');
      url = `/api/pick-orders/${pick}/sources/${kind}?${new URLSearchParams({filename: file.name, captured_at: captured, complete: 'true'})}`;
    }
    const result = await responseData(await fetch(url, {method: 'POST', headers: {'Content-Type': 'application/octet-stream'}, body: file}));
    if (kind === 'mb52') {
      const circle = byId('step-circle-1');
      if (circle) { circle.textContent = '✓'; circle.closest('.step-item')?.classList.add('completed'); }
      return `Đã tải snapshot MB52: ${result.total_rows} dòng. Snapshot mới thay thế dữ liệu đang dùng của từng Plant/SLoc, lịch sử được giữ riêng.`;
    }
    return `Đã tải ${kind.toUpperCase()}: ${result.row_count} dòng. Chưa đối soát; tải đủ hai nguồn rồi bấm Đối soát.`;
  }
  document.addEventListener('DOMContentLoaded', () => {
    health();
    const panel = document.createElement('section');
    panel.className = 'pilot-sources';
    panel.style.cssText = 'margin:16px 24px;padding:16px;border:1px solid #cbd5e1;border-radius:8px;background:white;';
    panel.innerHTML = `<details open><summary style="font-weight:600;cursor:pointer">⚡ Đối soát MB51 & Nguồn trực tiếp từ SAP</summary>
      <p style="margin:8px 0;font-size:13px;color:#475569">Luồng chính lấy MB51 và nguồn trực tiếp từ SAP theo đơn/phạm vi; không bắt buộc chọn file MB51 thủ công.</p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:12px">
        <select id="pilot-pick" style="padding:6px 10px;min-width:280px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;"></select>
        <button class="btn btn-secondary" id="pilot-refresh">Làm mới đơn pick</button>
        <button class="btn btn-primary" id="pilot-direct-reconcile-sap">⚡ Đối soát MB51 trực tiếp từ SAP</button>
        <button class="btn btn-secondary" id="pilot-export">Xuất / hoàn thiện BBBG</button>
      </div>
      <div id="pilot-source-result" role="status" aria-live="polite" style="font-size:13px;margin-top:12px;"></div>
      <details style="margin-top:14px;padding-top:10px;border-top:1px dashed #e2e8f0;">
        <summary style="font-size:12px;color:#64748b;cursor:pointer">📂 Tùy chọn dự phòng: Nạp file MB51/VL06I thủ công nếu chưa có phiên SAP live</summary>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:8px">
          <label style="font-size:12px">Thời điểm xuất nguồn <input type="datetime-local" id="pilot-captured-at"></label>
          <label style="font-size:12px"><input type="checkbox" id="pilot-complete"> Đã xuất đủ OB của đơn pick</label>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:8px">
          <label style="font-size:12px">MB51 <input id="pilot-file-mb51" type="file" accept=".csv,.tsv,.txt,.xlsx,.xlsm"></label>
          <button class="btn btn-secondary" id="pilot-upload-mb51" style="font-size:12px;padding:3px 8px;">Tải MB51</button>
          <label style="font-size:12px">VL06I <input id="pilot-file-vl06i" type="file" accept=".csv,.tsv,.txt,.xlsx,.xlsm"></label>
          <button class="btn btn-secondary" id="pilot-upload-vl06i" style="font-size:12px;padding:3px 8px;">Tải VL06I</button>
          <button class="btn btn-secondary" id="pilot-reconcile" style="font-size:12px;padding:3px 8px;">Đối soát từ file nạp</button>
          <label style="font-size:12px;margin-left:12px;">MB52 <input id="pilot-file-mb52" type="file" accept=".csv,.tsv,.txt,.xlsx,.xlsm"></label>
          <button class="btn btn-secondary" id="pilot-upload-mb52" style="font-size:12px;padding:3px 8px;">Nạp MB52</button>
        </div>
      </details>
    </details>`;
    const main = document.querySelector('.main-wrapper');
    main?.insertBefore(panel, main.querySelector('.stepper-container')?.nextSibling || null);
    orders().catch(error => { byId('pilot-source-result').textContent = error.message; });
    for (const kind of ['mb52', 'mb51', 'vl06i']) {
      const button = byId(`pilot-upload-${kind}`);
      button.onclick = () => action(button, () => upload(kind));
    }
    const refresh = byId('pilot-refresh');
    refresh.onclick = () => action(refresh, async () => { await orders(); return 'Đã làm mới đơn pick.'; });
    const directReconcile = byId('pilot-direct-reconcile-sap');
    if (directReconcile) {
      directReconcile.onclick = () => action(directReconcile, async () => {
        const pickId = selectedPick();
        const res = await responseData(await fetch(`/api/pick-orders/${pickId}/direct-reconcile-sap`, { method: 'POST' }));
        const summary = res.operator_summary || {};
        const isOk = summary.is_valid || res.status === 'SUCCESS';
        const actionsHtml = (summary.pending_actions || []).map(a => `<li>${a}</li>`).join('');
        const discHtml = (summary.discrepancies || []).map(d => `<li>${d.error || JSON.stringify(d)}</li>`).join('');
        const techJson = JSON.stringify(res.technical_details || res, null, 2);
        const resultEl = byId('pilot-source-result');
        resultEl.innerHTML = `
          <div style="background:${isOk ? '#f0fdf4' : '#fffbeb'};border:1px solid ${isOk ? '#86efac' : '#fde68a'};border-radius:6px;padding:12px;margin-top:6px;">
            <div style="font-weight:700;font-size:14px;color:${isOk ? '#166534' : '#92400e'};">
              ${isOk ? '✓ ' : '⚠️ '}${summary.result_label || res.message || 'Kết quả đối soát'}
            </div>
            <div style="font-size:13px;margin-top:6px;color:#334155;">
              Khớp: <strong>${summary.matched_count || 0}</strong> khóa dữ liệu • Trạng thái: <code>${res.status}</code>
            </div>
            ${discHtml ? `<div style="font-size:13px;margin-top:6px;color:#b91c1c;"><strong>Sai lệch phát hiện:</strong><ul style="margin:4px 0 0 16px;">${discHtml}</ul></div>` : ''}
            ${actionsHtml ? `<div style="font-size:13px;margin-top:6px;color:#475569;"><strong>Việc cần xử lý:</strong><ul style="margin:4px 0 0 16px;">${actionsHtml}</ul></div>` : ''}
            <details style="margin-top:10px;">
              <summary style="font-size:12px;color:#2563eb;cursor:pointer;">🔍 Xem chi tiết kỹ thuật (ẩn mặc định)</summary>
              <pre style="font-size:11px;background:#f8fafc;padding:8px;border-radius:4px;overflow-x:auto;max-height:160px;margin-top:4px;">${techJson}</pre>
            </details>
          </div>
        `;
        return isOk
          ? `Đối soát thành công! Đã khớp ${summary.matched_count || 0} dòng. Đủ điều kiện lập BBBG.`
          : (summary.result_label || 'Đối soát chưa hoàn tất.');
      });
    }
    const reconcile = byId('pilot-reconcile');
    reconcile.onclick = () => action(reconcile, async () => {
      const result = await responseData(await fetch(`/api/pick-orders/${selectedPick()}/reconcile-uploaded`, {method: 'POST'}));
      return `Đối soát đạt: ${result.rows.length} khóa OB–mã hàng–batch, đơn vị và lượng CF khớp. Có thể xuất BBBG từ nguồn đã nhập; chưa xác nhận GOS.`;
    });
    byId('pilot-export').onclick = () => {
      try {
        selectedPick();
        const code = byId('pilot-pick').selectedOptions[0].dataset.code;
        window.openBbbgModal(code);
      } catch (error) { byId('pilot-source-result').textContent = error.message; }
    };
  });
})();
