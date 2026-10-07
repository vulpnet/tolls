// Copy dữ liệu theo CỘT hoặc theo HÀNG cho mọi bảng kết quả trong bộ tools — dùng chung 1 lần cho
// nhiều tool (CompareText, ExcelToTable, TableBuilder, ExcelSplitMerge, CsvConverter, IpLookup,
// JsonConvert...) thay vì viết lặp lại từng tool, vì các bảng này re-render liên tục (đổi filter/sort/
// sheet/option) nên không thể bind trực tiếp vào từng <th>/<td> — phải dùng event delegation trên
// container cha và đọc DOM hiện tại (đúng với dữ liệu đang hiển thị, kể cả khi bị truncate).
//
// Cách dùng: enableTableCopy('#out') — gọi 1 lần sau khi DOM đã có container, không cần gọi lại sau
// mỗi lần re-render vì toàn bộ gắn qua delegation.
(function(){
  function cellText(el){
    // ưu tiên input/select (TableBuilder dùng <input> trong <td>) trước khi lấy textContent thô
    const input = el.querySelector(':scope > input, :scope > select');
    if (input) return String(input.value ?? '');
    // bỏ qua nút copy inject vào chính ô đó khi lấy text
    const clone = el.cloneNode(true);
    clone.querySelectorAll('.tblCopyBtn').forEach(b => b.remove());
    return (clone.textContent || '').replace(/\s+/g,' ').trim();
  }

  function isRownumCell(td){
    return td.classList.contains('rownum') || td.classList.contains('rowh');
  }

  // fix: fallback execCommand('copy') qua textarea ẩn — navigator.clipboard.writeText() có thể bị chặn
  // silently (promise reject không rõ lý do) khi trang chạy trong iframe thiếu allow="clipboard-write",
  // hoặc context không secure. Không có fallback thì nút vẫn hiện nhưng bấm hoàn toàn im lặng không copy
  // được, trông như tool bị lỗi mà không có cách nào biết tại sao.
  function legacyCopy(text){
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    let ok = false;
    try{ ok = document.execCommand('copy'); }catch(e){ ok = false; }
    document.body.removeChild(ta);
    return ok;
  }

  function flashBtn(btn){
    if (!btn) return;
    const old = btn.textContent;
    btn.textContent = '✓';
    setTimeout(() => { btn.textContent = old; }, 900);
  }

  function copyToClipboard(text, btn){
    if (navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(text).then(() => flashBtn(btn)).catch(() => {
        if (legacyCopy(text)) flashBtn(btn);
      });
    } else if (legacyCopy(text)){
      flashBtn(btn);
    }
  }

  function makeCopyBtn(title){
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'tblCopyBtn';
    b.textContent = '📋';
    b.title = title;
    b.addEventListener('click', e => {
      e.stopPropagation();
      e.preventDefault();
    });
    return b;
  }

  function copyColumn(table, colIndex, btn){
    const rows = table.querySelectorAll(':scope > tbody > tr, :scope > tr');
    const vals = [];
    rows.forEach(tr => {
      const cells = tr.children;
      if (colIndex >= cells.length) return;
      const cell = cells[colIndex];
      if (cell.tagName === 'TH') return; // bỏ header lặp (vd filter row dùng th) khỏi dữ liệu copy
      vals.push(cellText(cell));
    });
    copyToClipboard(vals.join('\n'), btn);
  }

  function copyRow(tr, btn){
    const cells = [...tr.children].filter(c => !isRownumCell(c) && c.tagName !== 'TH');
    const vals = cells.map(cellText);
    copyToClipboard(vals.join('\t'), btn);
  }

  function colIndexOf(th){
    return [...th.parentElement.children].indexOf(th);
  }

  function attachToHeader(th){
    if (th.querySelector('.tblCopyBtn')) return;
    if (isRownumCell(th) || th.classList.contains('rownumHead')) return; // cột # không có dữ liệu để copy
    const btn = makeCopyBtn('Copy cả cột này');
    btn.addEventListener('click', () => {
      const table = th.closest('table');
      copyColumn(table, colIndexOf(th), btn);
    });
    th.appendChild(btn);
  }

  function attachToRowHandle(cell){
    if (cell.querySelector('.tblCopyBtn')) return;
    const btn = makeCopyBtn('Copy cả hàng này');
    btn.addEventListener('click', () => {
      copyRow(cell.closest('tr'), btn);
    });
    cell.appendChild(btn);
  }

  function enableTableCopy(containerSelector){
    const container = typeof containerSelector === 'string' ? document.querySelector(containerSelector) : containerSelector;
    if (!container) return;

    container.addEventListener('mouseover', e => {
      const th = e.target.closest('th');
      if (th && container.contains(th)){ attachToHeader(th); return; }
      const rownumCell = e.target.closest('td.rownum, td.rowh');
      if (rownumCell && container.contains(rownumCell)){ attachToRowHandle(rownumCell); return; }
    });
  }

  window.enableTableCopy = enableTableCopy;
})();
