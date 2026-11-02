/* CompareText — rewritten engine (non-obfuscated), optimized for large text/file input.
   Preserves all existing features: line diff (side-by-side/unified), set mode,
   table mode (Excel cell-by-cell), rich mode (Word block diff / PDF page images).
   Performance changes vs. previous build:
   - 3-tier debounce instead of 2-tier (300/500/900ms) based on input size
   - render generation token: a stale render in flight is discarded if newer input arrived
   - checkbox state cached once per render() instead of re-read per comparison
   - inlineDiff() skips word-level diff for pathologically long single lines (token cap)
   - Excel import capped at MAX_EXCEL_ROWS rows/sheet, file-size warn/block thresholds
   - Word import capped at MAX_RICH_BLOCKS blocks (falls back to plain text diff beyond that)
   - PDF import rasterizes only the first MAX_PDF_PAGES pages by default, with a
     "load more pages" action instead of rasterizing an unbounded page count up front
*/
(function () {
'use strict';

const $ = id => document.getElementById(id);

const I18N = {
  vi: {
    title: "Compare Text",
    optIgnoreCase: "Bỏ qua hoa/thường",
    optIgnoreWs: "Bỏ qua khoảng trắng",
    optIgnoreBlank: "Bỏ qua dòng trống",
    optSideBySide: "Side-by-side",
    optOnlyDiff: "Chỉ hiện dòng khác",
    optSetMode: "So sánh danh sách (A/B khác gì nhau, không quan tâm thứ tự)",
    optTableMode: "So sánh dạng khối (Excel theo hàng/cột, Word/PDF giữ định dạng gốc) — chỉ khi A và B cùng loại",
    btnCompare: "So sánh",
    btnSwap: "Đổi chỗ",
    btnClear: "Xoá",
    navPrevTitle: "Đoạn khác trước đó (Alt+↑ / Shift+Enter)",
    navNextTitle: "Đoạn khác kế tiếp (Alt+↓ / Enter)",
    labelTextA: "Text A (gốc)",
    labelTextB: "Text B (mới)",
    btnImport: "📄 Import file (Excel/Word/PDF)",
    placeholderA: "Dán text A, hoặc bấm Import file để nạp từ Excel/Word/PDF...",
    placeholderB: "Dán text B, hoặc bấm Import file để nạp từ Excel/Word/PDF...",
    sheetALabel: "Sheet A:",
    sheetBLabel: "Sheet B:",
    optFirstRowHeader: "Dòng đầu là header",
    emptyPrompt: 'Nhập text rồi bấm <b>So sánh</b> (hoặc Ctrl+Enter).',
    emptyNoContent: 'Không có nội dung.',
    navPos: (i, n) => n ? (i < 0 ? "– / " + n : (i + 1) + " / " + n) : "0 / 0",
    unsupportedFormat: ext => "Định dạng file không được hỗ trợ: ." + ext,
    sheetUnreadable: name => "--- Sheet: " + name + " (không đọc được) ---",
    sheetHeader: name => "--- Sheet: " + name + " ---",
    pageHeader: n => "--- Trang " + n + " ---",
    readingFile: name => "Đang đọc " + name + "...",
    loadedFile: (name, lines) => "Đã nạp: " + name + " (" + lines + " dòng)",
    errReadingFile: msg => "Lỗi đọc file: " + msg,
    copied: "Đã copy!",
    setSame: n => "<b>Hai danh sách có cùng tập giá trị</b> (không quan tâm thứ tự/số lần lặp) — " + n + " giá trị duy nhất.",
    setDiff: (a, b, common, onlyA, onlyB) => "A: " + a + " giá trị duy nhất · B: " + b + " giá trị duy nhất · <b>" + common + "</b> chung · <b class=\"del\">" + onlyA + "</b> chỉ có ở A · <b class=\"add\">" + onlyB + "</b> chỉ có ở B",
    onlyEmpty: "(không có)",
    onlyALabel: "Chỉ có ở A",
    onlyBLabel: "Chỉ có ở B",
    labelPaneA: "TEXT A (GỐC)",
    labelPaneB: "TEXT B (MỚI)",
    copyBtn: "Copy",
    tableSame: (r, c) => "<b>Hai bảng giống nhau hoàn toàn</b> (theo option đang chọn) — " + r + " dòng × " + c + " cột.",
    tableDiff: (r, c, diffCells, onlyA, onlyB) => r + " dòng × " + c + " cột · <b class=\"del\">" + diffCells + "</b> ô khác nhau" + (onlyA ? " · <b class=\"del\">" + onlyA + "</b> dòng chỉ có ở A" : '') + (onlyB ? " · <b class=\"add\">" + onlyB + "</b> dòng chỉ có ở B" : ''),
    tableTruncated: (total, shown) => "⚠️ Bảng có " + total + " dòng — đang chỉ hiện " + shown + " dòng đầu để tránh treo trình duyệt.",
    lineTruncated: (total, shown) => "⚠️ Kết quả có " + total + " hàng — đang chỉ hiện " + shown + " hàng đầu để tránh treo trình duyệt.",
    showAllBtn: "Hiện tất cả",
    richZoomOut: "Thu nhỏ",
    richZoomIn: "Phóng to",
    richZoomReset: "100%",
    richPdfImageNote: "PDF hiển thị bằng ảnh trang thật (giữ nguyên layout/font gốc) — không tự động tô đỏ/xanh chỗ khác nhau, hãy tự đối chiếu 2 bên bằng mắt.",
    richSame: n => "<b>Nội dung giống nhau hoàn toàn</b> — " + n + " khối (đoạn/dòng).",
    richDiff: (a, b, onlyA, onlyB) => "A: " + a + " khối · B: " + b + " khối · <b class=\"del\">" + onlyA + "</b> khối chỉ có/khác ở A · <b class=\"add\">" + onlyB + "</b> khối chỉ có/khác ở B",
    bigTextWarning: n => "Văn bản đang so sánh rất lớn (>" + n.toLocaleString('vi-VN') + " ký tự), quá trình so sánh có thể chậm.",
    lineSame: n => "<b>Giống nhau hoàn toàn</b> (theo option đang chọn) — " + n + " dòng.",
    lineDiff: (a, b, del, add, mod) => "A: " + a + " dòng · B: " + b + " dòng · <b class=\"del\">-" + del + "</b> xoá · <b class=\"add\">+" + add + "</b> thêm · <b>" + mod + "</b> sửa",
    skippedLines: n => "… " + n + " dòng giống nhau",
    // new messages for the import/size-guard rewrite
    excelTooBig: (mb) => "File Excel quá lớn (" + mb + "MB, giới hạn 150MB) — vui lòng tách nhỏ file trước khi import.",
    excelWarnBig: (mb) => "File Excel khá lớn (" + mb + "MB) — việc đọc có thể mất vài giây.",
    excelRowsCapped: (sheet, total, cap) => "Sheet \"" + sheet + "\" có " + total.toLocaleString('vi-VN') + " dòng — chỉ đọc " + cap.toLocaleString('vi-VN') + " dòng đầu để tránh treo trình duyệt.",
    richBlocksCapped: (total, cap) => "Tài liệu có " + total.toLocaleString('vi-VN') + " đoạn — chỉ so sánh dạng khối (block) " + cap.toLocaleString('vi-VN') + " đoạn đầu, phần còn lại vẫn có trong văn bản thuần (không tô màu khác biệt).",
    pdfPagesCapped: (total, shown) => "PDF có " + total + " trang — đang hiện " + shown + " trang đầu.",
    pdfLoadMore: n => "Tải thêm " + n + " trang",
    hugeTextWarning: n => "Văn bản quá lớn (>" + n.toLocaleString('vi-VN') + " ký tự) — so sánh có thể rất chậm hoặc làm treo trình duyệt. Cân nhắc so sánh từng phần.",
  },
  en: {
    title: "Compare Text",
    optIgnoreCase: "Ignore case",
    optIgnoreWs: "Ignore whitespace",
    optIgnoreBlank: "Ignore blank lines",
    optSideBySide: "Side-by-side",
    optOnlyDiff: "Show only differing lines",
    optSetMode: "List comparison (what differs between A/B, ignoring order)",
    optTableMode: "Block comparison (Excel by row/column, Word/PDF keeps original formatting) — only when A and B are the same type",
    btnCompare: "Compare",
    btnSwap: "Swap",
    btnClear: "Clear",
    navPrevTitle: "Previous difference (Alt+↑ / Shift+Enter)",
    navNextTitle: "Next difference (Alt+↓ / Enter)",
    labelTextA: "Text A (original)",
    labelTextB: "Text B (new)",
    btnImport: "📄 Import file (Excel/Word/PDF)",
    placeholderA: "Paste text A, or click Import file to load from Excel/Word/PDF...",
    placeholderB: "Paste text B, or click Import file to load from Excel/Word/PDF...",
    sheetALabel: "Sheet A:",
    sheetBLabel: "Sheet B:",
    optFirstRowHeader: "First row is header",
    emptyPrompt: "Enter text then click <b>Compare</b> (or Ctrl+Enter).",
    emptyNoContent: 'No content.',
    navPos: (i, n) => n ? (i < 0 ? "– / " + n : (i + 1) + " / " + n) : "0 / 0",
    unsupportedFormat: ext => "Unsupported file format: ." + ext,
    sheetUnreadable: name => "--- Sheet: " + name + " (unreadable) ---",
    sheetHeader: name => "--- Sheet: " + name + " ---",
    pageHeader: n => "--- Page " + n + " ---",
    readingFile: name => "Reading " + name + "...",
    loadedFile: (name, lines) => "Loaded: " + name + " (" + lines + " lines)",
    errReadingFile: msg => "Error reading file: " + msg,
    copied: "Copied!",
    setSame: n => "<b>Both lists have the same set of values</b> (ignoring order/duplicates) — " + n + " unique values.",
    setDiff: (a, b, common, onlyA, onlyB) => "A: " + a + " unique values · B: " + b + " unique values · <b>" + common + "</b> common · <b class=\"del\">" + onlyA + "</b> only in A · <b class=\"add\">" + onlyB + "</b> only in B",
    onlyEmpty: "(none)",
    onlyALabel: "Only in A",
    onlyBLabel: "Only in B",
    labelPaneA: "TEXT A (ORIGINAL)",
    labelPaneB: "TEXT B (NEW)",
    copyBtn: "Copy",
    tableSame: (r, c) => "<b>Both tables are identical</b> (per selected options) — " + r + " rows × " + c + " columns.",
    tableDiff: (r, c, diffCells, onlyA, onlyB) => r + " rows × " + c + " columns · <b class=\"del\">" + diffCells + "</b> differing cells" + (onlyA ? " · <b class=\"del\">" + onlyA + "</b> rows only in A" : '') + (onlyB ? " · <b class=\"add\">" + onlyB + "</b> rows only in B" : ''),
    tableTruncated: (total, shown) => "⚠️ Table has " + total + " rows — showing only the first " + shown + " rows to avoid freezing the browser.",
    lineTruncated: (total, shown) => "⚠️ Result has " + total + " rows — showing only the first " + shown + " rows to avoid freezing the browser.",
    showAllBtn: "Show all",
    richZoomOut: "Zoom out",
    richZoomIn: "Zoom in",
    richZoomReset: "100%",
    richPdfImageNote: "PDF is shown as real page images (keeps original layout/fonts) — differences are NOT auto-highlighted, please compare both sides visually.",
    richSame: n => "<b>Content is identical</b> — " + n + " blocks (paragraphs/lines).",
    richDiff: (a, b, onlyA, onlyB) => "A: " + a + " blocks · B: " + b + " blocks · <b class=\"del\">" + onlyA + "</b> blocks only/different in A · <b class=\"add\">" + onlyB + "</b> blocks only/different in B",
    bigTextWarning: n => "The text being compared is very large (>" + n.toLocaleString('en-US') + " characters), comparison may be slow.",
    lineSame: n => "<b>Completely identical</b> (per selected options) — " + n + " lines.",
    lineDiff: (a, b, del, add, mod) => "A: " + a + " lines · B: " + b + " lines · <b class=\"del\">-" + del + "</b> deleted · <b class=\"add\">+" + add + "</b> added · <b>" + mod + "</b> modified",
    skippedLines: n => "… " + n + " identical lines",
    excelTooBig: (mb) => "Excel file is too large (" + mb + "MB, limit 150MB) — please split the file before importing.",
    excelWarnBig: (mb) => "Excel file is fairly large (" + mb + "MB) — reading it may take a few seconds.",
    excelRowsCapped: (sheet, total, cap) => "Sheet \"" + sheet + "\" has " + total.toLocaleString('en-US') + " rows — only the first " + cap.toLocaleString('en-US') + " rows were read to avoid freezing the browser.",
    richBlocksCapped: (total, cap) => "Document has " + total.toLocaleString('en-US') + " blocks — block comparison only covers the first " + cap.toLocaleString('en-US') + " blocks; the rest is still included in the plain text (without diff highlighting).",
    pdfPagesCapped: (total, shown) => "PDF has " + total + " pages — showing the first " + shown + ".",
    pdfLoadMore: n => "Load " + n + " more pages",
    hugeTextWarning: n => "Text is extremely large (>" + n.toLocaleString('en-US') + " characters) — comparison may be very slow or freeze the browser. Consider comparing in smaller parts.",
  },
};
function t() { return I18N[window.__toolsLang === 'en' ? 'en' : 'vi']; }

// ---------------------------------------------------------------------------
// Tunable limits
// ---------------------------------------------------------------------------
const BIG_TEXT_THRESHOLD = 300000;      // debounce bumps from 300ms to 500ms past this
const HUGE_TEXT_THRESHOLD = 1200000;    // debounce bumps to 900ms + visible warning banner past this
const HARD_TEXT_LIMIT = 8000000;        // refuse to diff past this many chars per side (would hang the tab)
const LINE_MODE_RENDER_LIMIT = 5000;    // DOM rows rendered by default in line mode
const TABLE_MODE_RENDER_LIMIT = 2000;   // DOM rows rendered by default in table mode
const DIFF_PRODUCT_CUTOFF = 4000000;    // linesA*linesB above this => naive diff fallback
const MYERS_D_CAP = 20000;              // iteration cap for the linear-space Myers path search
const INLINE_DIFF_TOKEN_CAP = 2000;     // per-line token count above which word-diff is skipped

const MAX_EXCEL_WARN_BYTES = 50 * 1024 * 1024;   // 50MB: show a "this may take a while" warning
const MAX_EXCEL_BLOCK_BYTES = 150 * 1024 * 1024; // 150MB: refuse outright (same ceiling as ExcelToTable)
const MAX_EXCEL_ROWS = 20000;           // rows read per sheet (sheet_to_json + sheet_to_csv)
const MAX_RICH_BLOCKS = 5000;           // Word blocks compared at block-diff granularity
const MAX_PDF_PAGES_INITIAL = 50;       // PDF pages rasterized on first import
const PDF_PAGE_BATCH = 50;              // pages rasterized per "load more" click
const PDF_RENDER_SCALE = 1.5;

// ---------------------------------------------------------------------------
// Diff engine
// ---------------------------------------------------------------------------

// Greedy O(ND) Myers diff over two arrays, returns an edit script of
// {type:'=', a, b} / {type:'-', a} / {type:'+', b} items (indices into
// the original arrays). Falls back to a naive full-delete+insert when the
// problem is too large to diff in bounded time/memory.
function diffArrays(a, b, eq) {
  const n = a.length, m = b.length;
  if (n === 0 && m === 0) return [];
  if (n === 0) return b.map((_, j) => ({ type: '+', b: j }));
  if (m === 0) return a.map((_, i) => ({ type: '-', a: i }));

  // Trim common prefix/suffix first — shrinks the DP problem a lot for
  // typical edits (most of a large document is usually unchanged).
  let prefix = 0;
  const maxPrefix = Math.min(n, m);
  while (prefix < maxPrefix && eq(a[prefix], b[prefix])) prefix++;
  let suffix = 0;
  const maxSuffix = Math.min(n, m) - prefix;
  while (suffix < maxSuffix && eq(a[n - 1 - suffix], b[m - 1 - suffix])) suffix++;

  const a2 = a.slice(prefix, n - suffix);
  const b2 = b.slice(prefix, m - suffix);
  const n2 = a2.length, m2 = b2.length;

  let middle;
  if (n2 === 0 && m2 === 0) {
    middle = [];
  } else if (n2 === 0) {
    middle = b2.map((_, j) => ({ type: '+', b: j }));
  } else if (m2 === 0) {
    middle = a2.map((_, i) => ({ type: '-', a: i }));
  } else if (n2 * m2 > DIFF_PRODUCT_CUTOFF) {
    // Too large to diff properly within a reasonable time/memory budget —
    // fall back to "everything changed" rather than hanging the tab.
    middle = a2.map((_, i) => ({ type: '-', a: i }))
      .concat(b2.map((_, j) => ({ type: '+', b: j })));
  } else {
    middle = myersPath(a2, b2, eq);
    if (middle === null) {
      middle = a2.map((_, i) => ({ type: '-', a: i }))
        .concat(b2.map((_, j) => ({ type: '+', b: j })));
    }
  }

  const out = [];
  for (let i = 0; i < prefix; i++) out.push({ type: '=', a: i, b: i });
  for (const op of middle) {
    if (op.type === '=') out.push({ type: '=', a: op.a + prefix, b: op.b + prefix });
    else if (op.type === '-') out.push({ type: '-', a: op.a + prefix });
    else out.push({ type: '+', b: op.b + prefix });
  }
  for (let i = 0; i < suffix; i++) {
    out.push({ type: '=', a: n - suffix + i, b: m - suffix + i });
  }
  return out;
}

// Linear-space greedy Myers diff (the "Myers' O(ND) Difference Algorithm" /
// greedy forward search). Bounded to MYERS_D_CAP edit-script-length
// iterations; returns null if no path is found within that bound (caller
// falls back to naive diff).
function myersPath(a, b, eq) {
  const n = a.length, m = b.length;
  const max = n + m;
  const vSize = 2 * max + 1;
  const v = new Int32Array(vSize);
  const offset = max;
  const trace = [];

  for (let d = 0; d <= Math.min(max, MYERS_D_CAP); d++) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x;
      if (k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])) {
        x = v[offset + k + 1];
      } else {
        x = v[offset + k - 1] + 1;
      }
      let y = x - k;
      while (x < n && y < m && eq(a[x], b[y])) { x++; y++; }
      v[offset + k] = x;
      if (x >= n && y >= m) {
        return backtrack(trace, a, b, eq, d, offset, n, m);
      }
    }
  }
  return null;
}

function backtrack(trace, a, b, eq, d0, offset, n, m) {
  const ops = [];
  let x = n, y = m;
  for (let d = d0; d > 0; d--) {
    const v = trace[d];
    const k = x - y;
    let prevK;
    if (k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])) {
      prevK = k + 1;
    } else {
      prevK = k - 1;
    }
    const prevX = v[offset + prevK];
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) {
      ops.push({ type: '=', a: x - 1, b: y - 1 });
      x--; y--;
    }
    if (d > 0) {
      if (x === prevX) { ops.push({ type: '+', b: prevY }); }
      else { ops.push({ type: '-', a: prevX }); }
    }
    x = prevX; y = prevY;
  }
  const v0 = trace[0];
  while (x > 0 && y > 0 && eq(a[x - 1], b[y - 1])) { ops.push({ type: '=', a: x - 1, b: y - 1 }); x--; y--; }
  ops.reverse();
  return ops;
}

// Word tokenizer — Unicode-aware (handles Vietnamese diacritics).
const TOKEN_RE = /\s+|[\wÀ-ỹ]+|[^\s\wÀ-ỹ]/g;
function tokenize(text) {
  return text.match(TOKEN_RE) || [];
}

// Word-level inline diff between two changed lines. Guards against
// pathologically long single lines (e.g. minified JSON/SQL on one line) by
// skipping word-diff beyond INLINE_DIFF_TOKEN_CAP tokens — callers fall back
// to whole-line add/del coloring for such lines instead of either hanging
// or silently producing the "everything changed" naive fallback.
function inlineDiff(lineA, lineB, ignoreCase) {
  const tokA = tokenize(lineA);
  const tokB = tokenize(lineB);
  if (tokA.length > INLINE_DIFF_TOKEN_CAP || tokB.length > INLINE_DIFF_TOKEN_CAP) {
    return null; // caller renders whole-line del/add instead
  }
  const cmp = ignoreCase
    ? (x, y) => x.toLowerCase() === y.toLowerCase()
    : (x, y) => x === y;
  const ops = diffArrays(tokA, tokB, cmp);
  let htmlA = '', htmlB = '';
  for (const op of ops) {
    if (op.type === '=') {
      const s = escHtml(tokA[op.a]);
      htmlA += s; htmlB += s;
    } else if (op.type === '-') {
      htmlA += '<del>' + escHtml(tokA[op.a]) + '</del>';
    } else {
      htmlB += '<ins>' + escHtml(tokB[op.b]) + '</ins>';
    }
  }
  return { htmlA, htmlB };
}

function escHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escHtmlAttr(s) {
  return escHtml(s).replace(/"/g, '&quot;');
}

// ---------------------------------------------------------------------------
// Line normalization — options are captured ONCE per render() call and
// passed down as plain booleans, instead of re-reading checkboxes from the
// DOM inside every comparison call (which previously happened up to
// millions of times for large inputs).
// ---------------------------------------------------------------------------
function readOptions() {
  return {
    ignoreCase: $('ignoreCase').checked,
    ignoreWs: $('ignoreWs').checked,
    ignoreBlank: $('ignoreBlank').checked,
    sideBySide: $('sideBySide').checked,
    onlyDiff: $('onlyDiff').checked,
    setMode: $('setMode').checked,
    tableMode: !$('tableModeWrap').classList.contains('hidden') && $('tableMode').checked,
  };
}

function normLine(line, opts) {
  let s = line;
  if (opts.ignoreWs) s = s.trim().replace(/\s+/g, ' ');
  if (opts.ignoreCase) s = s.toLowerCase();
  return s;
}

function splitLines(text, opts) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  if (opts.ignoreBlank) return lines.filter(l => l.trim() !== '');
  return lines;
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let debounceTimer = null;
let renderGeneration = 0;       // bumped on every scheduleRun(); render() checks it stays current
let totalHunks = 0, curHunk = -1;
let workbookA = null, workbookB = null;   // {sheetNames, sheets:{name:{rows, csv, truncated, totalRows}}}
let richDocA = null, richDocB = null;     // {kind:'blocks', blocks:[{html,text}], truncated?} | {kind:'pages', pages:[{dataUrl,pageNum}], totalPages, loadedPages}
let bigTextWarned = false;

function showBigTextBannerIfNeeded(lenA, lenB) {
  const banner = $('bigTextBanner');
  const maxLen = Math.max(lenA, lenB);
  if (maxLen > HUGE_TEXT_THRESHOLD) {
    banner.innerHTML = t().hugeTextWarning(HUGE_TEXT_THRESHOLD);
    banner.classList.remove('hidden');
  } else if (maxLen > BIG_TEXT_THRESHOLD) {
    banner.textContent = t().bigTextWarning(BIG_TEXT_THRESHOLD);
    banner.classList.remove('hidden');
  } else {
    banner.classList.add('hidden');
  }
}

function scheduleRun() {
  clearTimeout(debounceTimer);
  renderGeneration++;
  const lenA = $('a').value.length, lenB = $('b').value.length;
  const maxLen = Math.max(lenA, lenB);
  const delay = maxLen > HUGE_TEXT_THRESHOLD ? 900 : (maxLen > BIG_TEXT_THRESHOLD ? 500 : 300);
  const myGen = renderGeneration;
  debounceTimer = setTimeout(() => render(myGen), delay);
}

['ignoreCase', 'ignoreWs', 'ignoreBlank', 'sideBySide', 'onlyDiff', 'setMode', 'tableMode'].forEach(id => {
  const el = $(id);
  if (!el) return;
  el.addEventListener('change', () => {
    if (id === 'setMode') updateSetModeUi();
    if ($('a').value || $('b').value) scheduleRun();
  });
});

function updateSetModeUi() {
  const setMode = $('setMode').checked;
  $('tableModeWrap').classList.toggle('hidden', setMode || !(workbookA && workbookB) && !(richDocA && richDocB));
  $('onlyDiff').closest('label').classList.toggle('hidden', setMode);
  $('sideBySide').closest('label').classList.toggle('hidden', setMode);
}

// ---------------------------------------------------------------------------
// Main render dispatcher. `gen` is the generation token captured when this
// render was scheduled — if a newer input arrived and bumped
// renderGeneration since then, this call is stale and bails out before
// doing any expensive work or touching the DOM.
// ---------------------------------------------------------------------------
function render(gen) {
  if (gen !== undefined && gen !== renderGeneration) return; // superseded by newer input
  const textA = $('a').value, textB = $('b').value;
  showBigTextBannerIfNeeded(textA.length, textB.length);

  if (!textA && !textB) {
    $('out').innerHTML = '<div class="empty">' + t().emptyPrompt + '</div>';
    $('stats').textContent = '';
    totalHunks = 0; curHunk = -1; updateNav();
    return;
  }

  if (Math.max(textA.length, textB.length) > HARD_TEXT_LIMIT) {
    $('out').innerHTML = '<div class="empty">' + escHtml(t().hugeTextWarning(HARD_TEXT_LIMIT)) + '</div>';
    $('stats').textContent = '';
    return;
  }

  const opts = readOptions();

  if (opts.setMode) return renderSetMode(textA, textB, opts, gen);
  if (opts.tableMode && workbookA && workbookB) return renderTableMode(opts, gen);
  if (opts.tableMode && richDocA && richDocB) return renderRichMode(opts, gen);
  return renderLineMode(textA, textB, opts, gen);
}

// ---------------------------------------------------------------------------
// Line mode (default): side-by-side or unified, with word-level inline diff
// on changed-line pairs.
// ---------------------------------------------------------------------------
function renderLineMode(textA, textB, opts, gen) {
  const rawA = splitLines(textA, opts), rawB = splitLines(textB, opts);
  const normA = rawA.map(l => normLine(l, opts));
  const normB = rawB.map(l => normLine(l, opts));
  const ops = diffArrays(normA, normB, (x, y) => x === y);
  if (gen !== undefined && gen !== renderGeneration) return;

  // Pair up adjacent '-' then '+' runs into modify-pairs for word-diff.
  const rows = [];
  let i = 0;
  let dels = 0, adds = 0, mods = 0;
  while (i < ops.length) {
    const op = ops[i];
    if (op.type === '=') {
      rows.push({ kind: 'same', a: op.a, b: op.b });
      i++;
      continue;
    }
    // collect a run of consecutive -/+ ops
    const runDel = [], runAdd = [];
    while (i < ops.length && ops[i].type === '-') { runDel.push(ops[i].a); i++; }
    while (i < ops.length && ops[i].type === '+') { runAdd.push(ops[i].b); i++; }
    const pairCount = Math.min(runDel.length, runAdd.length);
    mods += pairCount;
    dels += runDel.length - pairCount;
    adds += runAdd.length - pairCount;
    for (let k = 0; k < pairCount; k++) rows.push({ kind: 'mod', a: runDel[k], b: runAdd[k] });
    for (let k = pairCount; k < runDel.length; k++) rows.push({ kind: 'del', a: runDel[k] });
    for (let k = pairCount; k < runAdd.length; k++) rows.push({ kind: 'add', b: runAdd[k] });
  }

  const sameCount = rows.filter(r => r.kind === 'same').length;
  if (dels === 0 && adds === 0 && mods === 0) {
    $('stats').innerHTML = t().lineSame(rawA.length);
  } else {
    $('stats').innerHTML = t().lineDiff(rawA.length, rawB.length, dels, adds, mods);
  }

  const showAll = $('showAllLines').dataset.active === '1';
  const limit = showAll ? Infinity : LINE_MODE_RENDER_LIMIT;
  let visibleRows = rows;
  if (opts.onlyDiff) visibleRows = rows.filter(r => r.kind !== 'same');
  const truncated = visibleRows.length > limit;
  const shownRows = truncated ? visibleRows.slice(0, limit) : visibleRows;

  const sideBySide = opts.sideBySide;
  const htmlParts = [];
  htmlParts.push('<table' + (sideBySide ? ' class="side"' : '') + '>');

  let hunkIndex = -1;
  let prevWasDiff = false;
  const hunkRowEls = []; // row index -> hunk number, for nav (built after insertion via data attr)

  for (const row of shownRows) {
    const isDiff = row.kind !== 'same';
    if (isDiff && !prevWasDiff) hunkIndex++;
    prevWasDiff = isDiff;
    const hunkAttr = isDiff ? ' data-hunk="' + hunkIndex + '"' : '';

    if (row.kind === 'same') {
      const lineA = rawA[row.a];
      if (sideBySide) {
        htmlParts.push('<tr><td class="num">' + (row.a + 1) + '</td><td class="pane">' + escHtml(lineA) + '</td><td class="num">' + (row.b + 1) + '</td><td class="pane">' + escHtml(rawB[row.b]) + '</td></tr>');
      } else {
        htmlParts.push('<tr><td class="num">' + (row.a + 1) + '</td><td>' + escHtml(lineA) + '</td></tr>');
      }
    } else if (row.kind === 'mod') {
      const lineA = rawA[row.a], lineB = rawB[row.b];
      const inline = inlineDiff(lineA, lineB, opts.ignoreCase);
      const cellA = inline ? inline.htmlA : escHtml(lineA);
      const cellB = inline ? inline.htmlB : escHtml(lineB);
      if (sideBySide) {
        htmlParts.push('<tr class="del"' + hunkAttr + '><td class="num">' + (row.a + 1) + '</td><td class="pane">' + cellA + '</td><td class="num">' + (row.b + 1) + '</td><td class="pane add-pane"></td></tr>'.replace('<td class="pane add-pane"></td>', '<td class="pane">' + cellB + '</td>').replace('<tr class="del"', '<tr class="del"'));
      } else {
        htmlParts.push('<tr class="del"' + hunkAttr + '><td class="num">' + (row.a + 1) + '</td><td>' + cellA + '</td></tr>');
        htmlParts.push('<tr class="add"' + hunkAttr + '><td class="num">' + (row.b + 1) + '</td><td>' + cellB + '</td></tr>');
      }
    } else if (row.kind === 'del') {
      const lineA = rawA[row.a];
      if (sideBySide) {
        htmlParts.push('<tr class="del"' + hunkAttr + '><td class="num">' + (row.a + 1) + '</td><td class="pane">' + escHtml(lineA) + '</td><td class="num"></td><td class="pane"></td></tr>');
      } else {
        htmlParts.push('<tr class="del"' + hunkAttr + '><td class="num">' + (row.a + 1) + '</td><td>' + escHtml(lineA) + '</td></tr>');
      }
    } else { // add
      const lineB = rawB[row.b];
      if (sideBySide) {
        htmlParts.push('<tr class="add"' + hunkAttr + '><td class="num"></td><td class="pane"></td><td class="num">' + (row.b + 1) + '</td><td class="pane">' + escHtml(lineB) + '</td></tr>');
      } else {
        htmlParts.push('<tr class="add"' + hunkAttr + '><td class="num">' + (row.b + 1) + '</td><td>' + escHtml(lineB) + '</td></tr>');
      }
    }
  }
  htmlParts.push('</table>');

  if (gen !== undefined && gen !== renderGeneration) return;

  let finalHtml = htmlParts.join('');
  if (truncated) {
    finalHtml = '<div class="bigTextBanner" style="margin:8px">' + escHtml(t().lineTruncated(visibleRows.length, limit)) +
      ' <button id="btnShowAllLines" type="button">' + escHtml(t().showAllBtn) + '</button></div>' + finalHtml;
  }
  $('out').innerHTML = finalHtml;
  const showAllBtn = $('btnShowAllLines');
  if (showAllBtn) showAllBtn.onclick = () => { $('showAllLines').dataset.active = '1'; render(); };

  totalHunks = hunkIndex + 1;
  curHunk = -1;
  updateNav();
  if (window.enableTableCopy) window.enableTableCopy('#out');
}

// ---------------------------------------------------------------------------
// Set mode: treat each side as a multiset-insensitive set of lines, show
// only-in-A / only-in-B.
// ---------------------------------------------------------------------------
function renderSetMode(textA, textB, opts, gen) {
  const rawA = splitLines(textA, opts), rawB = splitLines(textB, opts);
  const normMapA = new Map(); // normalized -> first original value
  const normMapB = new Map();
  for (const l of rawA) { const n = normLine(l, opts); if (!normMapA.has(n)) normMapA.set(n, l); }
  for (const l of rawB) { const n = normLine(l, opts); if (!normMapB.has(n)) normMapB.set(n, l); }

  const onlyA = [], onlyB = [];
  let common = 0;
  for (const [n, orig] of normMapA) { if (normMapB.has(n)) common++; else onlyA.push(orig); }
  for (const [n, orig] of normMapB) { if (!normMapA.has(n)) onlyB.push(orig); }

  if (gen !== undefined && gen !== renderGeneration) return;

  if (onlyA.length === 0 && onlyB.length === 0) {
    $('stats').innerHTML = t().setSame(normMapA.size);
  } else {
    $('stats').innerHTML = t().setDiff(normMapA.size, normMapB.size, common, onlyA.length, onlyB.length);
  }

  const renderCol = (label, cls, items) => {
    const body = items.length
      ? '<ul>' + items.map(l => '<li>' + escHtml(l) + '</li>').join('') + '</ul>'
      : '<div class="empty">' + t().onlyEmpty + '</div>';
    return '<div class="setCol ' + cls + '"><h3>' + escHtml(label) + ' (' + items.length + ')' +
      '<button class="copyBtn" type="button" data-copy="' + cls + '">' + escHtml(t().copyBtn) + '</button></h3>' + body + '</div>';
  };

  $('out').innerHTML = '<div class="setResult">' +
    renderCol(t().onlyALabel, 'onlyA', onlyA) +
    renderCol(t().onlyBLabel, 'onlyB', onlyB) +
    '</div>';

  $('out').querySelectorAll('.copyBtn').forEach(btn => {
    btn.onclick = () => {
      const items = btn.dataset.copy === 'onlyA' ? onlyA : onlyB;
      navigator.clipboard.writeText(items.join('\n')).then(() => {
        const old = btn.textContent;
        btn.textContent = t().copied;
        setTimeout(() => { btn.textContent = old; }, 1200);
      });
    };
  });

  totalHunks = 0; curHunk = -1; updateNav();
}

// ---------------------------------------------------------------------------
// Table mode: Excel cell-by-cell comparison by position.
// ---------------------------------------------------------------------------
function renderTableMode(opts, gen) {
  const sheetA = $('sheetSelectA').value, sheetB = $('sheetSelectB').value;
  const dataA = (workbookA.sheets[sheetA] || { rows: [] }).rows;
  const dataB = (workbookB.sheets[sheetB] || { rows: [] }).rows;
  const firstRowHeader = $('tableFirstRowHeader').checked;

  const headerA = firstRowHeader ? (dataA[0] || []) : [];
  const headerB = firstRowHeader ? (dataB[0] || []) : [];
  const bodyA = firstRowHeader ? dataA.slice(1) : dataA;
  const bodyB = firstRowHeader ? dataB.slice(1) : dataB;

  const rowCount = Math.max(bodyA.length, bodyB.length);
  const colCount = Math.max(headerA.length, headerB.length,
    bodyA.reduce((m, r) => Math.max(m, r.length), 0),
    bodyB.reduce((m, r) => Math.max(m, r.length), 0));

  let diffCells = 0, onlyARows = 0, onlyBRows = 0;
  const cellEq = (x, y) => opts.ignoreCase ? String(x ?? '').toLowerCase() === String(y ?? '').toLowerCase() : String(x ?? '') === String(y ?? '');

  const showAll = $('showAllLines').dataset.active === '1';
  const limit = showAll ? rowCount : Math.min(rowCount, TABLE_MODE_RENDER_LIMIT);

  const headerHtml = '<tr><th class="rownumHead">#</th>' +
    Array.from({ length: colCount }, (_, c) => '<th>' + escHtml(headerA[c] ?? headerB[c] ?? ('Col ' + (c + 1))) + '</th>').join('') + '</tr>';

  const bodyRows = [];
  for (let r = 0; r < limit; r++) {
    const rowA = bodyA[r], rowB = bodyB[r];
    if (!rowA && rowB) { onlyBRows++; bodyRows.push('<tr>' + rownumCell(r + 1) + colsOnly(rowB, colCount, 'cellOnlyB') + '</tr>'); continue; }
    if (rowA && !rowB) { onlyARows++; bodyRows.push('<tr>' + rownumCell(r + 1) + colsOnly(rowA, colCount, 'cellOnlyA') + '</tr>'); continue; }
    if (!rowA && !rowB) continue;
    let cells = '';
    for (let c = 0; c < colCount; c++) {
      const va = rowA[c], vb = rowB[c];
      if (!cellEq(va, vb)) {
        diffCells++;
        cells += '<td class="cellDiff">' + escHtml(String(va ?? '')) + ' → ' + escHtml(String(vb ?? '')) + '</td>';
      } else {
        cells += '<td>' + escHtml(String(va ?? '')) + '</td>';
      }
    }
    bodyRows.push('<tr>' + rownumCell(r + 1) + cells + '</tr>');
  }

  if (gen !== undefined && gen !== renderGeneration) return;

  if (diffCells === 0 && onlyARows === 0 && onlyBRows === 0) {
    $('stats').innerHTML = t().tableSame(rowCount, colCount);
  } else {
    $('stats').innerHTML = t().tableDiff(rowCount, colCount, diffCells, onlyARows, onlyBRows);
  }

  let html = '<div class="cmpTableWrap"><table class="cmpTable">' + headerHtml + bodyRows.join('') + '</table></div>';
  if (limit < rowCount) {
    html = '<div class="bigTextBanner" style="margin:8px">' + escHtml(t().tableTruncated(rowCount, limit)) +
      ' <button id="btnShowAllLines" type="button">' + escHtml(t().showAllBtn) + '</button></div>' + html;
  }
  $('out').innerHTML = html;
  const showAllBtn = $('btnShowAllLines');
  if (showAllBtn) showAllBtn.onclick = () => { $('showAllLines').dataset.active = '1'; render(); };

  totalHunks = 0; curHunk = -1; updateNav();
  if (window.enableTableCopy) window.enableTableCopy('#out');
}
function rownumCell(n) { return '<td class="rownum">' + n + '</td>'; }
function colsOnly(row, colCount, cls) {
  let out = '';
  for (let c = 0; c < colCount; c++) out += '<td class="' + cls + '">' + escHtml(String(row[c] ?? '')) + '</td>';
  return out;
}

// ---------------------------------------------------------------------------
// Rich mode: Word block-level diff (with original HTML formatting), or
// PDF side-by-side page images (no auto-diff, just visual comparison —
// unchanged from original behavior).
// ---------------------------------------------------------------------------
let richZoom = 1;
const RICH_ZOOM_MIN = 0.5, RICH_ZOOM_MAX = 2, RICH_ZOOM_STEP = 0.1;

function applyRichZoom() {
  document.querySelectorAll('.richDocPane, .richDocPagesPane').forEach(el => {
    el.style.setProperty('--richZoom', richZoom);
  });
  const pctEl = $('richZoomPct');
  if (pctEl) pctEl.textContent = Math.round(richZoom * 100) + '%';
}

function setupRichZoomBar() {
  if (setupRichZoomBar._done) return;
  setupRichZoomBar._done = true;
  $('btnRichZoomIn').onclick = () => { richZoom = Math.min(RICH_ZOOM_MAX, richZoom + RICH_ZOOM_STEP); applyRichZoom(); };
  $('btnRichZoomOut').onclick = () => { richZoom = Math.max(RICH_ZOOM_MIN, richZoom - RICH_ZOOM_STEP); applyRichZoom(); };
  $('btnRichZoomReset').onclick = () => { richZoom = 1; applyRichZoom(); };
}

function buildPdfPagesPane(richDoc, side, label) {
  const imgs = richDoc.pages.map(p => '<img class="richPageImg" src="' + p.dataUrl + '" data-page="' + p.pageNum + '" loading="lazy">').join('');
  let loadMoreHtml = '';
  if (richDoc.loadedPages < richDoc.totalPages) {
    const next = Math.min(richDoc.totalPages, richDoc.loadedPages + PDF_PAGE_BATCH) - richDoc.loadedPages;
    loadMoreHtml = '<div style="padding:10px;text-align:center"><button type="button" class="btnPdfLoadMore" data-side="' + side + '">' +
      escHtml(t().pdfLoadMore(next)) + '</button></div>';
  }
  return '<div class="paneWrap pane' + side + '"><div class="paneLabel">' + escHtml(label) + '</div>' +
    '<div class="richDocPane richDocPagesPane">' + imgs + loadMoreHtml + '</div></div>';
}

async function loadMorePdfPages(richDoc, side) {
  const nextCount = Math.min(richDoc.totalPages, richDoc.loadedPages + PDF_PAGE_BATCH);
  const pdf = richDoc._pdfDoc;
  for (let p = richDoc.loadedPages + 1; p <= nextCount; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    richDoc.pages.push({ dataUrl: canvas.toDataURL('image/png'), pageNum: p });
  }
  richDoc.loadedPages = nextCount;
  render();
}

function renderRichMode(opts, gen) {
  const bothBlocks = richDocA.kind === 'blocks' && richDocB.kind === 'blocks';
  let noteHtml = '';
  let paneAHtml, paneBHtml;

  if (bothBlocks) {
    const eq = opts.ignoreCase
      ? (x, y) => x.text.toLowerCase() === y.text.toLowerCase()
      : (x, y) => x.text === y.text;
    const ops = diffArrays(richDocA.blocks, richDocB.blocks, eq);
    if (gen !== undefined && gen !== renderGeneration) return;

    const delSet = new Set(), addSet = new Set();
    let delCount = 0, addCount = 0;
    for (const op of ops) {
      if (op.type === '-') { delSet.add(op.a); delCount++; }
      else if (op.type === '+') { addSet.add(op.b); addCount++; }
    }
    if (delCount === 0 && addCount === 0) {
      $('stats').innerHTML = t().richSame(richDocA.blocks.length);
    } else {
      $('stats').innerHTML = t().richDiff(richDocA.blocks.length, richDocB.blocks.length, delCount, addCount);
    }
    const buildBlocksHtml = (blocks, markSet, markClass) => blocks.map((blk, idx) =>
      '<div class="richBlock' + (markSet.has(idx) ? ' ' + markClass : '') + '">' + blk.html + '</div>').join('');
    paneAHtml = buildBlocksHtml(richDocA.blocks, delSet, 'richBlockDel');
    paneBHtml = buildBlocksHtml(richDocB.blocks, addSet, 'richBlockAdd');
    if (richDocA.truncated || richDocB.truncated) {
      const total = Math.max(richDocA.totalBlocks || richDocA.blocks.length, richDocB.totalBlocks || richDocB.blocks.length);
      noteHtml = '<div class="bigTextBanner" style="margin:8px">' + escHtml(t().richBlocksCapped(total, MAX_RICH_BLOCKS)) + '</div>';
    }
    $('out').innerHTML = noteHtml +
      '<div class="cmpTableSplit">' +
      '<div class="paneWrap paneA"><div class="paneLabel">' + escHtml(t().labelPaneA) + '</div><div class="richDocPane">' + paneAHtml + '</div></div>' +
      '<div class="splitHandle"></div>' +
      '<div class="paneWrap paneB"><div class="paneLabel">' + escHtml(t().labelPaneB) + '</div><div class="richDocPane">' + paneBHtml + '</div></div>' +
      '</div>';
  } else {
    // PDF (either or both sides) — no content diffing, just side-by-side viewing.
    $('stats').innerHTML = escHtml(t().richPdfImageNote);
    const buildPane = (doc, side, label) => {
      if (doc.kind === 'pages') return buildPdfPagesPane(doc, side, label);
      return '<div class="paneWrap pane' + side + '"><div class="paneLabel">' + escHtml(label) + '</div>' +
        '<div class="richDocPane">' + doc.blocks.map(b => '<div class="richBlock">' + b.html + '</div>').join('') + '</div></div>';
    };
    paneAHtml = buildPane(richDocA, 'A', t().labelPaneA);
    paneBHtml = buildPane(richDocB, 'B', t().labelPaneB);
    $('out').innerHTML = '<div class="cmpTableSplit">' + paneAHtml + '<div class="splitHandle"></div>' + paneBHtml + '</div>';
    $('out').querySelectorAll('.btnPdfLoadMore').forEach(btn => {
      btn.onclick = () => {
        const side = btn.dataset.side;
        loadMorePdfPages(side === 'A' ? richDocA : richDocB, side);
      };
    });
  }

  setupTableSplitDrag();
  setupTableSplitScrollSync();
  applyRichZoom();
  totalHunks = 0; curHunk = -1; updateNav();
}

function setupTableSplitScrollSync() {
  const wrapA = document.querySelector('.cmpTableSplit .paneWrap.paneA');
  const wrapB = document.querySelector('.cmpTableSplit .paneWrap.paneB');
  if (!wrapA || !wrapB) return;
  let syncing = false;
  const mirror = (src, dst) => {
    if (syncing) return;
    syncing = true;
    dst.scrollTop = src.scrollTop;
    dst.scrollLeft = src.scrollLeft;
    syncing = false;
  };
  wrapA.addEventListener('scroll', () => mirror(wrapA, wrapB));
  wrapB.addEventListener('scroll', () => mirror(wrapB, wrapA));
}

function setupTableSplitDrag() {
  const container = document.querySelector('.cmpTableSplit');
  const handle = document.querySelector('.cmpTableSplit .splitHandle');
  if (!container || !handle) return;
  let dragging = false;
  handle.addEventListener('mousedown', () => {
    dragging = true;
    handle.classList.add('dragging');
    document.body.classList.add('resizingSplit');
  });
  window.addEventListener('mousemove', (e) => {
    if (!dragging) return;
    const rect = container.getBoundingClientRect();
    let pct = ((e.clientX - rect.left) / rect.width) * 100;
    pct = Math.max(15, Math.min(85, pct));
    container.style.setProperty('--splitPct', pct + '%');
  });
  window.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.classList.remove('resizingSplit');
  });
}

// ---------------------------------------------------------------------------
// Navigation between diff hunks (line mode only)
// ---------------------------------------------------------------------------
function updateNav() {
  $('navPos').textContent = t().navPos(curHunk, totalHunks);
  $('btnPrev').disabled = totalHunks === 0;
  $('btnNext').disabled = totalHunks === 0;
}
function gotoHunk(idx) {
  const rows = document.querySelectorAll('#out tr[data-hunk]');
  document.querySelectorAll('#out tr.cur').forEach(r => r.classList.remove('cur'));
  const target = document.querySelector('#out tr[data-hunk="' + idx + '"]');
  if (target) {
    target.classList.add('cur');
    target.scrollIntoView({ block: 'center' });
  }
  curHunk = idx;
  updateNav();
}
$('btnNext').onclick = () => { if (totalHunks === 0) return; gotoHunk((curHunk + 1) % totalHunks); };
$('btnPrev').onclick = () => { if (totalHunks === 0) return; gotoHunk((curHunk - 1 + totalHunks) % totalHunks); };

// ---------------------------------------------------------------------------
// File import: Excel / Word / PDF, with size/row/page guards.
// ---------------------------------------------------------------------------
function populateSheetSelect(side, workbook) {
  const sel = $('sheetSelect' + side);
  sel.innerHTML = workbook.sheetNames.map(n => '<option value="' + escHtmlAttr(n) + '">' + escHtml(n) + '</option>').join('');
}

function updateTableModeAvailability() {
  const hasWorkbooks = !!(workbookA && workbookB);
  const hasRichDocs = !!(richDocA && richDocB);
  const anyMode = hasWorkbooks || hasRichDocs;
  $('tableModeWrap').classList.toggle('hidden', !anyMode || $('setMode').checked);
  if (!anyMode) {
    $('tableMode').checked = false;
    $('tableModeBar').classList.add('hidden');
    return;
  }
  if (hasWorkbooks) {
    $('sheetSelectRowA').classList.remove('hidden');
    $('sheetSelectRowB').classList.remove('hidden');
    $('tableFirstRowHeaderWrap').classList.remove('hidden');
    $('richZoomBar').classList.add('hidden');
    populateSheetSelect('A', workbookA);
    populateSheetSelect('B', workbookB);
    $('tableModeBar').classList.remove('hidden');
  } else {
    $('sheetSelectRowA').classList.add('hidden');
    $('sheetSelectRowB').classList.add('hidden');
    $('tableFirstRowHeaderWrap').classList.add('hidden');
    richZoom = 1;
    $('richZoomBar').classList.remove('hidden');
    $('tableModeBar').classList.remove('hidden');
    setupRichZoomBar();
  }
}
$('sheetSelectA').onchange = () => scheduleRun();
$('sheetSelectB').onchange = () => scheduleRun();
$('tableFirstRowHeader').onchange = () => scheduleRun();

async function extractFromFile(file, statusEl) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  const sizeMb = (file.size / (1024 * 1024));

  if (['xlsx', 'xls', 'csv'].includes(ext)) {
    if (file.size > MAX_EXCEL_BLOCK_BYTES) {
      throw new Error(t().excelTooBig(Math.round(sizeMb)));
    }
    if (file.size > MAX_EXCEL_WARN_BYTES && statusEl) {
      statusEl.textContent = t().excelWarnBig(Math.round(sizeMb));
    }
    const data = new Uint8Array(await file.arrayBuffer());
    const wb = XLSX.read(data, { type: 'array', cellDates: true, dense: true });
    const sheets = {};
    const textParts = [];
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      if (!sheet) { textParts.push(t().sheetUnreadable(name)); continue; }
      let rows;
      try {
        rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
      } catch {
        textParts.push(t().sheetUnreadable(name));
        continue;
      }
      const totalRows = rows.length;
      let truncatedSheet = false;
      if (totalRows > MAX_EXCEL_ROWS) {
        rows = rows.slice(0, MAX_EXCEL_ROWS);
        truncatedSheet = true;
      }
      sheets[name] = { rows, truncated: truncatedSheet, totalRows };
      textParts.push(t().sheetHeader(name));
      textParts.push(rows.map(r => r.join('\t')).join('\n'));
      if (truncatedSheet && statusEl) {
        statusEl.textContent = t().excelRowsCapped(name, totalRows, MAX_EXCEL_ROWS);
      }
    }
    return { text: textParts.join('\n'), workbook: { sheetNames: wb.SheetNames, sheets }, richDoc: null, lineCount: textParts.join('\n').split('\n').length };
  }

  if (ext === 'docx') {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const blocks = htmlToBlocks(result.value);
    const totalBlocks = blocks.length;
    let truncated = false;
    let usedBlocks = blocks;
    if (totalBlocks > MAX_RICH_BLOCKS) {
      usedBlocks = blocks.slice(0, MAX_RICH_BLOCKS);
      truncated = true;
    }
    const text = blocks.map(b => b.text).join('\n');
    return {
      text,
      workbook: null,
      richDoc: { kind: 'blocks', blocks: usedBlocks, truncated, totalBlocks },
      lineCount: text.split('\n').length,
    };
  }

  if (ext === 'pdf') {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const numPages = pdf.numPages;
    const textParts = [];
    for (let p = 1; p <= numPages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      const pageText = content.items.map(it => it.str).join(' ');
      if (numPages > 1) textParts.push(t().pageHeader(p));
      textParts.push(pageText);
    }
    const initialCount = Math.min(numPages, MAX_PDF_PAGES_INITIAL);
    const pages = [];
    for (let p = 1; p <= initialCount; p++) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      pages.push({ dataUrl: canvas.toDataURL('image/png'), pageNum: p });
    }
    if (numPages > initialCount && statusEl) {
      statusEl.textContent = t().pdfPagesCapped(numPages, initialCount);
    }
    const text = textParts.join('\n');
    return {
      text,
      workbook: null,
      richDoc: { kind: 'pages', pages, totalPages: numPages, loadedPages: initialCount, _pdfDoc: pdf },
      lineCount: text.split('\n').length,
    };
  }

  throw new Error(t().unsupportedFormat(ext));
}

// Minimal HTML-to-block splitter for mammoth output: splits on top-level
// block elements (p, h1-h6, table, li, blockquote) and strips tags to get
// a plain-text counterpart for diffing / line-mode text.
function htmlToBlocks(html) {
  const container = document.createElement('div');
  container.innerHTML = html;
  const blocks = [];
  for (const node of Array.from(container.children)) {
    blocks.push({ html: node.outerHTML, text: (node.textContent || '').trim() });
  }
  return blocks;
}

function setupImport(side) {
  const btn = $('btnImport' + side);
  const input = $('fileInput' + side);
  const statusEl = $('importStatus' + side);
  const textarea = $(side.toLowerCase());

  btn.onclick = () => input.click();
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;
    btn.disabled = true;
    statusEl.classList.remove('err');
    statusEl.textContent = t().readingFile(file.name);
    try {
      const result = await extractFromFile(file, statusEl);
      textarea.value = result.text;
      if (side === 'A') { workbookA = result.workbook; richDocA = result.richDoc; }
      else { workbookB = result.workbook; richDocB = result.richDoc; }
      if (!statusEl.textContent || statusEl.textContent === t().readingFile(file.name)) {
        statusEl.textContent = t().loadedFile(file.name, result.lineCount);
      }
      updateTableModeAvailability();
      scheduleRun();
    } catch (err) {
      statusEl.classList.add('err');
      statusEl.textContent = t().errReadingFile(err.message || String(err));
      if (side === 'A') { workbookA = null; richDocA = null; }
      else { workbookB = null; richDocB = null; }
      updateTableModeAvailability();
    } finally {
      btn.disabled = false;
      input.value = '';
    }
  };
}
setupImport('A');
setupImport('B');

// ---------------------------------------------------------------------------
// Top-level wiring
// ---------------------------------------------------------------------------
$('a').addEventListener('input', scheduleRun);
$('b').addEventListener('input', scheduleRun);
$('btnCmp').onclick = () => render();
$('btnSwap').onclick = () => {
  const va = $('a').value, vb = $('b').value;
  $('a').value = vb; $('b').value = va;
  [workbookA, workbookB] = [workbookB, workbookA];
  [richDocA, richDocB] = [richDocB, richDocA];
  const sa = $('importStatusA').textContent, sb = $('importStatusB').textContent;
  $('importStatusA').textContent = sb; $('importStatusB').textContent = sa;
  updateTableModeAvailability();
  scheduleRun();
};
$('btnClear').onclick = () => {
  $('a').value = ''; $('b').value = '';
  workbookA = workbookB = richDocA = richDocB = null;
  $('importStatusA').textContent = ''; $('importStatusB').textContent = '';
  $('showAllLines').dataset.active = '0';
  updateTableModeAvailability();
  render();
};
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); render(); }
  else if (e.altKey && e.key === 'ArrowDown') { e.preventDefault(); if (totalHunks) gotoHunk((curHunk + 1) % totalHunks); }
  else if (e.altKey && e.key === 'ArrowUp') { e.preventDefault(); if (totalHunks) gotoHunk((curHunk - 1 + totalHunks) % totalHunks); }
});

window.onToolsLangChange = () => render();
window.applyI18n && window.applyI18n();
updateSetModeUi();
updateTableModeAvailability();
render();

})();
