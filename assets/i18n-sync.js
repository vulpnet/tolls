// Đồng bộ ngôn ngữ giao diện (vi/en) giữa menu chính và mọi tool chạy trong iframe.
// Menu chính lưu localStorage 'toolsLang' rồi postMessage xuống từng iframe;
// mỗi tool tự đọc localStorage lúc load (trước khi user đổi) + lắng nghe message (đổi ngôn ngữ realtime).
//
// Cách dùng trong 1 tool:
//   1. Định nghĩa DICT = { vi: {key: '...'}, en: {key: '...'} }
//   2. Gọi applyI18n(DICT) sau khi DOM sẵn sàng để render lần đầu theo ngôn ngữ đã lưu
//   3. Đăng ký callback qua window.onToolsLangChange = (lang) => { ...render lại UI... }
//      applyI18n() tự gọi callback này mỗi khi ngôn ngữ đổi (kể cả lần đầu tiên)
(function(){
  const LANG_KEY = 'toolsLang';

  function getSavedLang(){
    try{ return localStorage.getItem(LANG_KEY) || 'vi'; }catch(e){ return 'vi'; }
  }

  window.__toolsLang = getSavedLang();

  window.addEventListener('message', e => {
    if (!e.data || e.data.type !== 'toolsLang') return;
    window.__toolsLang = e.data.value || 'vi';
    if (typeof window.onToolsLangChange === 'function') window.onToolsLangChange(window.__toolsLang);
  });

  // gọi ngay khi tool tự khai báo applyI18n — cho phép tool set callback trước rồi mới gọi render lần đầu
  window.applyI18n = function(){
    if (typeof window.onToolsLangChange === 'function') window.onToolsLangChange(window.__toolsLang);
  };
})();
