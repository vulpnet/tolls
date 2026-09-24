// Đồng bộ màu nền tùy chỉnh (--bg) giữa menu chính và mọi tool chạy trong iframe.
// Menu chính lưu localStorage 'toolsCustomBg' rồi postMessage xuống từng iframe;
// mỗi tool tự đọc localStorage lúc load (page đầu, trước khi user đổi màu) + lắng nghe message (đổi màu realtime).
(function(){
  function applyBg(hex){
    if (hex) document.documentElement.style.setProperty('--bg', hex);
    else document.documentElement.style.removeProperty('--bg'); // fix: rỗng = trả về màu mặc định theo light/dark system
  }

  try{
    const saved = localStorage.getItem('toolsCustomBg');
    if (saved) applyBg(saved);
  }catch(e){ /* localStorage có thể bị chặn (chế độ ẩn danh nghiêm ngặt) — bỏ qua, không vỡ trang */ }

  window.addEventListener('message', e => {
    if (!e.data || e.data.type !== 'toolsCustomBg') return;
    applyBg(e.data.value || null);
  });
})();
