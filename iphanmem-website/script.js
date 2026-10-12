/* Iphanmem: menu di động, trang tải về (đọc releases.json), nút chép. Không phụ thuộc thư viện. */
(function () {
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  var btn = $('.menu-btn'), nav = $('.site-h nav');
  if (btn && nav) btn.addEventListener('click', function () {
    var open = nav.classList.toggle('open');
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
  });

  /* Trang tải về: đọc releases.json. Chưa có bộ cài thật thì hiện "sắp phát hành" và dẫn sang Zalo. */
  var rels = $$('.rel');
  if (rels.length && window.fetch) {
    fetch('releases.json').then(function (r) { return r.json(); }).then(function (data) {
      var anySoon = false;
      rels.forEach(function (card) {
        var r = (data.products || {})[card.getAttribute('data-id')] || { available: false };
        var set = function (k, v) { var el = $('[data-f="' + k + '"]', card); if (el && v) el.textContent = v; };
        var a = $('[data-f="link"]', card), det = $('details', card);
        if (!r.available || !r.file) {
          anySoon = true;
          set('sha', 'Sẽ có khi phát hành');
          if (a) a.textContent = 'Nhắn Zalo nhận bản dùng thử';
          if (det) det.hidden = true;
          return;
        }
        ['version', 'date', 'size', 'channel'].forEach(function (k) { set(k, r[k]); });
        set('sha', r.sha256);
        var ul = $('[data-f="notes"]', card);
        if (ul) (r.notes || []).forEach(function (n) { var li = document.createElement('li'); li.textContent = n; ul.appendChild(li); });
        if (!(r.notes || []).length && det) det.hidden = true;
        if (a) { a.href = data.baseUrl.replace(/\/$/, '') + '/' + r.file; a.removeAttribute('target'); a.setAttribute('download', ''); a.textContent = 'Tải về'; }
      });
      var note = $('#soon-note'); if (note) note.hidden = !anySoon;
    }).catch(function () { /* giữ nội dung mặc định */ });
  }
  $$('[data-copy],[data-copy-text]').forEach(function (b) {
    b.addEventListener('click', function () {
      var code = b.parentNode.querySelector('code'), txt = b.getAttribute('data-copy-text') || (code ? code.textContent : '');
      var label = b.textContent;
      var done = function () { b.textContent = 'Đã chép'; setTimeout(function () { b.textContent = label; }, 1600); };
      try { navigator.clipboard.writeText(txt).then(done, function () { selectText(code); }); } catch (e) { selectText(code); }
    });
  });
  function selectText(el) {
    if (!el) return;
    var r = document.createRange(); r.selectNodeContents(el);
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
  }

})();
