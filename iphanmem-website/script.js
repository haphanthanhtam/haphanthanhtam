/* Iphanmem: menu di động, trang tải xuống (đọc releases.json), form liên hệ. Không phụ thuộc thư viện. */
(function () {
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  var btn = $('.menu-btn'), nav = $('.site-h nav');
  if (btn && nav) btn.addEventListener('click', function () {
    var open = nav.classList.toggle('open');
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
  });

  /* Trang tải xuống: điền phiên bản, mã SHA-256 và liên kết từ releases.json */
  var rels = $$('.rel');
  if (rels.length && window.fetch) {
    fetch('releases.json').then(function (r) { return r.json(); }).then(function (data) {
      var note = $('#sample-note'); if (note) note.hidden = !data.sample;
      rels.forEach(function (card) {
        var r = (data.products || {})[card.getAttribute('data-id')]; if (!r) return;
        var set = function (k, v) { var el = $('[data-f="' + k + '"]', card); if (el && v) el.textContent = v; };
        ['version', 'date', 'size', 'channel', 'sha'].forEach(function (k) { set(k, r[k === 'sha' ? 'sha256' : k]); });
        var ul = $('[data-f="notes"]', card);
        if (ul) ul.innerHTML = (r.notes || []).map(function (n) { var li = document.createElement('li'); li.textContent = n; return li.outerHTML; }).join('');
        var a = $('[data-f="link"]', card);
        if (a && r.file && !data.sample) { a.href = data.baseUrl.replace(/\/$/, '') + '/' + r.file; a.setAttribute('download', ''); }
        else if (a && data.sample) { a.textContent = 'Tải về (bản mẫu)'; a.setAttribute('aria-disabled', 'true'); }
      });
    }).catch(function () { /* giữ nội dung mặc định */ });
  }
  $$('[data-copy]').forEach(function (b) {
    b.addEventListener('click', function () {
      var code = b.parentNode.querySelector('code'), txt = code ? code.textContent : '';
      var done = function () { b.textContent = 'Đã chép'; setTimeout(function () { b.textContent = 'Chép'; }, 1600); };
      try { navigator.clipboard.writeText(txt).then(done, function () { selectText(code); }); } catch (e) { selectText(code); }
    });
  });
  function selectText(el) {
    if (!el) return;
    var r = document.createRange(); r.selectNodeContents(el);
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
  }

  /* Liên hệ: chọn sẵn quan tâm từ ?i= ; gửi bằng ứng dụng email của khách */
  var form = $('#contact-form');
  if (form) {
    try {
      var q = new URLSearchParams(location.search), i = q.get('i'), p = q.get('p');
      var sel = $('#f-int');
      if (i && sel) $$('option', sel).forEach(function (o) { if (o.value.toLowerCase().indexOf(i) === 0) sel.value = o.value; });
      if (p && $('#f-msg')) $('#f-msg').value = 'Tôi cần: ' + p + '.';
    } catch (e) { /* bỏ qua */ }
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var v = function (id) { return ($('#' + id) || {}).value || ''; };
      var body = 'Họ tên: ' + v('f-name') + '\nEmail: ' + v('f-mail') + '\nĐiện thoại/Zalo: ' + v('f-phone') +
        '\nCông ty: ' + v('f-co') + '\nQuan tâm: ' + v('f-int') + '\n\n' + v('f-msg');
      var out = $('#form-out'); if (out) out.hidden = false;
      window.location.href = 'mailto:' + form.getAttribute('data-to') + '?subject=' + encodeURIComponent('[Iphanmem] ' + v('f-int')) + '&body=' + encodeURIComponent(body);
    });
  }
})();
