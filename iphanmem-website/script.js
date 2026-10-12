/* iPhanMem: menu di động, tab và chu kỳ trên bảng giá, form liên hệ. Không phụ thuộc thư viện. */
(function () {
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  var btn = $('.menu-btn'), nav = $('.site-h nav');
  if (btn && nav) btn.addEventListener('click', function () {
    var open = nav.classList.toggle('open');
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
  });

  /* Bảng giá: tab sản phẩm + số license (từ 5 giảm 10%, từ 10 giảm 20%) */
  var tabs = $$('[data-tab]'), panels = $$('.panel');
  function fmt(n) { return Math.round(n).toLocaleString('vi-VN') + ' đ'; }
  function paint() {
    var q = Math.max(1, Math.min(999, parseInt(($('#qty') || {}).value, 10) || 1));
    var d = q >= 10 ? 0.2 : q >= 5 ? 0.1 : 0;
    $$('[data-p]').forEach(function (el) {
      el.textContent = fmt(parseFloat(el.getAttribute('data-p')) * q * (1 - d));
    });
    var note = $('#qty-note');
    if (note) note.textContent = (d ? 'Đã áp dụng giảm ' + d * 100 + '% cho ' + q + ' license. ' : '') +
      'Mua từ 5 license giảm 10%, từ 10 license giảm 20% (đề xuất).';
  }
  tabs.forEach(function (t) {
    t.addEventListener('click', function () {
      var id = t.getAttribute('data-tab');
      tabs.forEach(function (x) { x.classList.toggle('on', x === t); });
      panels.forEach(function (p) { p.classList.toggle('on', p.id === 'tab-' + id); });
      var c = $('#qty-ctl'), n = $('#qty-note'); if (c) c.hidden = n.hidden = (id === 'erp');
    });
  });
  var qty = $('#qty'); if (qty) qty.addEventListener('input', paint);

  /* Liên hệ: chọn sẵn quan tâm từ ?i= ; gửi bằng ứng dụng email của khách */
  var form = $('#contact-form');
  if (form) {
    try {
      var q = new URLSearchParams(location.search), i = q.get('i'), p = q.get('p');
      var sel = $('#f-int');
      if (i && sel) $$('option', sel).forEach(function (o) { if (o.value.toLowerCase().indexOf(i) === 0) sel.value = o.value; });
      if (p && $('#f-msg')) $('#f-msg').value = 'Tôi quan tâm gói ' + p + '.';
    } catch (e) { /* bỏ qua */ }
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var v = function (id) { return ($('#' + id) || {}).value || ''; };
      var body = 'Họ tên: ' + v('f-name') + '\nEmail: ' + v('f-mail') + '\nĐiện thoại/Zalo: ' + v('f-phone') +
        '\nCông ty: ' + v('f-co') + '\nQuan tâm: ' + v('f-int') + '\n\n' + v('f-msg');
      var to = form.getAttribute('data-to');
      var out = $('#form-out'); if (out) { out.hidden = false; }
      window.location.href = 'mailto:' + to + '?subject=' + encodeURIComponent('[iPhanMem] ' + v('f-int')) + '&body=' + encodeURIComponent(body);
    });
  }
})();
