/* Scansione scontrino: la foto viene ridotta e migliorata sul telefono, letta dal Raspberry con Tesseract
   e usata per precompilare il modulo. Non salva nulla da sola. */
(function () {
  var btn = document.getElementById('scanb');
  if (!btn) return;
  var cam = document.getElementById('scanf'), gal = document.getElementById('scang'),
      st = document.getElementById('scans'), gl = document.getElementById('scangl');

  function field(n) { return document.querySelector('[name=' + n + ']'); }
  function say(msg, bad, text) {
    st.hidden = false; st.style.color = bad ? '#c62828' : ''; st.textContent = msg;
    if (text) {
      var d = document.createElement('details'), s = document.createElement('summary'), p = document.createElement('pre');
      s.textContent = 'Testo letto'; p.textContent = text; p.style.cssText = 'white-space:pre-wrap;font-size:12px;margin:6px 0 0';
      d.appendChild(s); d.appendChild(p); st.appendChild(d);
    }
  }
  function set(n, v) {
    var el = field(n); if (!el || v === '' || v == null) return false;
    el.value = v;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  /* Riduce a 2000 px, scala di grigi e contrasto: aiuta l'OCR e alleggerisce l'invio */
  function prepare(file, done) {
    var img = new Image(), url = URL.createObjectURL(file);
    img.onload = function () {
      var k = Math.min(1, 2000 / Math.max(img.naturalWidth, img.naturalHeight)),
          c = document.createElement('canvas'), w = c.width = Math.round(img.naturalWidth * k), h = c.height = Math.round(img.naturalHeight * k),
          x = c.getContext('2d');
      x.drawImage(img, 0, 0, w, h); URL.revokeObjectURL(url);
      try {
        var im = x.getImageData(0, 0, w, h), d = im.data, hist = new Array(256).fill(0), i, g;
        for (i = 0; i < d.length; i += 4) { g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0; d[i] = g; hist[g]++; }
        var n = w * h, lo = 0, hi = 255, acc = 0;
        for (i = 0; i < 256; i++) { acc += hist[i]; if (acc > n * 0.02) { lo = i; break; } }
        acc = 0;
        for (i = 255; i >= 0; i--) { acc += hist[i]; if (acc > n * 0.2) { hi = i; break; } }   /* la carta è la parte chiara */
        var sc = 255 / Math.max(hi - lo, 30);
        for (i = 0; i < d.length; i += 4) { g = Math.max(0, Math.min(255, ((d[i] - lo) * sc) | 0)); d[i] = d[i + 1] = d[i + 2] = g; }
        x.putImageData(im, 0, 0);
      } catch (e) { /* si invia l'immagine ridotta così com'è */ }
      c.toBlob(function (b) { done(b || file); }, 'image/jpeg', 0.9);
    };
    img.onerror = function () { URL.revokeObjectURL(url); done(file); };   /* Formato non decodificabile: si invia l'originale */
    img.src = url;
  }

  function run(file) {
    if (!file) return;
    btn.disabled = true; say('Leggo lo scontrino… può servire qualche secondo.');
    prepare(file, function (blob) {
      var fd = new FormData(); fd.append('image', blob, 'scontrino.jpg');
      fetch('scan', { method: 'POST', body: fd, credentials: 'same-origin' })
        .then(function (r) { return r.json().catch(function () { return { error: 'Risposta non valida dal server (timeout?)' }; }); })
        .then(function (r) {
          btn.disabled = false;
          if (r.error) { say(r.error, true); return; }
          var got = [];
          if (set('amount', r.amount)) got.push(r.amount + ' €');
          if (set('voucher', r.voucher)) got.push('buoni ' + r.voucher + ' €');
          if (set('description', r.store)) got.unshift(r.store);
          if (r.date && set('date', r.date)) got.push(r.date.split('-').reverse().join('/'));
          var cat = field('category_id');
          if (r.category_id && cat && Array.prototype.some.call(cat.options, function (o) { return o.value == r.category_id; })) {
            cat.value = r.category_id; cat.dispatchEvent(new Event('change', { bubbles: true }));
          }
          if (got.length) { var fm = document.querySelector('form.card'); if (fm) fm.scrollIntoView({ behavior: 'smooth', block: 'start' }); }   /* Il pulsante sta sotto il modulo: si torna in alto a controllarlo */
          say((got.length ? 'Letto: ' + got.join(' · ') + '. Controlla i campi e salva.' : 'Non sono riuscito a leggere nulla di utile: riprova con una foto più nitida.') +
              (r.note ? ' ' + r.note + '.' : ''), !got.length || !r.amount, r.text);
        })
        .catch(function () { btn.disabled = false; say('Non riesco a contattare il server.', true); });
    });
  }

  btn.addEventListener('click', function () { cam.click(); });
  gl.addEventListener('click', function (e) { e.preventDefault(); gal.click(); });
  [cam, gal].forEach(function (inp) {
    inp.addEventListener('change', function () { run(inp.files[0]); inp.value = ''; });
  });
})();
