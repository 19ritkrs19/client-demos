// LaunchLocal - shared JS (nav toggle + scroll reveal + year)
document.addEventListener('DOMContentLoaded', function () {
  // mobile nav toggle
  var toggle = document.querySelector('.nav-toggle');
  var links = document.querySelector('.nav-links');
  if (toggle && links) {
    toggle.addEventListener('click', function () {
      links.classList.toggle('open');
      toggle.textContent = links.classList.contains('open') ? '✕' : '☰';
    });
    links.querySelectorAll('a').forEach(function (a) {
      a.addEventListener('click', function () { links.classList.remove('open'); toggle.textContent = '☰'; });
    });
  }

  // scroll reveal
  var io = new IntersectionObserver(function (entries) {
    entries.forEach(function (e) {
      if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
    });
  }, { threshold: 0.12 });
  document.querySelectorAll('.reveal').forEach(function (el, i) {
    el.style.transitionDelay = (i % 6 * 60) + 'ms';
    io.observe(el);
  });

  // footer year
  var yr = document.getElementById('yr');
  if (yr) yr.textContent = new Date().getFullYear();

  // Lead form: email (Formspree) + WhatsApp, no page break
  var form = document.getElementById('leadForm');
  if (form) {
    var WA_NUMBER = '918130283848';
    var msg = document.getElementById('formMsg');
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var btn = form.querySelector('button[type=submit]');
      var data = {
        name: form.name.value.trim(),
        phone: form.phone.value.trim(),
        business: form.business.value.trim(),
        need: form.need.value,
        message: form.message.value.trim()
      };
      if (!data.name || !data.phone) { return; }

      btn.disabled = true; btn.textContent = 'Sending...';

      // 1) Email via Formspree (AJAX - page break nahi)
      fetch(form.dataset.formspree, {
        method: 'POST',
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      }).catch(function(){/* email fail ho to bhi WhatsApp chalega */})
      .finally(function () {
        // 2) WhatsApp pre-filled message
        var text = 'Hi LaunchLocal! My details:%0A' +
          'Name: ' + encodeURIComponent(data.name) + '%0A' +
          'Phone: ' + encodeURIComponent(data.phone) + '%0A' +
          'Business: ' + encodeURIComponent(data.business || '-') + '%0A' +
          'Need: ' + encodeURIComponent(data.need) + '%0A' +
          'Message: ' + encodeURIComponent(data.message || '-');
        var waUrl = 'https://wa.me/' + WA_NUMBER + '?text=' + text;

        // 3) success message + open WhatsApp
        if (msg) {
          msg.style.display = 'block';
          msg.style.color = '#00d4b3';
          msg.innerHTML = '✅ Details sent! Opening WhatsApp... if it does not open, ' +
            '<a href="' + waUrl + '" style="color:#4f7cff">click here</a>.';
        }
        form.reset();
        btn.disabled = false; btn.textContent = 'Send & Get Free Demo';
        window.open(waUrl, '_blank');
      });
    });
  }
});
