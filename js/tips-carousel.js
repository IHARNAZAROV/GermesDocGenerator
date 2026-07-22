'use strict';
/* ── Tips Carousel ─────────────────────────────────────── */
(function () {
  var track  = document.getElementById('tips-track');
  var dots   = document.querySelectorAll('.tips-dot');
  var total  = dots.length;
  var current = 0;
  var timer  = null;

  if (!track || !total) return;

  function goTo(idx) {
    current = ((idx % total) + total) % total;
    // Fade-transition при смене слайда
    track.style.opacity = '0';
    track.style.transition = 'opacity 160ms ease';
    setTimeout(function () {
      track.style.transform = 'translateX(-' + (current * 100) + '%)';
      track.style.opacity = '1';
    }, 160);
    dots.forEach(function (d, i) {
      d.classList.toggle('tips-dot--active', i === current);
    });
  }

  function startAuto() {
    clearInterval(timer);
    timer = setInterval(function () { goTo(current + 1); }, 6000);
  }

  dots.forEach(function (d) {
    d.addEventListener('click', function () {
      goTo(+d.dataset.tip);
      startAuto();
    });
  });

  startAuto();
})();
