(function() {
  'use strict';

  // ===== 1. Hover prefetch: start loading pages on hover so clicks feel instant =====
  var prefetched = {};
  document.addEventListener('mouseover', function(e) {
    var link = e.target.closest('a');
    if (!link) return;
    var href = link.getAttribute('href');
    if (!href || href.startsWith('#') || href.startsWith('http') || href.startsWith('mailto') || link.target === '_blank') return;
    if (prefetched[href]) return;
    prefetched[href] = true;
    var l = document.createElement('link');
    l.rel = 'prefetch';
    l.href = href;
    document.head.appendChild(l);
  });

  // ===== 2. Typing effect =====
  var typingEl = document.querySelector('.hero-typing');
  if (typingEl) {
    var fullText = typingEl.textContent;
    typingEl.textContent = '';
    typingEl.style.visibility = 'visible';
    var charIndex = 0;
    function typeNext() {
      if (charIndex < fullText.length) {
        typingEl.textContent += fullText.charAt(charIndex);
        charIndex++;
        setTimeout(typeNext, 20);
      } else {
        typingEl.classList.add('typing-done');
      }
    }
    setTimeout(typeNext, 400);
  }

  // ===== 3. Animated number counters =====
  var counters = document.querySelectorAll('.scholar-metric__value');
  if (counters.length) {
    var counterObs = new IntersectionObserver(function(entries) {
      entries.forEach(function(entry) {
        if (!entry.isIntersecting) return;
        var el = entry.target;
        var target = parseInt(el.textContent, 10);
        if (isNaN(target)) return;
        counterObs.unobserve(el);
        var start = 0;
        var duration = 800;
        var startTime = null;
        function step(ts) {
          if (!startTime) startTime = ts;
          var progress = Math.min((ts - startTime) / duration, 1);
          var eased = 1 - Math.pow(1 - progress, 3);
          el.textContent = Math.floor(eased * target);
          if (progress < 1) requestAnimationFrame(step);
          else el.textContent = target;
        }
        requestAnimationFrame(step);
      });
    }, { threshold: 0.5 });
    counters.forEach(function(c) { counterObs.observe(c); });
  }

  // ===== 4. Particle constellation background =====
  var canvas = document.getElementById('hero-particles');
  if (canvas) {
    var ctx = canvas.getContext('2d');
    var particles = [];
    var PARTICLE_COUNT = 50;
    var CONNECT_DIST = 120;
    var animId;

    function resize() {
      var rect = canvas.parentElement.getBoundingClientRect();
      canvas.width = rect.width;
      canvas.height = rect.height;
    }

    function initParticles() {
      particles = [];
      for (var i = 0; i < PARTICLE_COUNT; i++) {
        particles.push({
          x: Math.random() * canvas.width,
          y: Math.random() * canvas.height,
          vx: (Math.random() - 0.5) * 0.4,
          vy: (Math.random() - 0.5) * 0.4,
          r: Math.random() * 2 + 1
        });
      }
    }

    function draw() {
      ctx.clearRect(0, 0, canvas.width, canvas.height);

      for (var i = 0; i < particles.length; i++) {
        var p = particles[i];
        p.x += p.vx;
        p.y += p.vy;
        if (p.x < 0) p.x = canvas.width;
        if (p.x > canvas.width) p.x = 0;
        if (p.y < 0) p.y = canvas.height;
        if (p.y > canvas.height) p.y = 0;

        ctx.beginPath();
        ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(16, 112, 112, 0.35)';
        ctx.fill();

        for (var j = i + 1; j < particles.length; j++) {
          var q = particles[j];
          var dx = p.x - q.x;
          var dy = p.y - q.y;
          var dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < CONNECT_DIST) {
            ctx.beginPath();
            ctx.moveTo(p.x, p.y);
            ctx.lineTo(q.x, q.y);
            ctx.strokeStyle = 'rgba(16, 112, 112, ' + (0.12 * (1 - dist / CONNECT_DIST)) + ')';
            ctx.lineWidth = 0.8;
            ctx.stroke();
          }
        }
      }
      animId = requestAnimationFrame(draw);
    }

    resize();
    initParticles();
    draw();
    window.addEventListener('resize', function() { resize(); initParticles(); });
  }

  // ===== 5. Parallax on scroll =====
  var parallaxEls = document.querySelectorAll('[data-parallax]');
  if (parallaxEls.length) {
    window.addEventListener('scroll', function() {
      var scrollY = window.scrollY;
      parallaxEls.forEach(function(el) {
        var speed = parseFloat(el.getAttribute('data-parallax')) || 0.15;
        var rect = el.getBoundingClientRect();
        var offset = (rect.top + scrollY - window.innerHeight / 2) * speed;
        el.style.transform = 'translateY(' + (-offset) + 'px)';
      });
    }, { passive: true });
  }

  // ===== 6. Research card expand (static panel below grid) =====
  var rdPanel = document.getElementById('rd-detail-panel');
  if (rdPanel) {
    var rdCards = Array.from(document.querySelectorAll('.rd-card'));
    var rdCloseBtn = rdPanel.querySelector('.rd-detail__close');
    var rdHeaderH3 = rdPanel.querySelector('.rd-detail__header h3');
    var rdBody = rdPanel.querySelector('.rd-detail__body');

    function closeRd() {
      rdPanel.style.display = 'none';
      rdCards.forEach(function(c) { c.classList.remove('open'); });
    }

    rdCloseBtn.addEventListener('click', function(ev) {
      ev.stopPropagation();
      closeRd();
    });

    rdCards.forEach(function(card) {
      card.addEventListener('click', function(e) {
        if (e.target.closest('a') || e.target.closest('button')) return;

        var wasOpen = card.classList.contains('open');
        closeRd();
        if (wasOpen) return;

        var expanded = card.querySelector('.rd-card__expanded');
        var title = card.querySelector('.rd-card__front h3');
        if (!expanded || !title) return;

        card.classList.add('open');
        rdHeaderH3.textContent = title.textContent;
        rdBody.innerHTML = expanded.innerHTML;
        rdPanel.style.display = 'block';

        setTimeout(function() {
          rdPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 50);
      });
    });

    document.addEventListener('click', function(e) {
      if (!e.target.closest('.rd-card') && !e.target.closest('#rd-detail-panel')) {
        closeRd();
      }
    });
  }

})();
