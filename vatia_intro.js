/* vatia_intro.js — Splash VATIA · 4 estilos (port de Python/Tkinter) */
(function () {
  const CFG = {
    bg:     '#1E3A5F',
    fg:     '#FFFFFF',
    accent: '#0E7490',
    slack:  '#047857',
    pv:     '#6D28D9',
    pq:     '#C2410C',
    style:  'nexus',   // 'nexus' | 'pulse' | 'wave' | 'tower'
    duration: 1200,
    fade:     300
  };

  /* ---------- utilidades de color ---------- */
  const hex2rgb = h => {
    h = h.replace('#', '');
    return [parseInt(h.substr(0, 2), 16),
            parseInt(h.substr(2, 2), 16),
            parseInt(h.substr(4, 2), 16)];
  };
  const rgb2hex = (r, g, b) =>
    '#' + [r, g, b].map(v => Math.round(v).toString(16).padStart(2, '0')).join('');

  function interpColor(a, b, t) {
    const A = hex2rgb(a), B = hex2rgb(b);
    return rgb2hex(
      A[0] + (B[0] - A[0]) * t,
      A[1] + (B[1] - A[1]) * t,
      A[2] + (B[2] - A[2]) * t
    );
  }

  /* ---------- helpers de dibujo ---------- */
  function strokePolyline(ctx, pts, color, width, smooth) {
    ctx.beginPath();
    if (smooth && pts.length > 4) {
      // Catmull-Rom simplificado → quadraticCurveTo
      ctx.moveTo(pts[0], pts[1]);
      for (let i = 2; i < pts.length - 2; i += 2) {
        const xc = (pts[i] + pts[i + 2]) / 2;
        const yc = (pts[i + 1] + pts[i + 3]) / 2;
        ctx.quadraticCurveTo(pts[i], pts[i + 1], xc, yc);
      }
      ctx.lineTo(pts[pts.length - 2], pts[pts.length - 1]);
    } else {
      ctx.moveTo(pts[0], pts[1]);
      for (let i = 2; i < pts.length; i += 2) ctx.lineTo(pts[i], pts[i + 1]);
    }
    ctx.strokeStyle = color;
    ctx.lineWidth   = width;
    ctx.lineCap     = 'round';
    ctx.lineJoin    = 'round';
    ctx.stroke();
  }

  /* ---------- los 4 estilos ---------- */
  function drawLogo(ctx, cx, cy, size, style, phase) {
    const r       = size / 2;
    const lwMain  = Math.max(2, size / 14);
    const lwThin  = Math.max(1, size / 26);
    const { bg, fg, accent, slack, pv, pq } = CFG;
    const tipos   = [slack, pv, pq, pq, pv, slack];
    const nodoR   = size * 0.06;

    /* ================= NEXUS ================= */
    if (style === 'nexus') {
      ctx.beginPath();
      for (let i = 0; i < 6; i++) {
        const ang = (2 * Math.PI * i) / 6 - Math.PI / 2;
        const x = cx + r * Math.cos(ang);
        const y = cy + r * Math.sin(ang);
        i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.fillStyle = bg;
      ctx.fill();

      for (let i = 0; i < 6; i++) {
        const ang = (2 * Math.PI * i) / 6 - Math.PI / 2;
        const nx  = cx + r * 0.88 * Math.cos(ang);
        const ny  = cy + r * 0.88 * Math.sin(ang);

        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.lineTo(nx, ny);
        ctx.strokeStyle = accent;
        ctx.lineWidth   = lwThin;
        ctx.stroke();

        const pulse = 0.5 + 0.5 * Math.abs(Math.sin((phase + i * 0.15) * 2 * Math.PI));
        const haloR = nodoR * (1.4 + 0.6 * pulse);
        ctx.beginPath();
        ctx.arc(nx, ny, haloR, 0, 7);
        ctx.fillStyle = accent;
        ctx.fill();

        ctx.beginPath();
        ctx.arc(nx, ny, nodoR, 0, 7);
        ctx.fillStyle   = tipos[i];
        ctx.fill();
        ctx.strokeStyle = fg;
        ctx.lineWidth   = Math.max(1, size / 40);
        ctx.stroke();
      }

      const vH = r * 1.05, vW = r * 0.55;
      ctx.beginPath();
      ctx.moveTo(cx - vW, cy - vH / 2);
      ctx.lineTo(cx,      cy + vH / 2);
      ctx.lineTo(cx + vW, cy - vH / 2);
      ctx.strokeStyle = fg;
      ctx.lineWidth   = lwMain + 1;
      ctx.lineCap     = 'round';
      ctx.lineJoin    = 'round';
      ctx.stroke();
    }

    /* ================= PULSE ================= */
    else if (style === 'pulse') {
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, 7);
      ctx.fillStyle   = bg;
      ctx.fill();
      ctx.strokeStyle = accent;
      ctx.lineWidth   = lwMain;
      ctx.stroke();

      for (let k = 0; k < 3; k++) {
        const t = (phase + k / 3.0) % 1.0;
        const ringR = t * r * 0.85;
        const col   = interpColor(bg, accent, 1.0 - t);
        ctx.beginPath();
        ctx.arc(cx, cy, ringR, 0, 7);
        ctx.strokeStyle = col;
        ctx.lineWidth   = 2;
        ctx.stroke();
      }

      ctx.beginPath();
      ctx.arc(cx, cy, r * 0.35, 0, 7);
      ctx.fillStyle = bg;
      ctx.fill();

      const vH = r * 0.75, vW = r * 0.42;
      ctx.beginPath();
      ctx.moveTo(cx - vW, cy - vH / 2);
      ctx.lineTo(cx,      cy + vH / 2);
      ctx.lineTo(cx + vW, cy - vH / 2);
      ctx.strokeStyle = fg;
      ctx.lineWidth   = lwMain;
      ctx.lineCap     = 'round';
      ctx.lineJoin    = 'round';
      ctx.stroke();
    }

    /* ================= WAVE ================= */
    else if (style === 'wave') {
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, 7);
      ctx.fillStyle   = bg;
      ctx.fill();
      ctx.strokeStyle = accent;
      ctx.lineWidth   = lwThin * 2;
      ctx.stroke();

      const baseY = cy + r * 0.35;

      // línea punteada
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(cx - r * 0.8, baseY);
      ctx.lineTo(cx + r * 0.8, baseY);
      ctx.strokeStyle = accent;
      ctx.lineWidth   = lwThin;
      ctx.stroke();
      ctx.setLineDash([]);

      // onda principal
      const n = 60, pts = [];
      for (let i = 0; i <= n; i++) {
        const t   = i / n;
        const x   = cx - r * 0.82 + r * 1.64 * t;
        const amp = r * 0.55 * (1 - Math.abs(t - 0.5) * 2) + r * 0.15;
        const vShape = Math.abs(t - 0.5) * 2 * r * 0.35;
        const y = baseY - vShape - amp * 0.3 * Math.abs(Math.sin(t * Math.PI));
        pts.push(x, y);
      }
      strokePolyline(ctx, pts, fg, lwMain, true);

      // punto que recorre la onda
      const tP   = phase;
      const xP   = cx - r * 0.82 + r * 1.64 * tP;
      const vP   = Math.abs(tP - 0.5) * 2 * r * 0.35;
      const ampP = r * 0.55 * (1 - Math.abs(tP - 0.5) * 2) + r * 0.15;
      const yP   = baseY - vP - ampP * 0.3 * Math.abs(Math.sin(tP * Math.PI));

      ctx.beginPath();
      ctx.arc(xP, yP, size * 0.06, 0, 7);
      ctx.fillStyle   = accent;
      ctx.fill();
      ctx.strokeStyle = fg;
      ctx.lineWidth   = Math.max(1, size / 30);
      ctx.stroke();
    }

    /* ================= TOWER ================= */
    else if (style === 'tower') {
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, 7);
      ctx.fillStyle   = bg;
      ctx.fill();
      ctx.strokeStyle = accent;
      ctx.lineWidth   = lwMain;
      ctx.stroke();

      const topY = cy - r * 0.55, botY = cy + r * 0.65;
      const topW = r * 0.55,      botW = r * 0.18;

      // torres laterales
      ctx.strokeStyle = fg;
      ctx.lineCap = 'round';
      ctx.lineWidth = lwMain;
      ctx.beginPath(); ctx.moveTo(cx - topW, topY); ctx.lineTo(cx - botW, botY); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(cx + topW, topY); ctx.lineTo(cx + botW, botY); ctx.stroke();

      // travesaños horizontales
      for (const t of [0.15, 0.4, 0.65, 0.88]) {
        const yT = topY + (botY - topY) * t;
        const wT = topW + (botW - topW) * t;
        ctx.beginPath();
        ctx.moveTo(cx - wT, yT);
        ctx.lineTo(cx + wT, yT);
        ctx.strokeStyle = fg;
        ctx.lineWidth   = lwThin;
        ctx.stroke();
      }

      // cruces
      for (let i = 0; i < 3; i++) {
        const t1 = 0.15 + i * 0.25, t2 = t1 + 0.25;
        const y1 = topY + (botY - topY) * t1;
        const y2 = topY + (botY - topY) * t2;
        const w1 = topW + (botW - topW) * t1;
        const w2 = topW + (botW - topW) * t2;
        ctx.strokeStyle = accent;
        ctx.lineWidth   = lwThin;
        ctx.beginPath(); ctx.moveTo(cx - w1, y1); ctx.lineTo(cx + w2, y2); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(cx + w1, y1); ctx.lineTo(cx - w2, y2); ctx.stroke();
      }

      // barra superior
      ctx.strokeStyle = fg;
      ctx.lineWidth   = lwMain;
      ctx.beginPath();
      ctx.moveTo(cx - topW * 1.3, topY - r * 0.05);
      ctx.lineTo(cx + topW * 1.3, topY - r * 0.05);
      ctx.stroke();

      // aisladores
      for (const sx of [-1, 1]) {
        for (const hy of [-r * 0.05, r * 0.08]) {
          ctx.beginPath();
          ctx.arc(cx + sx * topW * 1.3, topY + hy, size * 0.04, 0, 7);
          ctx.fillStyle   = accent;
          ctx.fill();
          ctx.strokeStyle = fg;
          ctx.lineWidth   = 1;
          ctx.stroke();
        }
      }

      // chispa superior
      const pulse   = 0.5 + 0.5 * Math.abs(Math.sin(phase * 2 * Math.PI));
      const sparkR  = size * 0.06 * (1 + 0.5 * pulse);
      ctx.beginPath();
      ctx.arc(cx, topY - r * 0.18, sparkR, 0, 7);
      ctx.fillStyle   = accent;
      ctx.fill();
      ctx.strokeStyle = fg;
      ctx.lineWidth   = Math.max(1, size / 30);
      ctx.stroke();
    }
  }

  /* ---------- splash ---------- */
  function mount(opts = {}) {
    const duration = opts.duration ?? CFG.duration;
    const fade     = opts.fade     ?? CFG.fade;
    const onDone   = opts.onDone   ?? (() => {});
    const style    = opts.style    ?? CFG.style;

    const ov = document.createElement('div');
    ov.id = 'vatia-intro';
    Object.assign(ov.style, {
      position: 'fixed', inset: 0, zIndex: 99999,
      background: 'radial-gradient(circle at 50% 40%, #1E3A5F 0%, #0B1728 100%)',
      display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center',
      transition: `opacity ${fade}ms ease`,
      fontFamily: '"Segoe UI", system-ui, sans-serif'
    });

    const cv = document.createElement('canvas');
    const S  = Math.min(window.innerWidth, window.innerHeight) * 0.42;
    cv.width = cv.height = S * (window.devicePixelRatio || 1);
    cv.style.width = cv.style.height = S + 'px';
    ov.appendChild(cv);

    const title = document.createElement('div');
    title.textContent = 'VATIA';
    Object.assign(title.style, {
      color: '#fff', fontWeight: '700', letterSpacing: '14px',
      fontSize: Math.max(20, S * 0.11) + 'px',
      marginTop: '18px', textShadow: '0 0 24px #0E7490'
    });
    ov.appendChild(title);

    const sub = document.createElement('div');
    sub.textContent = 'SIMULADOR DE SISTEMAS DE POTENCIA KEVIN ANDRÉS GÓMEZ PEÑA  ASEP 2026-2';
    Object.assign(sub.style, {
      color: '#94A3B8', fontSize: Math.max(9, S * 0.032) + 'px',
      letterSpacing: '5px', marginTop: '6px'
    });
    ov.appendChild(sub);

    document.body.appendChild(ov);

    const ctx = cv.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    ctx.scale(dpr, dpr);
    const size = S;

    const t0 = performance.now();
    let raf;
    let finished = false;

    function finish() {
      if (finished) return;
      finished = true;
      cancelAnimationFrame(raf);
      ov.style.opacity = '0';
      setTimeout(() => {
        ov.remove();
        onDone();
      }, fade + 50);
    }

    function loop(t) {
      const elapsed = t - t0;
      const phase   = (elapsed % 2400) / 2400;
      ctx.clearRect(0, 0, size, size);
      drawLogo(ctx, size / 2, size / 2, size * 0.92, style, phase);

      if (elapsed < duration) raf = requestAnimationFrame(loop);
      else finish();
    }
    raf = requestAnimationFrame(loop);

    const skip  = () => finish();
    const onKey = e => { if (e.key === 'Escape') skip(); };
    ov.addEventListener('click', skip);
    window.addEventListener('keydown', onKey);
  }

  window.VatiaIntro = { mount, drawLogo, CFG };
})();