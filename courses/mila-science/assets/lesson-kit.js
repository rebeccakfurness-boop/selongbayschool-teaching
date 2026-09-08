/* Shared interactive-lesson kit — JS.
   Referenced by every lessons/lesson-NN/lesson.html via a relative <script>.
   Exposes window.LessonKit with reusable widget builders + the step
   navigation engine. Each lesson's own inline script supplies the content
   (data arrays) and calls these builders — bespoke one-off widgets for a
   single lesson (e.g. a custom diagram) can live directly in that lesson's
   own <script> block using the same CSS classes for visual consistency. */
(function (global) {
  const LessonKit = {};

  // ---------- Flip cards: term (front) -> definition (back) ----------
  // data: [{ term, hint, back, backClass? }]
  LessonKit.buildFlipGrid = function (container, data, opts = {}) {
    data.forEach((item) => {
      const el = document.createElement('div');
      el.className = 'flip';
      const backClass = item.backClass || opts.backClass || '';
      el.innerHTML = `<div class="flip-inner">
        <div class="flip-face flip-front"><div class="flip-term">${item.term}</div><div class="flip-hint">${item.hint || 'Tap to reveal ↻'}</div></div>
        <div class="flip-face flip-back ${backClass}"><p>${item.back}</p></div>
      </div>`;
      el.onclick = () => el.classList.toggle('flipped');
      container.appendChild(el);
    });
  };

  // ---------- Sort / classify activity ----------
  // data: [{ name, correct, reason }], opts: { choices: [...] }
  LessonKit.buildSortActivity = function (container, data, opts) {
    const choices = opts.choices;
    data.forEach((item) => {
      const row = document.createElement('div');
      row.className = 'sort-item';
      const optsHtml = choices.map((o) => `<button class="opt-btn" data-val="${o}">${o}</button>`).join('');
      row.innerHTML = `<div style="flex:1;"><div class="sort-name">${item.name}</div><div class="sort-reason">${item.reason}</div></div><div class="sort-options">${optsHtml}</div>`;
      row.querySelectorAll('.opt-btn').forEach((btn) => {
        btn.onclick = () => {
          if (row.classList.contains('answered')) return;
          row.classList.add('answered');
          row.querySelectorAll('.opt-btn').forEach((b) => {
            if (b.dataset.val === item.correct) b.classList.add('correct');
            else if (b === btn) b.classList.add('wrong');
          });
        };
      });
      container.appendChild(row);
    });
  };

  // ---------- Clickable reveal grid (e.g. determinants, shifters) ----------
  // data: [{ emj, lbl, desc }]
  LessonKit.buildShiftGrid = function (container, data) {
    data.forEach((item) => {
      const el = document.createElement('div');
      el.className = 'shift-card';
      el.innerHTML = `<div class="emj">${item.emj}</div><div class="lbl">${item.lbl}</div><div class="desc">${item.desc}</div>`;
      el.onclick = () => el.classList.toggle('open');
      container.appendChild(el);
    });
  };

  // ---------- Guess-then-reveal cards ----------
  // data: [{ text, ans, tags?: [{label, cls}] }]
  // cls one of: tag-up, tag-down, tag-neutral
  LessonKit.buildRevealCards = function (container, data, opts = {}) {
    const buttonLabel = opts.buttonLabel || 'Reveal answer';
    data.forEach((item) => {
      const el = document.createElement('div');
      el.className = 'reveal-card';
      const tagsHtml = (item.tags || []).map((t) => `<span class="tag ${t.cls}">${t.label}</span>`).join('');
      el.innerHTML = `<div class="reveal-top"><div class="reveal-text">${item.text}</div><button class="btn btn-sm btn-coral">${buttonLabel}</button></div>
        <div class="reveal-ans">${tagsHtml}<p style="margin:${tagsHtml ? '8px 0 0' : '0'};">${item.ans}</p></div>`;
      el.onclick = () => el.classList.toggle('open');
      container.appendChild(el);
    });
  };

  // ---------- Interactive reader: buttons that swap computed readouts ----------
  // data: [{ label, ...values }], readouts: [{ key, label, format? }]
  // Optionally pass onSelect(item, index) for extra behaviour (e.g. highlighting an SVG point).
  LessonKit.buildReader = function (btnContainer, readoutContainer, data, readouts, onSelect) {
    readoutContainer.innerHTML = readouts.map((r) => `<div class="readout-box"><div class="num" data-key="${r.key}">—</div><div class="lbl">${r.label}</div></div>`).join('');
    const numEls = {};
    readouts.forEach((r) => { numEls[r.key] = readoutContainer.querySelector(`[data-key="${r.key}"]`); });
    data.forEach((item, i) => {
      const b = document.createElement('button');
      b.className = 'reader-btn';
      b.textContent = item.label;
      b.onclick = () => {
        btnContainer.querySelectorAll('.reader-btn').forEach((x) => x.classList.remove('sel'));
        b.classList.add('sel');
        readouts.forEach((r) => {
          const val = item[r.key];
          numEls[r.key].textContent = r.format ? r.format(val) : val;
        });
        if (onSelect) onSelect(item, i);
      };
      btnContainer.appendChild(b);
    });
  };

  // ---------- Data table with an optional highlighted "best" row ----------
  LessonKit.buildDataTable = function (container, headers, rows, bestIndex) {
    let html = '<tr>' + headers.map((h) => `<th>${h}</th>`).join('') + '</tr>';
    rows.forEach((r, i) => {
      html += `<tr class="${i === bestIndex ? 'best' : ''}">` + r.map((c) => `<td>${c}</td>`).join('') + '</tr>';
    });
    container.innerHTML = html;
  };

  // ---------- Quiz with instant feedback ----------
  // data: [{ q, opts: [...], correct: index, fb }]
  LessonKit.buildQuiz = function (container, data) {
    data.forEach((item) => {
      const el = document.createElement('div');
      el.className = 'quiz-card';
      const optsHtml = item.opts.map((o, i) => `<button class="btn btn-sm btn-ghost" data-i="${i}">${o}</button>`).join('');
      el.innerHTML = `<div class="quiz-q">${item.q}</div><div class="quiz-opts">${optsHtml}</div><div class="quiz-fb"></div>`;
      const fb = el.querySelector('.quiz-fb');
      el.querySelectorAll('.quiz-opts button').forEach((btn) => {
        btn.onclick = () => {
          if (fb.classList.contains('correct')) return;
          const ok = Number(btn.dataset.i) === item.correct;
          fb.textContent = (ok ? '✓ ' : '✗ Not quite — ') + item.fb;
          fb.className = 'quiz-fb ' + (ok ? 'correct' : 'wrong');
          if (ok) { btn.style.background = 'var(--teal)'; btn.style.color = '#fff'; }
        };
      });
      container.appendChild(el);
    });
  };

  // ---------- Vocab recap list ----------
  LessonKit.buildVocabList = function (container, data) {
    data.forEach((item) => {
      const el = document.createElement('div');
      el.className = 'vocab-card';
      el.innerHTML = `<b>${item.t}:</b> ${item.d}`;
      container.appendChild(el);
    });
  };

  // ---------- Homework checklist ----------
  LessonKit.buildHomeworkChecklist = function (container, items) {
    container.innerHTML = items.map((t) => `<li><div class="hw-box"></div>${t}</li>`).join('');
  };

  // ---------- Worked-example step-by-step walkthrough ----------
  // steps: [{ label, detail }] — the last item is auto-highlighted as the final answer.
  LessonKit.buildStepsList = function (container, steps) {
    steps.forEach((s, i) => {
      const isFinal = i === steps.length - 1;
      const item = document.createElement('div');
      item.className = 'step-item' + (isFinal ? ' final' : '');
      item.innerHTML = `<div class="step-num">${isFinal ? '✓' : i + 1}</div>
        <div class="step-body"><div class="step-label">${s.label}</div><div class="step-detail">${s.detail}</div></div>`;
      container.appendChild(item);
    });
  };

  // ---------- Bar comparison (e.g. quantity demanded vs supplied) ----------
  // bars: [{ label, value, color }], opts: { max } (defaults to the largest value present)
  LessonKit.buildBarCompare = function (container, bars, opts = {}) {
    const max = opts.max || Math.max(...bars.map((b) => b.value)) * 1.1;
    bars.forEach((b) => {
      const pct = Math.max(6, Math.round((b.value / max) * 100));
      const row = document.createElement('div');
      row.className = 'bar-row';
      row.innerHTML = `<div class="bar-top"><span>${b.label}</span><span>${b.value}</span></div>
        <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:${b.color};"><span>${b.value}</span></div></div>`;
      container.appendChild(row);
    });
  };

  // ---------- Stepper navigation engine ----------
  // Call once after all step content is built. Wires dots, prev/next buttons, keyboard arrows.
  LessonKit.initStepper = function (opts = {}) {
    const steps = document.querySelectorAll('.step');
    const dotsWrap = document.getElementById('dots');
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const stepNow = document.getElementById('stepNow');
    const stepTotal = document.getElementById('stepTotal');
    let current = 0;

    steps.forEach((s, i) => {
      const d = document.createElement('button');
      d.className = 'dot' + (i === 0 ? ' active' : '');
      d.onclick = () => goTo(i);
      dotsWrap.appendChild(d);
    });
    if (stepTotal) stepTotal.textContent = steps.length;

    function goTo(i) {
      if (i < 0 || i >= steps.length) return;
      steps[current].classList.remove('active');
      dotsWrap.children[current].classList.remove('active');
      if (i > current) dotsWrap.children[current].classList.add('done');
      current = i;
      steps[current].classList.add('active');
      Array.from(dotsWrap.children).forEach((d, idx) => d.classList.toggle('active', idx === current));
      if (stepNow) stepNow.textContent = current + 1;
      if (prevBtn) prevBtn.disabled = current === 0;
      if (nextBtn) nextBtn.textContent = current === steps.length - 1 ? (opts.finishLabel || 'Finish 🎉') : 'Next →';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
    if (prevBtn) prevBtn.onclick = () => goTo(current - 1);
    if (nextBtn) nextBtn.onclick = () => { if (current < steps.length - 1) goTo(current + 1); };
    document.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight') goTo(current + 1);
      if (e.key === 'ArrowLeft') goTo(current - 1);
    });
    goTo(0);
    return { goTo };
  };

  global.LessonKit = LessonKit;
})(window);
