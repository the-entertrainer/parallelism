/**
 * PARALLELISM — Interactive Training Web App
 * 3D glassmorphism presentation with morph-inspired transitions
 * for Instructional Designers & Content Developers
 */

import * as THREE from 'three';

// ---------------------------------------------------------------------------
// Slide data
// ---------------------------------------------------------------------------
const SLIDES = [
  {
    id: 'hook',
    title: 'Hook',
    html: `
      <div class="slide-inner">
        <p class="kicker">Icebreaker · 60 seconds</p>
        <h1 class="display">Which of these are<br><em style="font-style:normal;background:linear-gradient(90deg,#00e5ff,#a855f7);-webkit-background-clip:text;background-clip:text;color:transparent">truly parallel</em>?</h1>
        <p class="lead">Tap the pairs that share the same structure or direction. There are more than you think.</p>
        <div class="icebreaker-grid" id="ice-grid" role="group" aria-label="Parallel pairs challenge">
          <button class="ice-card glass" data-pair="lanes" aria-pressed="false">
            <span class="emoji" aria-hidden="true">🛣️</span>
            <span>Highway lanes</span>
          </button>
          <button class="ice-card glass" data-pair="cores" aria-pressed="false">
            <span class="emoji" aria-hidden="true">💻</span>
            <span>CPU cores</span>
          </button>
          <button class="ice-card glass" data-pair="swim" aria-pressed="false">
            <span class="emoji" aria-hidden="true">🏊</span>
            <span>Synchronized swimming</span>
          </button>
          <button class="ice-card glass" data-pair="rails" aria-pressed="false">
            <span class="emoji" aria-hidden="true">🚂</span>
            <span>Train tracks</span>
          </button>
          <button class="ice-card glass" data-pair="obj" aria-pressed="false">
            <span class="emoji" aria-hidden="true">🎯</span>
            <span>Learning objectives</span>
          </button>
          <button class="ice-card glass" data-pair="lists" aria-pressed="false">
            <span class="emoji" aria-hidden="true">📋</span>
            <span>Bullet lists</span>
          </button>
        </div>
        <p class="body" id="ice-feedback" style="margin-top:1rem;min-height:1.5em;color:var(--accent-cyan)"></p>
      </div>
    `
  },
  {
    id: 'why',
    title: 'Why it matters',
    html: `
      <div class="slide-inner">
        <p class="kicker">The case for parallelism</p>
        <h1 class="display">Why instructional designers obsess over parallel structure</h1>
        <p class="lead">Parallelism is not just grammar. It is cognitive kindness.</p>
        <div class="card-grid">
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">⚡</div>
            <h3>Scannability</h3>
            <p>Readers process patterned lists 40% faster. Your objectives, steps, and features become instantly graspable.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">🧠</div>
            <h3>Memory</h3>
            <p>The brain loves rhythm. Parallel constructions create a mental beat that sticks long after the course ends.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">♿</div>
            <h3>Accessibility</h3>
            <p>Screen readers and cognitive load both benefit from predictable, consistent phrasing.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">✨</div>
            <h3>Professional polish</h3>
            <p>Non-parallel lists scream “rushed draft.” Parallel ones signal craft and care.</p>
          </article>
        </div>
      </div>
    `
  },
  {
    id: 'what',
    title: 'What is it',
    html: `
      <div class="slide-inner">
        <p class="kicker">Definition</p>
        <h1 class="display">Parallel structure<br>in plain language</h1>
        <div class="glass" style="padding:1.5rem 1.75rem;max-width:540px;text-align:left">
          <p style="font-size:var(--text-lg);color:var(--fg-primary);margin-bottom:1rem">
            When items in a series, list, comparison, or set of headings share the <strong style="color:var(--accent-cyan)">same grammatical form</strong>.
          </p>
          <p class="body" style="margin:0">
            Nouns with nouns. Verbs with verbs. “-ing” with “-ing”. Complete sentences with complete sentences. The pattern itself becomes meaning.
          </p>
        </div>
        <p class="lead" style="margin-top:1.5rem">It is the difference between noise and music in your content.</p>
      </div>
    `
  },
  {
    id: 'examples',
    title: 'Examples',
    html: `
      <div class="slide-inner">
        <p class="kicker">See the difference</p>
        <h2>Learning objectives — before & after</h2>
        <div class="example-pair">
          <div class="example-box bad glass">
            <div class="example-label"><span aria-hidden="true">✗</span> Not parallel</div>
            <ul>
              <li>Understand the sales process</li>
              <li>How to close a deal</li>
              <li>Writing follow-up emails</li>
              <li>Learners will practice objection handling</li>
            </ul>
          </div>
          <div class="example-box good glass">
            <div class="example-label"><span aria-hidden="true">✓</span> Parallel</div>
            <ul>
              <li>Map the sales process</li>
              <li>Close a deal confidently</li>
              <li>Write effective follow-up emails</li>
              <li>Handle objections in real time</li>
            </ul>
          </div>
        </div>
        <p class="body" style="margin-top:1.25rem">Notice the consistent verb form? That is the entire game.</p>
      </div>
    `
  },
  {
    id: 'examples2',
    title: 'More examples',
    html: `
      <div class="slide-inner">
        <p class="kicker">Everyday content</p>
        <h2>Bullets, UI text & headings</h2>
        <div class="example-pair">
          <div class="example-box bad glass">
            <div class="example-label"><span aria-hidden="true">✗</span> Mixed forms</div>
            <ul>
              <li>Fast onboarding</li>
              <li>You can customize templates</li>
              <li>Analytics that matter</li>
              <li>Integrating with Slack</li>
            </ul>
          </div>
          <div class="example-box good glass">
            <div class="example-label"><span aria-hidden="true">✓</span> Clean pattern</div>
            <ul>
              <li>Onboard in minutes</li>
              <li>Customize every template</li>
              <li>See analytics that matter</li>
              <li>Integrate with Slack</li>
            </ul>
          </div>
        </div>
        <div class="glass" style="padding:1rem 1.25rem;margin-top:1.25rem;max-width:520px;text-align:left">
          <p style="font-size:var(--text-sm);color:var(--fg-secondary);margin:0">
            <strong style="color:var(--accent-cyan)">Pro tip:</strong> Read the list aloud. If the rhythm breaks, the structure is broken.
          </p>
        </div>
      </div>
    `
  },
  {
    id: 'rules',
    title: 'Rules',
    html: `
      <div class="slide-inner">
        <p class="kicker">Practical rules</p>
        <h1 class="display">Five rules you can apply today</h1>
        <ol class="takeaway-list" style="counter-reset:none;list-style:none">
          <li>
            <span class="takeaway-num">1</span>
            <span>Match form: all gerunds, all infinitives, all imperative verbs, or all noun phrases.</span>
          </li>
          <li>
            <span class="takeaway-num">2</span>
            <span>Keep articles & prepositions consistent (or omit them all).</span>
          </li>
          <li>
            <span class="takeaway-num">3</span>
            <span>In comparisons use the same structure on both sides of “than” or “as”.</span>
          </li>
          <li>
            <span class="takeaway-num">4</span>
            <span>Headings and sub-headings in a section should share the same pattern.</span>
          </li>
          <li>
            <span class="takeaway-num">5</span>
            <span>When in doubt, rewrite the entire series rather than forcing one item.</span>
          </li>
        </ol>
      </div>
    `
  },
  {
    id: 'practice1',
    title: 'Practice 1',
    html: `
      <div class="slide-inner">
        <p class="kicker">Practice · Round 1</p>
        <h2>Which list is parallel?</h2>
        <p class="body">Select the correctly structured set of course outcomes.</p>
        <div class="practice-area" id="practice-1" role="radiogroup" aria-label="Choose the parallel list">
          <button class="practice-item glass" data-correct="false" role="radio" aria-checked="false">
            <strong>A.</strong> Identify customer needs · Building rapport · How to ask discovery questions · Close the sale
          </button>
          <button class="practice-item glass" data-correct="true" role="radio" aria-checked="false">
            <strong>B.</strong> Identify customer needs · Build rapport quickly · Ask discovery questions · Close the sale
          </button>
          <button class="practice-item glass" data-correct="false" role="radio" aria-checked="false">
            <strong>C.</strong> Identifying needs · Build rapport · Discovery questions · The close
          </button>
        </div>
        <div class="feedback" id="feedback-1" role="status" aria-live="polite"></div>
      </div>
    `
  },
  {
    id: 'practice2',
    title: 'Practice 2',
    html: `
      <div class="slide-inner">
        <p class="kicker">Practice · Round 2</p>
        <h2>Fix the non-parallel objective</h2>
        <p class="body">Original: “By the end of this module, learners will understand the platform, be able to create a project, and navigation of the dashboard.”</p>
        <div class="practice-area" id="practice-2" role="radiogroup" aria-label="Choose the best rewrite">
          <button class="practice-item glass" data-correct="false" role="radio" aria-checked="false">
            Understand the platform, be able to create a project, and navigation of the dashboard.
          </button>
          <button class="practice-item glass" data-correct="true" role="radio" aria-checked="false">
            Understand the platform, create a project, and navigate the dashboard.
          </button>
          <button class="practice-item glass" data-correct="false" role="radio" aria-checked="false">
            Understanding the platform, creating a project, and the ability to navigate.
          </button>
        </div>
        <div class="feedback" id="feedback-2" role="status" aria-live="polite"></div>
        <p class="body" style="margin-top:1rem;font-size:var(--text-sm)">Hint: pure verb forms in series create the strongest parallel.</p>
      </div>
    `
  },
  {
    id: 'advanced',
    title: 'Advanced',
    html: `
      <div class="slide-inner">
        <p class="kicker">Beyond sentences</p>
        <h1 class="display">Parallelism lives everywhere</h1>
        <div class="card-grid">
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">🖼️</div>
            <h3>Visual design</h3>
            <p>Icons, card layouts, and illustration styles that share the same visual language feel parallel and intentional.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">🎬</div>
            <h3>Multimedia</h3>
            <p>Consistent shot framing, motion timing, and voice-over cadence create a parallel experience across modules.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">🔀</div>
            <h3>Branching scenarios</h3>
            <p>Choice labels written in the same grammatical form keep decision points fair and scannable.</p>
          </article>
          <article class="card glass">
            <div class="card-icon" aria-hidden="true">📐</div>
            <h3>Assessment items</h3>
            <p>Options in multiple-choice questions that are parallel in length and structure reduce bias and cognitive noise.</p>
          </article>
        </div>
      </div>
    `
  },
  {
    id: 'summary',
    title: 'Summary',
    html: `
      <div class="slide-inner">
        <p class="kicker">Key takeaways</p>
        <h1 class="display">Leave with these</h1>
        <ul class="takeaway-list">
          <li>
            <span class="takeaway-num">1</span>
            <span>Parallel structure is a design decision, not just a grammar rule.</span>
          </li>
          <li>
            <span class="takeaway-num">2</span>
            <span>Match grammatical form across lists, objectives, headings, and choices.</span>
          </li>
          <li>
            <span class="takeaway-num">3</span>
            <span>Read your content aloud — broken rhythm reveals broken parallelism.</span>
          </li>
          <li>
            <span class="takeaway-num">4</span>
            <span>Extend the principle to visuals, motion, and interaction patterns.</span>
          </li>
          <li>
            <span class="takeaway-num">5</span>
            <span>When one item fights the pattern, rewrite the whole series.</span>
          </li>
        </ul>
        <p class="lead" style="margin-top:1.5rem">Now go make your next course feel like music.</p>
        <span class="badge" style="margin-top:0.75rem">You finished the training · Parallelism</span>
      </div>
    `
  }
];

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let current = 0;
let isTransitioning = false;
const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

// ---------------------------------------------------------------------------
// DOM refs
// ---------------------------------------------------------------------------
const stage = document.getElementById('stage');
const progressFill = document.getElementById('progress-fill');
const progressDots = document.getElementById('progress-dots');
const btnPrev = document.getElementById('btn-prev');
const btnNext = document.getElementById('btn-next');
const keyHint = document.getElementById('key-hint');
const motionNotice = document.getElementById('motion-notice');

// ---------------------------------------------------------------------------
// Build slides & dots
// ---------------------------------------------------------------------------
function build() {
  SLIDES.forEach((s, i) => {
    const el = document.createElement('section');
    el.className = 'slide' + (i === 0 ? ' active' : '');
    el.id = `slide-${s.id}`;
    el.setAttribute('aria-hidden', i === 0 ? 'false' : 'true');
    el.setAttribute('aria-label', `Slide ${i + 1} of ${SLIDES.length}: ${s.title}`);
    el.innerHTML = s.html;
    stage.appendChild(el);

    const dot = document.createElement('button');
    dot.className = 'progress-dot';
    dot.setAttribute('aria-label', `Go to slide ${i + 1}: ${s.title}`);
    if (i === 0) dot.setAttribute('aria-current', 'true');
    dot.addEventListener('click', () => goTo(i));
    progressDots.appendChild(dot);
  });

  if (reducedMotion) {
    motionNotice.hidden = false;
  }

  updateControls();
  initIcebreaker();
  initPractice();
}

// ---------------------------------------------------------------------------
// Navigation
// ---------------------------------------------------------------------------
function goTo(index) {
  if (isTransitioning || index === current || index < 0 || index >= SLIDES.length) return;
  isTransitioning = true;

  const slides = stage.querySelectorAll('.slide');
  const from = slides[current];
  const to = slides[index];

  from.classList.remove('active');
  from.classList.add('exiting');
  from.setAttribute('aria-hidden', 'true');

  to.classList.add('active');
  to.setAttribute('aria-hidden', 'false');

  // Update dots
  const dots = progressDots.querySelectorAll('.progress-dot');
  dots[current].removeAttribute('aria-current');
  dots[index].setAttribute('aria-current', 'true');

  current = index;
  updateControls();
  updateProgress();

  // Clean up after transition
  setTimeout(() => {
    from.classList.remove('exiting');
    isTransitioning = false;
    // Focus management for a11y
    const heading = to.querySelector('h1, h2');
    if (heading) heading.setAttribute('tabindex', '-1');
  }, reducedMotion ? 20 : 720);
}

function next() {
  if (current < SLIDES.length - 1) goTo(current + 1);
}

function prev() {
  if (current > 0) goTo(current - 1);
}

function updateControls() {
  btnPrev.disabled = current === 0;
  const isLast = current === SLIDES.length - 1;
  btnNext.querySelector('.btn-label').textContent =
    current === 0 ? 'Start' : isLast ? 'Restart' : 'Next';
  if (isLast) {
    btnNext.onclick = () => goTo(0);
  } else {
    btnNext.onclick = next;
  }
}

function updateProgress() {
  const pct = ((current + 1) / SLIDES.length) * 100;
  progressFill.style.width = `${pct}%`;
}

// ---------------------------------------------------------------------------
// Icebreaker interaction
// ---------------------------------------------------------------------------
function initIcebreaker() {
  const grid = document.getElementById('ice-grid');
  const feedback = document.getElementById('ice-feedback');
  if (!grid) return;

  const allPairs = new Set(['lanes', 'cores', 'swim', 'rails', 'obj', 'lists']);
  let selected = new Set();

  grid.addEventListener('click', (e) => {
    const card = e.target.closest('.ice-card');
    if (!card) return;
    const pair = card.dataset.pair;
    const pressed = card.getAttribute('aria-pressed') === 'true';
    card.setAttribute('aria-pressed', String(!pressed));
    card.classList.toggle('selected', !pressed);
    if (!pressed) selected.add(pair);
    else selected.delete(pair);

    if (selected.size === allPairs.size) {
      feedback.textContent = 'Perfect. Everything here can be parallel — including your learning content.';
    } else if (selected.size > 0) {
      feedback.textContent = `${selected.size} of ${allPairs.size} selected. Keep going…`;
    } else {
      feedback.textContent = '';
    }
  });
}

// ---------------------------------------------------------------------------
// Practice interactions
// ---------------------------------------------------------------------------
function initPractice() {
  setupPractice('practice-1', 'feedback-1', 'Yes! All imperative verbs in the same form create clean parallel structure.');
  setupPractice('practice-2', 'feedback-2', 'Excellent. Consistent infinitives keep the series parallel and professional.');
}

function setupPractice(areaId, feedbackId, successMsg) {
  const area = document.getElementById(areaId);
  const feedback = document.getElementById(feedbackId);
  if (!area) return;

  area.addEventListener('click', (e) => {
    const item = e.target.closest('.practice-item');
    if (!item) return;

    // Clear previous
    area.querySelectorAll('.practice-item').forEach(el => {
      el.classList.remove('selected', 'correct', 'incorrect');
      el.setAttribute('aria-checked', 'false');
    });

    item.classList.add('selected');
    item.setAttribute('aria-checked', 'true');

    const isCorrect = item.dataset.correct === 'true';
    item.classList.add(isCorrect ? 'correct' : 'incorrect');

    feedback.className = 'feedback show ' + (isCorrect ? 'success' : 'error');
    feedback.textContent = isCorrect
      ? successMsg
      : 'Not quite. Look for consistent grammatical form across every item.';
  });
}

// ---------------------------------------------------------------------------
// Keyboard & touch
// ---------------------------------------------------------------------------
document.addEventListener('keydown', (e) => {
  if (e.target.closest('button, a, input, textarea')) return;
  if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
    e.preventDefault();
    next();
  } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
    e.preventDefault();
    prev();
  } else if (e.key === 'Home') {
    e.preventDefault();
    goTo(0);
  } else if (e.key === 'End') {
    e.preventDefault();
    goTo(SLIDES.length - 1);
  }
});

// Touch swipe
let touchStartX = 0;
stage.addEventListener('touchstart', (e) => {
  touchStartX = e.changedTouches[0].screenX;
}, { passive: true });

stage.addEventListener('touchend', (e) => {
  const dx = e.changedTouches[0].screenX - touchStartX;
  if (Math.abs(dx) > 60) {
    if (dx < 0) next();
    else prev();
  }
}, { passive: true });

// Hide key hint after first interaction
let hintHidden = false;
function hideHint() {
  if (hintHidden) return;
  hintHidden = true;
  keyHint.classList.add('hidden');
}
btnNext.addEventListener('click', hideHint);
btnPrev.addEventListener('click', hideHint);
document.addEventListener('keydown', hideHint, { once: true });

// ---------------------------------------------------------------------------
// Three.js background — parallel lines / data streams
// ---------------------------------------------------------------------------
function initThree() {
  const canvas = document.getElementById('bg-canvas');
  if (!canvas || reducedMotion) return;

  const renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.setSize(window.innerWidth, window.innerHeight);
  renderer.setClearColor(0x000000, 0);

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(55, window.innerWidth / window.innerHeight, 0.1, 100);
  camera.position.z = 18;

  // Parallel lines geometry
  const group = new THREE.Group();
  const lineCount = 28;
  const positions = [];
  const colors = [];

  const colorA = new THREE.Color(0x00e5ff);
  const colorB = new THREE.Color(0xa855f7);

  for (let i = 0; i < lineCount; i++) {
    const x = (i / (lineCount - 1) - 0.5) * 28;
    const yStart = -14 + Math.random() * 4;
    const yEnd = 14 - Math.random() * 4;
    const z = (Math.random() - 0.5) * 10;

    positions.push(x, yStart, z, x, yEnd, z);

    const t = i / (lineCount - 1);
    const c = colorA.clone().lerp(colorB, t);
    colors.push(c.r, c.g, c.b, c.r, c.g, c.b);
  }

  const geo = new THREE.BufferGeometry();
  geo.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
  geo.setAttribute('color', new THREE.Float32BufferAttribute(colors, 3));

  const mat = new THREE.LineBasicMaterial({
    vertexColors: true,
    transparent: true,
    opacity: 0.22,
    blending: THREE.AdditiveBlending
  });

  const lines = new THREE.LineSegments(geo, mat);
  group.add(lines);
  scene.add(group);

  // Soft particles
  const particleCount = 120;
  const pGeo = new THREE.BufferGeometry();
  const pPos = new Float32Array(particleCount * 3);
  for (let i = 0; i < particleCount * 3; i++) {
    pPos[i] = (Math.random() - 0.5) * 40;
  }
  pGeo.setAttribute('position', new THREE.BufferAttribute(pPos, 3));
  const pMat = new THREE.PointsMaterial({
    size: 0.08,
    color: 0x00e5ff,
    transparent: true,
    opacity: 0.5,
    blending: THREE.AdditiveBlending,
    depthWrite: false
  });
  const particles = new THREE.Points(pGeo, pMat);
  scene.add(particles);

  let mouseX = 0, mouseY = 0;
  window.addEventListener('pointermove', (e) => {
    mouseX = (e.clientX / window.innerWidth - 0.5) * 2;
    mouseY = (e.clientY / window.innerHeight - 0.5) * 2;
  });

  function onResize() {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
  }
  window.addEventListener('resize', onResize);

  const clock = new THREE.Clock();
  function animate() {
    requestAnimationFrame(animate);
    const t = clock.getElapsedTime();

    group.rotation.z = Math.sin(t * 0.08) * 0.04;
    group.rotation.y = mouseX * 0.08;
    group.rotation.x = mouseY * 0.05;

    particles.rotation.y = t * 0.03;
    particles.position.y = Math.sin(t * 0.2) * 0.4;

    // Subtle opacity pulse
    mat.opacity = 0.18 + Math.sin(t * 0.5) * 0.06;

    renderer.render(scene, camera);
  }
  animate();
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
build();
updateProgress();
initThree();

// Expose for debugging if needed
window.__parallelism = { goTo, next, prev, current: () => current };
