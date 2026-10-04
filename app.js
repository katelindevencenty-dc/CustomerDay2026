(() => {
  "use strict";

  const POINTER_ANGLE = -Math.PI / 2; // pointer sits at the top of the wheel
  const PALETTE = [
    "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
    "#42d4f4", "#f032e6", "#469990", "#9a6324", "#800000",
    "#808000", "#000075", "#e6a817", "#2f8f5b", "#c2185b",
  ];

  const $ = (id) => document.getElementById(id);
  const wheelCanvas = $("wheel");
  const wctx = wheelCanvas.getContext("2d");
  const fxCanvas = $("fx");
  const fctx = fxCanvas.getContext("2d");

  const state = {
    entries: [],      // [{ name, account }] currently on the wheel
    filled: [],       // during the build animation: which slices have landed
    rotation: 0,
    building: false,
    spinning: false,
    winnerIndex: -1,
    winners: [],
    flyers: [],
    confetti: [],
    fxRunning: false,
    soundOn: true,
  };

  // ---------- Randomness ----------

  function randFloat() {
    const buf = new Uint32Array(1);
    crypto.getRandomValues(buf);
    return buf[0] / 2 ** 32;
  }
  const randInt = (n) => Math.floor(randFloat() * n);

  function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = randInt(i + 1);
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }

  // ---------- Sound ----------

  let audio = null;
  let lastSound = 0;

  function beep(freq, duration, type = "sine", volume = 0.08, delay = 0) {
    if (!state.soundOn) return;
    audio = audio || new (window.AudioContext || window.webkitAudioContext)();
    const t = audio.currentTime + delay;
    const osc = audio.createOscillator();
    const gain = audio.createGain();
    osc.type = type;
    osc.frequency.setValueAtTime(freq, t);
    gain.gain.setValueAtTime(volume, t);
    gain.gain.exponentialRampToValueAtTime(0.0001, t + duration);
    osc.connect(gain).connect(audio.destination);
    osc.start(t);
    osc.stop(t + duration);
  }

  function throttledBeep(...args) {
    const now = performance.now();
    if (now - lastSound < 35) return;
    lastSound = now;
    beep(...args);
  }

  const fanfare = () =>
    [523, 659, 784, 1047].forEach((f, i) => beep(f, 0.35, "triangle", 0.12, i * 0.12));

  // ---------- File parsing ----------

  const normalizeHeader = (h) => String(h).trim().toLowerCase().replace(/[\s_]+/g, "_");
  const isYes = (v) => /^(y|yes|true|1)$/i.test(String(v).trim());

  function parseWorkbook(data) {
    const wb = XLSX.read(data, { type: "array" });

    for (const sheetName of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
        header: 1, defval: "", raw: false, blankrows: false,
      });

      // The table may not start on the first row, so look for its header row.
      const headerRow = rows.slice(0, 25).findIndex((r) => r.map(normalizeHeader).includes("contact_name"));
      if (headerRow === -1) continue;

      const headers = rows[headerRow].map(normalizeHeader);
      const col = (name) => headers.indexOf(name);
      const nameCol = col("contact_name");
      const accountCol = col("account_name");
      const raffleCol = col("raffle");
      if (raffleCol === -1) {
        throw new Error(`Found "contact_name" on sheet "${sheetName}" but no "raffle" column.`);
      }

      const entries = [];
      let excluded = 0;
      for (const row of rows.slice(headerRow + 1)) {
        const name = String(row[nameCol] ?? "").trim();
        if (!name) continue;
        if (!isYes(row[raffleCol])) { excluded++; continue; }
        entries.push({ name, account: accountCol === -1 ? "" : String(row[accountCol] ?? "").trim() });
      }
      return { entries, excluded };
    }
    throw new Error('Couldn\'t find a table with a "contact_name" column in this file.');
  }

  async function handleFile(file) {
    $("error").classList.add("hidden");
    try {
      const { entries, excluded } = parseWorkbook(await file.arrayBuffer());
      if (entries.length === 0) {
        throw new Error(`No one in "${file.name}" has raffle = Y (${excluded} row(s) excluded).`);
      }
      state.entries = shuffle(entries);
      state.winners = [];
      renderWinners();
      $("summary").textContent =
        `${entries.length} name${entries.length === 1 ? "" : "s"} in the draw` +
        (excluded ? ` · ${excluded} excluded (raffle ≠ Y)` : "") +
        ` · ${file.name}`;
      showView("wheel");
      startBuild();
    } catch (err) {
      $("error").textContent = err.message;
      $("error").classList.remove("hidden");
    }
  }

  // ---------- Views ----------

  function showView(which) {
    $("uploadView").classList.toggle("hidden", which !== "upload");
    $("wheelView").classList.toggle("hidden", which !== "wheel");
    $("newFileBtn").classList.toggle("hidden", which !== "wheel");
    if (which === "upload") $("summary").textContent = "";
    if (which === "wheel") resizeWheel();
  }

  function renderWinners() {
    $("winnerList").innerHTML = "";
    for (const w of state.winners) {
      const li = document.createElement("li");
      li.textContent = w.name;
      if (w.account) {
        const span = document.createElement("span");
        span.textContent = w.account;
        li.appendChild(span);
      }
      $("winnerList").appendChild(li);
    }
    $("noWinners").classList.toggle("hidden", state.winners.length > 0);
  }

  // ---------- Wheel drawing ----------

  function resizeWheel() {
    const dpr = window.devicePixelRatio || 1;
    const size = wheelCanvas.clientWidth;
    wheelCanvas.width = size * dpr;
    wheelCanvas.height = size * dpr;
    wctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    fxCanvas.width = window.innerWidth * dpr;
    fxCanvas.height = window.innerHeight * dpr;
    fctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    drawWheel();
  }

  function wheelGeometry() {
    const size = wheelCanvas.clientWidth;
    const outer = size / 2;
    const radius = outer * 0.93; // slice radius, inside the rim
    return { size, cx: outer, cy: outer, outer, radius };
  }

  function sliceColor(i, n) {
    // Avoid the last slice matching the first when they meet.
    if (n > 1 && i === n - 1 && i % PALETTE.length === 0) return PALETTE[2];
    return PALETTE[i % PALETTE.length];
  }

  function labelFontSize(radius, n) {
    const seg = (2 * Math.PI) / n;
    return Math.min(radius * 0.065, radius * seg * 0.55);
  }

  function fitText(ctx, text, maxWidth) {
    if (ctx.measureText(text).width <= maxWidth) return text;
    let t = text;
    while (t.length > 1 && ctx.measureText(t + "…").width > maxWidth) t = t.slice(0, -1);
    return t + "…";
  }

  function drawWheel() {
    const { size, cx, cy, outer, radius } = wheelGeometry();
    const n = state.entries.length;
    if (size === 0) return; // wheel view is hidden
    wctx.clearRect(0, 0, size, size);

    // Rim with light bulbs
    wctx.beginPath();
    wctx.arc(cx, cy, outer - 1, 0, Math.PI * 2);
    wctx.fillStyle = "#0b1422";
    wctx.fill();
    const bulbs = 36;
    for (let b = 0; b < bulbs; b++) {
      const a = state.rotation + (b / bulbs) * Math.PI * 2;
      const r = (outer + radius) / 2;
      wctx.beginPath();
      wctx.arc(cx + Math.cos(a) * r, cy + Math.sin(a) * r, (outer - radius) * 0.22, 0, Math.PI * 2);
      wctx.fillStyle = b % 2 ? "#ffe9a8" : "#ffc94a";
      wctx.fill();
    }

    // Empty face (visible while the wheel is being built)
    wctx.beginPath();
    wctx.arc(cx, cy, radius, 0, Math.PI * 2);
    wctx.fillStyle = "#22344f";
    wctx.fill();

    if (n === 0) return;

    const seg = (2 * Math.PI) / n;
    const fontSize = labelFontSize(radius, n);
    const showLabels = fontSize >= 6;
    wctx.font = `600 ${fontSize}px "Segoe UI", system-ui, sans-serif`;
    wctx.textAlign = "right";
    wctx.textBaseline = "middle";

    for (let i = 0; i < n; i++) {
      if (state.building && !state.filled[i]) continue;
      const start = state.rotation + i * seg;

      wctx.beginPath();
      wctx.moveTo(cx, cy);
      wctx.arc(cx, cy, radius, start, start + seg);
      wctx.closePath();
      wctx.fillStyle = sliceColor(i, n);
      wctx.fill();
      if (n > 1) {
        wctx.strokeStyle = "rgba(255,255,255,0.55)";
        wctx.lineWidth = n > 120 ? 0.5 : 1.5;
        wctx.stroke();
      }

      if (showLabels) {
        wctx.save();
        wctx.translate(cx, cy);
        wctx.rotate(start + seg / 2);
        wctx.fillStyle = "#fff";
        wctx.fillText(fitText(wctx, state.entries[i].name, radius * 0.7), radius - 12, 0);
        wctx.restore();
      }
    }
  }

  // ---------- Build animation: names fly onto the wheel ----------

  function startBuild() {
    const n = state.entries.length;
    state.building = true;
    state.filled = new Array(n).fill(false);
    state.rotation = 0;
    state.flyers = [];
    $("hub").disabled = true;
    $("wheelHint").textContent = "Adding everyone to the wheel…";

    const stagger = Math.max(15, Math.min(120, 3500 / n));
    const now = performance.now();
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    state.entries.forEach((entry, i) => {
      // Start each name at a random spot around the edge of the screen.
      const side = randInt(4);
      const along = randFloat();
      const from = [
        { x: along * vw, y: -30 },
        { x: vw + 30, y: along * vh },
        { x: along * vw, y: vh + 30 },
        { x: -30, y: along * vh },
      ][side];
      state.flyers.push({ index: i, name: entry.name, from, start: now + i * stagger, duration: 900 });
    });

    runFx();
  }

  function sliceTarget(i) {
    // Where slice i's label sits on screen right now.
    const rect = wheelCanvas.getBoundingClientRect();
    const { radius } = wheelGeometry();
    const n = state.entries.length;
    const seg = (2 * Math.PI) / n;
    const angle = state.rotation + (i + 0.5) * seg;
    const r = radius * 0.62;
    return {
      x: rect.left + rect.width / 2 + Math.cos(angle) * r,
      y: rect.top + rect.height / 2 + Math.sin(angle) * r,
      angle: Math.atan2(Math.sin(angle), Math.cos(angle)),
      fontSize: Math.max(6, labelFontSize(radius, n)),
    };
  }

  const easeInOut = (t) => (t < 0.5 ? 2 * t * t : 1 - (-2 * t + 2) ** 2 / 2);

  function drawFlyers(now) {
    let remaining = 0;
    for (const f of state.flyers) {
      if (f.done) continue;
      const t = (now - f.start) / f.duration;
      if (t < 0) { remaining++; continue; }
      if (t >= 1) {
        f.done = true;
        state.filled[f.index] = true;
        throttledBeep(500 + randInt(500), 0.06, "sine", 0.05);
        continue;
      }
      remaining++;
      const e = easeInOut(t);
      const target = sliceTarget(f.index);
      const x = f.from.x + (target.x - f.from.x) * e;
      const y = f.from.y + (target.y - f.from.y) * e;
      const size = 26 + (target.fontSize - 26) * e;

      fctx.save();
      fctx.translate(x, y);
      fctx.rotate(target.angle * e);
      fctx.font = `700 ${size}px "Segoe UI", system-ui, sans-serif`;
      fctx.textAlign = "center";
      fctx.textBaseline = "middle";
      fctx.shadowColor = "rgba(0,0,0,0.6)";
      fctx.shadowBlur = 6;
      fctx.fillStyle = "#fff";
      fctx.fillText(f.name, 0, 0);
      fctx.restore();
    }
    return remaining;
  }

  function finishBuild() {
    state.building = false;
    state.flyers = [];
    drawWheel();
    $("hub").disabled = false;
    $("wheelHint").textContent = "Click the wheel (or press Space) to spin!";
  }

  // ---------- Confetti ----------

  function launchConfetti() {
    const vw = window.innerWidth;
    for (let i = 0; i < 260; i++) {
      state.confetti.push({
        x: randFloat() * vw,
        y: -20 - randFloat() * 300,
        vx: (randFloat() - 0.5) * 4,
        vy: 2 + randFloat() * 4,
        size: 6 + randFloat() * 8,
        spin: randFloat() * Math.PI,
        vspin: (randFloat() - 0.5) * 0.3,
        color: PALETTE[randInt(PALETTE.length)],
      });
    }
    runFx();
  }

  function drawConfetti() {
    const vh = window.innerHeight;
    state.confetti = state.confetti.filter((p) => p.y < vh + 30);
    for (const p of state.confetti) {
      p.x += p.vx;
      p.y += p.vy;
      p.vy += 0.05;
      p.spin += p.vspin;
      fctx.save();
      fctx.translate(p.x, p.y);
      fctx.rotate(p.spin);
      fctx.fillStyle = p.color;
      fctx.fillRect(-p.size / 2, -p.size / 4, p.size, p.size / 2);
      fctx.restore();
    }
    return state.confetti.length;
  }

  // ---------- Shared animation loop for the overlay ----------

  function runFx() {
    if (state.fxRunning) return;
    state.fxRunning = true;

    const frame = (now) => {
      fctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
      let active = 0;

      if (state.building) {
        state.rotation += 0.004; // gentle turn while names land
        const flying = drawFlyers(now);
        drawWheel();
        if (flying === 0) finishBuild();
        else active++;
      }
      if (state.confetti.length) active += drawConfetti();

      if (active) requestAnimationFrame(frame);
      else state.fxRunning = false;
    };
    requestAnimationFrame(frame);
  }

  // ---------- Spinning ----------

  const mod = (a, m) => ((a % m) + m) % m;

  function indexUnderPointer() {
    const n = state.entries.length;
    const seg = (2 * Math.PI) / n;
    return Math.floor(mod(POINTER_ANGLE - state.rotation, 2 * Math.PI) / seg) % n;
  }

  function spin() {
    const n = state.entries.length;
    if (state.spinning || state.building || n === 0 || !$("winnerModal").classList.contains("hidden")) return;

    state.spinning = true;
    $("hub").disabled = true;
    $("wheelHint").textContent = "Good luck everyone…";

    const seg = (2 * Math.PI) / n;
    const winner = randInt(n);
    const jitter = (randFloat() - 0.5) * 0.7 * seg; // don't always stop dead-centre
    const target = POINTER_ANGLE - (winner + 0.5) * seg - jitter;
    const startRot = state.rotation;
    const endRot = startRot + mod(target - startRot, 2 * Math.PI) + 2 * Math.PI * (6 + randInt(3));
    const duration = 7000 + randInt(2000);
    const t0 = performance.now();
    let lastIndex = indexUnderPointer();

    const frame = (now) => {
      const t = Math.min(1, (now - t0) / duration);
      const eased = 1 - (1 - t) ** 4;
      state.rotation = startRot + (endRot - startRot) * eased;
      drawWheel();

      const idx = indexUnderPointer();
      if (idx !== lastIndex) {
        lastIndex = idx;
        throttledBeep(1400, 0.025, "square", 0.04);
      }

      if (t < 1) requestAnimationFrame(frame);
      else finishSpin();
    };
    requestAnimationFrame(frame);
  }

  function finishSpin() {
    state.spinning = false;
    state.winnerIndex = indexUnderPointer();
    const w = state.entries[state.winnerIndex];
    $("winnerName").textContent = w.name;
    $("winnerAccount").textContent = w.account;
    $("winnerModal").classList.remove("hidden");
    fanfare();
    launchConfetti();
  }

  function closeWinner(remove) {
    const w = state.entries[state.winnerIndex];
    state.winners.push(w);
    renderWinners();
    if (remove) state.entries.splice(state.winnerIndex, 1);
    state.winnerIndex = -1;
    $("winnerModal").classList.add("hidden");
    drawWheel();

    const left = state.entries.length;
    $("hub").disabled = left === 0;
    $("wheelHint").textContent = left === 0
      ? "Everyone has won — load another file to start again."
      : "Click the wheel (or press Space) to spin again!";
  }

  // ---------- Events ----------

  $("fileInput").addEventListener("change", (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
    e.target.value = "";
  });

  const dropZone = $("dropZone");
  ["dragenter", "dragover"].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => { e.preventDefault(); dropZone.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => { e.preventDefault(); dropZone.classList.remove("over"); }));
  dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  // Stop the browser opening a file dropped outside the drop zone.
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());

  wheelCanvas.addEventListener("click", spin);
  $("hub").addEventListener("click", spin);
  $("removeBtn").addEventListener("click", () => closeWinner(true));
  $("keepBtn").addEventListener("click", () => closeWinner(false));

  $("newFileBtn").addEventListener("click", () => {
    if (state.spinning || state.building) return;
    state.entries = [];
    showView("upload");
  });

  $("soundBtn").addEventListener("click", () => {
    state.soundOn = !state.soundOn;
    $("soundBtn").textContent = state.soundOn ? "🔊 Sound on" : "🔇 Sound off";
  });

  document.addEventListener("keydown", (e) => {
    if ($("wheelView").classList.contains("hidden")) return;
    if (e.code === "Space") {
      e.preventDefault();
      if ($("winnerModal").classList.contains("hidden")) spin();
    } else if (e.code === "Enter" && !$("winnerModal").classList.contains("hidden")) {
      closeWinner(true);
    }
  });

  window.addEventListener("resize", resizeWheel);
  resizeWheel();
})();
