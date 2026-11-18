/* Excelクエスト 画面と進行 */
(function () {
  'use strict';
  const { CHAPTERS, STAGES } = window.XQ;
  const $app = document.getElementById('app');
  const KEY = 'excel-quest-v1';

  // ---------- 記録 ----------
  const blank = () => ({ stars: {}, xp: 0, streak: { last: '', n: 0 }, missed: {}, examBest: null, unlockAll: false, answered: 0 });
  function load() {
    try { const s = JSON.parse(localStorage.getItem(KEY)); if (s && typeof s === 'object') return Object.assign(blank(), s); } catch (e) { /* 読めないときは新規 */ }
    return blank();
  }
  let S = load();
  function save() { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) { /* 保存できない環境 */ } }
  const today = () => { const d = new Date(); return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate(); };
  function touchStreak() {
    const t = today();
    if (S.streak.last === t) return;
    const y = new Date(); y.setDate(y.getDate() - 1);
    const ys = y.getFullYear() + '-' + (y.getMonth() + 1) + '-' + y.getDate();
    S.streak.n = S.streak.last === ys ? S.streak.n + 1 : 1;
    S.streak.last = t;
  }
  const level = () => Math.floor(S.xp / 100) + 1;
  const isCleared = (id) => (S.stars[id] || 0) > 0;
  function isUnlocked(i) { return S.unlockAll || i === 0 || isCleared(STAGES[i - 1].id); }

  // ---------- 小物 ----------
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const shuffle = (a) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
  function h(html) { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstChild; }
  function screen(html) { $app.innerHTML = html; window.scrollTo(0, 0); }

  // ---------- 表 ----------
  function gridHTML(g, opt) {
    opt = opt || {};
    const cells = Object.assign({}, g.cells, opt.override || {});
    const sheet = XL.makeSheet(cells, g.names);
    let maxC = 1, maxR = 1;
    for (const k of Object.keys(cells)) { const a = XL.parseAddr(k); if (a) { maxC = Math.max(maxC, a.col); maxR = Math.max(maxR, a.row); } }
    if (opt.target) { const a = XL.parseAddr(opt.target); maxC = Math.max(maxC, a.col); maxR = Math.max(maxR, a.row); }
    if (opt.fill) { const a = XL.parseAddr(opt.fill); maxR = Math.max(maxR, a.row); }
    const t = opt.target ? XL.parseAddr(opt.target) : null;
    const f = opt.fill ? XL.parseAddr(opt.fill) : null;
    const hl = new Set(g.hl || []);
    let s = '<div class="gridwrap"><table class="grid' + (opt.tap ? ' tap' : '') + '"><thead><tr><th></th>';
    for (let c = 1; c <= maxC; c++) s += '<th>' + XL.numToCol(c) + '</th>';
    s += '</tr></thead><tbody>';
    for (let r = 1; r <= maxR; r++) {
      s += '<tr><th>' + r + '</th>';
      for (let c = 1; c <= maxC; c++) {
        const a = XL.addr(c, r);
        const raw = cells[a];
        let v = XL.cellValue(sheet, a);
        const isDate = raw && typeof raw === 'object' && 'd' in raw;
        let txt = XL.format(v, isDate || (opt.fmt === 'date' && opt.override && a in opt.override) ? 'date' : null);
        const cls = [];
        if (typeof v === 'number') cls.push('num');
        if (r === 1) cls.push('head');
        if (hl.has(a)) cls.push(g.hlColor === 'red' ? 'hlred' : 'hl');
        if (t && c === t.col && r === t.row) { cls.push('target'); if (!opt.override || !(a in opt.override)) txt = '?'; }
        else if (t && f && c === t.col && r > t.row && r <= f.row) { cls.push('fillcell'); if (!opt.override || !(a in opt.override)) txt = '↓'; }
        s += '<td data-a="' + a + '" class="' + cls.join(' ') + '">' + esc(txt) + '</td>';
      }
      s += '</tr>';
    }
    s += '</tbody></table></div>';
    if (g.names) s += '<div class="names">' + Object.entries(g.names).map(([k, v]) => '🏷️ 名前 <b>' + esc(k) + '</b> ＝ ' + v).join('<br>') + '</div>';
    return s;
  }

  // ---------- ホーム ----------
  function home() {
    const totalStars = STAGES.reduce((n, s) => n + (S.stars[s.id] || 0), 0);
    const cleared = STAGES.filter((s) => isCleared(s.id)).length;
    const lv = level(), into = S.xp % 100;
    let html = '<header class="top"><div class="brand">📗 Excelクエスト</div><button class="icon" id="menuBtn" aria-label="設定">⚙️</button></header>';
    html += '<section class="stats">' +
      '<div class="stat"><div class="big">Lv.' + lv + '</div><div class="bar"><i style="width:' + into + '%"></i></div><div class="small">次まで ' + (100 - into) + ' XP</div></div>' +
      '<div class="stat"><div class="big">🔥 ' + (S.streak.n || 0) + '</div><div class="small">日連続</div></div>' +
      '<div class="stat"><div class="big">⭐ ' + totalStars + '</div><div class="small">' + cleared + '/' + STAGES.length + ' クリア</div></div>' +
      '</section>';
    const nextIdx = STAGES.findIndex((s, i) => !isCleared(s.id) && isUnlocked(i));
    if (nextIdx >= 0) html += '<button class="cta" data-go="' + nextIdx + '">▶ つづきから：' + esc(STAGES[nextIdx].title) + '</button>';
    const missedN = Object.values(S.missed).filter((n) => n > 0).length;
    if (cleared > 0) html += '<button class="cta sub" id="reviewBtn">🔁 復習モード' + (missedN ? '（苦手 ' + missedN + '問）' : '') + '</button>';
    for (const ch of CHAPTERS) {
      html += '<h2 class="chap">' + ch.emoji + ' ' + esc(ch.name) + '</h2><div class="stages">';
      STAGES.forEach((s, i) => {
        if (s.ch !== ch.id) return;
        const un = isUnlocked(i), st = S.stars[s.id] || 0;
        html += '<button class="stage' + (un ? '' : ' locked') + (st ? ' done' : '') + '" data-go="' + i + '"' + (un ? '' : ' disabled') + '>' +
          '<span class="num">' + (un ? (st ? '✓' : i + 1) : '🔒') + '</span>' +
          '<span class="txt"><b>' + esc(s.title) + '</b><small>' + esc(s.sub) + '</small></span>' +
          '<span class="st">' + (s.exam && S.examBest !== null ? S.examBest + '点' : '★'.repeat(st) + '<i>' + '★'.repeat(3 - st) + '</i>') + '</span></button>';
      });
      html += '</div>';
    }
    html += '<p class="foot">記録はこのiPhoneの中に保存されます．ホーム画面のアイコンから開いてください．</p>';
    screen(html);
    $app.querySelectorAll('[data-go]').forEach((b) => b.addEventListener('click', () => lessonScreen(+b.dataset.go)));
    document.getElementById('menuBtn').addEventListener('click', settings);
    const rb = document.getElementById('reviewBtn'); if (rb) rb.addEventListener('click', review);
  }

  // ---------- レッスン ----------
  function lessonScreen(i) {
    const s = STAGES[i];
    let k = 0;
    function draw() {
      const L = s.lesson[k];
      let html = '<header class="top"><button class="icon" id="back">✕</button><div class="ttl">' + esc(s.title) + ' の説明</div><div class="pg">' + (k + 1) + '/' + s.lesson.length + '</div></header>';
      html += '<div class="card lesson"><h3>' + esc(L.h) + '</h3><div class="body">' + L.p + '</div>';
      if (L.g) html += gridHTML(L.g);
      if (L.f) html += '<div class="formula"><span>fx</span>' + esc(L.f) + (L.r ? '<b>→ ' + esc(L.r) + '</b>' : '') + '</div>';
      html += '</div><div class="row">';
      html += k > 0 ? '<button class="btn ghost" id="prev">← 戻る</button>' : '<button class="btn ghost" id="skip">問題へ ≫</button>';
      html += k < s.lesson.length - 1 ? '<button class="btn" id="next">次へ →</button>' : '<button class="btn go" id="start">' + (s.exam ? '試験スタート' : '問題にチャレンジ！') + '</button>';
      html += '</div>';
      screen(html);
      document.getElementById('back').onclick = home;
      const p = document.getElementById('prev'); if (p) p.onclick = () => { k--; draw(); };
      const n = document.getElementById('next'); if (n) n.onclick = () => { k++; draw(); };
      const sk = document.getElementById('skip'); if (sk) sk.onclick = () => quiz({ stage: s, qs: s.qs.slice() });
      const st = document.getElementById('start'); if (st) st.onclick = () => quiz({ stage: s, qs: s.qs.slice() });
    }
    draw();
  }

  // ---------- クイズ ----------
  function quiz(run) {
    const exam = !!(run.stage && run.stage.exam);
    const queue = run.qs.slice();
    const total = queue.length;
    let solved = 0, combo = 0, xpGain = 0;
    const wrongOnce = new Set();
    const results = [];
    const title = run.stage ? run.stage.title : '復習モード';

    function next() {
      if (!queue.length) return finish();
      ask(queue.shift());
    }

    function ask(q) {
      let html = '<header class="top"><button class="icon" id="back">✕</button><div class="prog"><i style="width:' + (solved / total * 100) + '%"></i></div><div class="pg">' + (combo >= 2 ? '🔥' + combo : '') + '</div></header>';
      html += '<div class="card q"><div class="qtag">' + ({ c: '選択', f: '数式を入力', o: '並べ替え' }[q.t]) + (wrongOnce.has(q.id) && !exam ? '・再挑戦' : '') + '</div><div class="qtext">' + q.q + '</div>';
      if (q.t === 'f') html += gridHTML(q.g, { target: q.target, fill: q.fill, tap: true });
      else if (q.g) html += gridHTML(q.g);
      html += '</div><div id="ans"></div><div id="fb"></div>';
      screen(html);
      document.getElementById('back').onclick = () => { if (confirm('やめてホームに戻りますか？（このステージの途中経過は消えます）')) home(); };
      ({ c: askChoice, f: askFormula, o: askOrder })[q.t](q);
    }

    function judge(q, ok, detail) {
      S.answered++;
      touchStreak();
      if (ok) {
        solved++; combo++;
        const g = wrongOnce.has(q.id) ? 5 : 10;
        xpGain += g; S.xp += g;
        if (S.missed[q.id]) S.missed[q.id] = Math.max(0, S.missed[q.id] - 1);
      } else {
        combo = 0;
        S.missed[q.id] = (S.missed[q.id] || 0) + 1;
        if (exam) solved++;
        else queue.push(q);
        wrongOnce.add(q.id);
      }
      results.push({ q, ok });
      save();
      const fb = document.getElementById('fb');
      fb.innerHTML = '<div class="fb ' + (ok ? 'ok' : 'ng') + '"><div class="fbh">' + (ok ? pick(['正解！', 'いいね！', 'その調子！', 'すごい！']) + (combo >= 3 ? ' 🔥' + combo + '連続' : '') : (exam ? '不正解' : '不正解…あとでもう一度出ます')) + '</div>' +
        (detail ? '<div class="detail">' + detail + '</div>' : '') +
        '<div class="ex">💡 ' + q.ex + '</div><button class="btn go" id="nextQ">次へ</button></div>';
      document.getElementById('ans').querySelectorAll('button,input').forEach((b) => { b.disabled = true; });
      document.getElementById('nextQ').onclick = next;
      fb.scrollIntoView({ behavior: 'smooth', block: 'end' });
    }
    const pick = (a) => a[Math.floor(Math.random() * a.length)];

    function askChoice(q) {
      const opts = shuffle(q.o.map((t, i) => ({ t, ok: i === 0 })));
      const box = document.getElementById('ans');
      box.innerHTML = '<div class="opts">' + opts.map((o, i) => '<button class="opt" data-i="' + i + '">' + esc(o.t) + '</button>').join('') + '</div>';
      box.querySelectorAll('.opt').forEach((b) => b.addEventListener('click', () => {
        const o = opts[+b.dataset.i];
        b.classList.add(o.ok ? 'right' : 'wrong');
        if (!o.ok) box.querySelectorAll('.opt').forEach((x) => { if (opts[+x.dataset.i].ok) x.classList.add('right'); });
        judge(q, o.ok);
      }));
    }

    function askOrder(q) {
      const pool = shuffle(q.items.map((t, i) => ({ t, i })));
      const chosen = [];
      const box = document.getElementById('ans');
      function draw() {
        box.innerHTML = '<div class="olist">' + (chosen.length ? chosen.map((o, k) => '<button class="ochip on" data-k="' + k + '"><b>' + (k + 1) + '</b> ' + esc(o.t) + '</button>').join('') : '<div class="hint">下から順番にタップ（まちがえたらタップで戻す）</div>') + '</div>' +
          '<div class="opool">' + pool.map((o, k) => '<button class="ochip" data-p="' + k + '">' + esc(o.t) + '</button>').join('') + '</div>' +
          (pool.length ? '' : '<button class="btn go" id="chk">答え合わせ</button>');
        box.querySelectorAll('[data-p]').forEach((b) => b.onclick = () => { chosen.push(pool.splice(+b.dataset.p, 1)[0]); draw(); });
        box.querySelectorAll('[data-k]').forEach((b) => b.onclick = () => { pool.push(chosen.splice(+b.dataset.k, 1)[0]); draw(); });
        const c = document.getElementById('chk');
        if (c) c.onclick = () => {
          const ok = chosen.every((o, k) => o.i === k);
          judge(q, ok, ok ? '' : '正しい順番：<ol>' + q.items.map((t) => '<li>' + esc(t) + '</li>').join('') + '</ol>');
        };
      }
      draw();
    }

    function askFormula(q) {
      const box = document.getElementById('ans');
      const fns = XL.functionsUsed(q.ans);
      const chipFns = Array.from(new Set(fns.concat(run.stage && run.stage.exam ? ['SUM', 'COUNT', 'IF'] : [])));
      const syms = ['(', ')', ',', ':', '$', '"', '>=', '<=', '<>', '>', '<', '=', '&', '*', '/', '+', '-'];
      box.innerHTML = '<div class="fxbar"><span>fx</span><input id="fx" type="text" autocomplete="off" autocapitalize="characters" autocorrect="off" spellcheck="false" enterkeyhint="done" placeholder="=SUM(A1:A5)" value="="></div>' +
        '<div class="hint">表のセルをタップすると番地が入ります</div>' +
        '<div class="chips">' + chipFns.map((f) => '<button class="chip fn" data-ins="' + f + '(">' + f + '</button>').join('') +
        syms.map((s) => '<button class="chip" data-ins="' + esc(s) + '">' + esc(s) + '</button>').join('') +
        '<button class="chip" data-del="1">⌫</button><button class="chip" data-clr="1">クリア</button></div>' +
        '<div id="err" class="err"></div>' +
        '<div class="row"><button class="btn ghost" id="give">答えを見る</button><button class="btn go" id="chk">答え合わせ</button></div>';
      const fx = document.getElementById('fx');
      let caret = 1;
      const remember = () => { caret = fx.selectionStart == null ? fx.value.length : fx.selectionStart; };
      ['keyup', 'click', 'input', 'select', 'blur'].forEach((ev) => fx.addEventListener(ev, remember));
      function insert(s) {
        const v = fx.value;
        const pos = Math.min(caret, v.length);
        fx.value = v.slice(0, pos) + s + v.slice(pos);
        caret = pos + s.length;
        try { fx.setSelectionRange(caret, caret); } catch (e) { /* noop */ }
      }
      box.querySelectorAll('[data-ins]').forEach((b) => b.addEventListener('click', () => insert(b.dataset.ins)));
      box.querySelector('[data-del]').onclick = () => { const p = Math.min(caret, fx.value.length); if (p > 0) { fx.value = fx.value.slice(0, p - 1) + fx.value.slice(p); caret = p - 1; } };
      box.querySelector('[data-clr]').onclick = () => { fx.value = '='; caret = 1; };
      $app.querySelectorAll('.grid.tap td').forEach((td) => td.addEventListener('click', () => {
        insert(td.dataset.a);
        td.classList.add('flash'); setTimeout(() => td.classList.remove('flash'), 300);
      }));
      fx.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); check(); } });
      document.getElementById('chk').onclick = check;
      document.getElementById('give').onclick = () => {
        fx.value = q.ans;
        judge(q, false, '模範解答：<code>' + esc(q.ans) + '</code>');
        showResultGrid(q, q.ans);
      };

      function check() {
        let src = XL.normalizeInput(fx.value);
        if (!src.startsWith('=')) src = '=' + src;
        const err = document.getElementById('err');
        err.textContent = '';
        if (src === '=') { err.textContent = '数式を入力してください'; return; }
        try { XL.parse(src.slice(1)); } catch (e) { err.textContent = '⚠️ ' + e.message; return; }
        const up = src.toUpperCase();
        const used = XL.functionsUsed(src.slice(1));
        const unknown = used.filter((f) => !['SUM', 'MAX', 'MIN', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'SUMIF', 'AVERAGE', 'AVERAGEA', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'ABS', 'TODAY', 'IF', 'AND', 'OR', 'NOT', 'VLOOKUP'].includes(f));
        if (unknown.length) { err.textContent = '⚠️ 関数名「' + unknown[0] + '」はスペルが違うかもしれません'; return; }
        const rows = rowsOf(q);
        const lines = [];
        let ok = true;
        for (const r of rows) {
          const mine = XL.evaluate(XL.shiftFormula(src.slice(1), r.d, 0), XL.makeSheet(q.g.cells, q.g.names));
          const want = XL.evaluate(XL.shiftFormula(q.ans.slice(1), r.d, 0), XL.makeSheet(q.g.cells, q.g.names));
          const same = XL.sameValue(mine, want);
          if (!same) ok = false;
          lines.push('<tr class="' + (same ? 'y' : 'n') + '"><td>' + r.a + '</td><td>' + esc(XL.format(mine, q.fmt) || '（空）') + '</td><td>' + esc(XL.format(want, q.fmt) || '（空）') + '</td><td>' + (same ? '○' : '×') + '</td></tr>');
        }
        const FNS = ['SUM', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'SUMIF', 'AVERAGE', 'AVERAGEA', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TODAY', 'IF', 'AND', 'OR', 'VLOOKUP'];
        const lack = (q.must || []).filter((m) => !(FNS.includes(m) ? used.includes(m) : up.includes(m)));
        let detail = '<table class="cmp"><tr><th>セル</th><th>あなた</th><th>正解</th><th></th></tr>' + lines.join('') + '</table>';
        if (ok && lack.length) {
          ok = false;
          detail += '<div class="warn">値は合っていますが，' + lack.map((m) => m === '$' ? '「$」（絶対参照）' : '「' + esc(m) + '」').join('と') + 'を使って書いてみよう．</div>';
        }
        if (!ok && rows.length > 1 && !src.includes('$') && q.ans.includes('$')) detail += '<div class="warn">下にコピーするとずれる部分はありませんか？ $ で固定してみよう．</div>';
        detail += '<div>あなたの数式：<code>' + esc(src) + '</code>' + (ok && src.toUpperCase().replace(/\s/g, '') !== q.ans.toUpperCase() ? '<br>模範解答：<code>' + esc(q.ans) + '</code>' : '') + (ok ? '' : '<br>模範解答：<code>' + esc(q.ans) + '</code>') + '</div>';
        judge(q, ok, detail);
        showResultGrid(q, src);
      }
    }

    function rowsOf(q) {
      const t = XL.parseAddr(q.target);
      const end = q.fill ? XL.parseAddr(q.fill).row : t.row;
      const out = [];
      for (let r = t.row; r <= end; r++) out.push({ d: r - t.row, a: XL.addr(t.col, r) });
      return out;
    }
    function showResultGrid(q, src) {
      const over = {};
      for (const r of rowsOf(q)) over[r.a] = '=' + XL.shiftFormula(src.slice(1), r.d, 0);
      const wrap = $app.querySelector('.card.q .gridwrap');
      if (!wrap) return;
      const tmp = document.createElement('div');
      tmp.innerHTML = gridHTML(q.g, { target: q.target, fill: q.fill, override: over, fmt: q.fmt });
      wrap.replaceWith(tmp.querySelector('.gridwrap'));
    }

    function finish() {
      const mist = wrongOnce.size;
      let html = '<div class="finish">';
      if (exam) {
        const correct = results.filter((r) => r.ok).length;
        const score = Math.round(correct / total * 100);
        const pass = score >= 80;
        S.examBest = Math.max(S.examBest || 0, score);
        if (pass) S.stars[run.stage.id] = Math.max(S.stars[run.stage.id] || 0, score === 100 ? 3 : score >= 90 ? 2 : 1);
        save();
        html += '<div class="medal">' + (pass ? '🏆' : '📘') + '</div><h2>' + score + ' 点</h2><p>' + (pass ? '合格です！本番もこの調子で．' : 'あと少し！まちがえた問題のステージを復習しよう．') + '</p>';
        html += '<div class="review">' + results.map((r, i) => '<div class="rv ' + (r.ok ? 'y' : 'n') + '"><b>' + (i + 1) + '. ' + (r.ok ? '○' : '×') + '</b> ' + r.q.q.replace(/<[^>]+>/g, '') + (r.ok ? '' : '<div class="ex">💡 ' + r.q.ex + '</div>') + '</div>').join('') + '</div>';
      } else {
        const stars = mist === 0 ? 3 : mist <= 2 ? 2 : 1;
        if (run.stage) { S.stars[run.stage.id] = Math.max(S.stars[run.stage.id] || 0, stars); }
        save();
        html += '<div class="medal">' + (stars === 3 ? '🎉' : '✨') + '</div><h2>' + esc(title) + ' クリア！</h2>' +
          '<div class="bigstars">' + '★'.repeat(stars) + '<i>' + '★'.repeat(3 - stars) + '</i></div>' +
          '<p>+' + xpGain + ' XP　／　まちがい ' + mist + ' 問</p>' +
          (stars < 3 && run.stage ? '<p class="small">ノーミスで ★3 ！もう一度挑戦できます．</p>' : '');
      }
      html += '</div><div class="row">' + (run.stage ? '<button class="btn ghost" id="again">もう一度</button>' : '') + '<button class="btn go" id="home">ホームへ</button></div>';
      screen(html);
      if (!exam || results.filter((r) => r.ok).length / total >= 0.8) confetti();
      document.getElementById('home').onclick = home;
      const a = document.getElementById('again'); if (a) a.onclick = () => quiz({ stage: run.stage, qs: run.stage.qs.slice() });
    }

    next();
  }

  // ---------- 復習 ----------
  function review() {
    const pool = [];
    STAGES.forEach((s) => { if (isCleared(s.id) && !s.exam) s.qs.forEach((q) => pool.push(q)); });
    const weighted = shuffle(pool).sort((a, b) => (S.missed[b.id] || 0) - (S.missed[a.id] || 0));
    const qs = shuffle(weighted.slice(0, 8));
    if (!qs.length) { home(); return; }
    // 復習では数式チップ用に元のステージを引けるようにする
    quiz({ stage: null, qs });
  }

  // ---------- 設定 ----------
  function settings() {
    screen('<header class="top"><button class="icon" id="back">←</button><div class="ttl">設定・使い方</div><div></div></header>' +
      '<div class="card"><h3>📱 ホーム画面に追加</h3><div class="body">Safari で開いて，下の<b>共有ボタン</b>（□に↑）→<b>「ホーム画面に追加」</b>．<br>以後はホーム画面のアイコンから開くと，アプリのように全画面で使えて，電波が弱くても動きます．<br><b>記録はアイコンから開いたときの分が保存されます</b>（Safariで直接開いたときとは別）．</div></div>' +
      '<div class="card"><h3>💾 記録のバックアップ</h3><div class="body">機種変更や念のための保存用．コピーした文字をメモ等に貼っておけば，「読み込む」で戻せます．</div>' +
      '<div class="row"><button class="btn ghost" id="exp">記録をコピー</button><button class="btn ghost" id="imp">記録を読み込む</button></div></div>' +
      '<div class="card"><h3>🔓 全ステージを開放</h3><div class="body">順番を飛ばして好きなステージに挑戦できます．</div><label class="sw"><input type="checkbox" id="ua"' + (S.unlockAll ? ' checked' : '') + '> すべて開放する</label></div>' +
      '<div class="card"><h3>📊 これまでの記録</h3><div class="body">解いた問題：' + S.answered + ' 問<br>XP：' + S.xp + '<br>模擬試験の最高点：' + (S.examBest === null ? 'まだ' : S.examBest + ' 点') + '</div></div>' +
      '<div class="card"><h3>🗑️ リセット</h3><button class="btn danger" id="reset">記録をすべて消す</button></div>');
    document.getElementById('back').onclick = home;
    document.getElementById('ua').onchange = (e) => { S.unlockAll = e.target.checked; save(); };
    document.getElementById('exp').onclick = async () => {
      const code = btoa(unescape(encodeURIComponent(JSON.stringify(S))));
      try { await navigator.clipboard.writeText(code); alert('コピーしました．メモなどに貼り付けて保管してください．'); }
      catch (e) { prompt('この文字を長押ししてコピーしてください', code); }
    };
    document.getElementById('imp').onclick = () => {
      const code = prompt('バックアップした文字を貼り付けてください');
      if (!code) return;
      try { const obj = JSON.parse(decodeURIComponent(escape(atob(code.trim())))); S = Object.assign(blank(), obj); save(); alert('読み込みました'); home(); }
      catch (e) { alert('読み込めませんでした．文字が途中で切れていないか確認してください．'); }
    };
    document.getElementById('reset').onclick = () => { if (confirm('本当にすべての記録を消しますか？元に戻せません．')) { S = blank(); save(); home(); } };
  }

  // ---------- 演出 ----------
  function confetti() {
    const box = h('<div class="confetti"></div>');
    const em = ['🎉', '⭐', '✨', '📗', '💚'];
    for (let i = 0; i < 24; i++) {
      const s = document.createElement('span');
      s.textContent = em[i % em.length];
      s.style.left = Math.random() * 100 + '%';
      s.style.animationDelay = Math.random() * 0.6 + 's';
      s.style.fontSize = 16 + Math.random() * 16 + 'px';
      box.appendChild(s);
    }
    document.body.appendChild(box);
    setTimeout(() => box.remove(), 2600);
  }

  home();
  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
