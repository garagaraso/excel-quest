/* Excelクエスト 数式エンジン
 * Excel の数式をブラウザ内で計算する小さな実装．
 * 対応：四則演算・比較・& 連結・セル参照（$ 付き）・範囲・名前付き範囲
 * 関数：SUM MAX MIN COUNT COUNTA COUNTBLANK COUNTIF SUMIF AVERAGE AVERAGEA
 *       ROUND ROUNDUP ROUNDDOWN INT ABS TODAY IF AND OR NOT VLOOKUP
 */
(function (global) {
  'use strict';

  // ---------- 基本 ----------
  const ERR = (code) => ({ err: code });
  const isErr = (v) => v && typeof v === 'object' && 'err' in v;
  const isRange = (v) => v && typeof v === 'object' && v.type === 'range';

  function colToNum(s) {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function parseAddr(a) {
    const m = /^(\$?)([A-Z]{1,2})(\$?)(\d+)$/i.exec(a);
    if (!m) return null;
    return { col: colToNum(m[2]), row: parseInt(m[4], 10), absCol: !!m[1], absRow: !!m[3] };
  }
  const addr = (col, row) => numToCol(col) + row;

  function todaySerial() {
    const d = new Date();
    const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    return Math.round(utc / 86400000) + 25569;
  }
  function serialToDateStr(s) {
    const ms = (s - 25569) * 86400000;
    const d = new Date(ms);
    return d.getUTCFullYear() + '/' + (d.getUTCMonth() + 1) + '/' + d.getUTCDate();
  }

  // 全角→半角（iPhone の日本語入力対策）
  function normalizeInput(s) {
    return s
      .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[“”″]/g, '"')
      .replace(/[、，]/g, ',')
      .replace(/　/g, ' ')
      .replace(/＄/g, '$')
      .trim();
  }

  // ---------- 字句解析 ----------
  function tokenize(src) {
    const toks = [];
    let i = 0;
    const s = src;
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"') {
        let j = i + 1, str = '';
        for (;;) {
          if (j >= s.length) throw new SyntaxError('文字列の " が閉じていません');
          if (s[j] === '"') {
            if (s[j + 1] === '"') { str += '"'; j += 2; continue; }
            break;
          }
          str += s[j++];
        }
        toks.push({ t: 'str', v: str, raw: s.slice(i, j + 1) });
        i = j + 1; continue;
      }
      let m;
      const rest = s.slice(i);
      if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(rest)) || (m = /^\.\d+/.exec(rest))) {
        toks.push({ t: 'num', v: parseFloat(m[0]), raw: m[0] }); i += m[0].length; continue;
      }
      if ((m = /^(\$?[A-Za-z]{1,2}\$?\d+)(:(\$?[A-Za-z]{1,2}\$?\d+))?(?![A-Za-z0-9_(])/.exec(rest))) {
        toks.push({ t: 'ref', v: m[0].toUpperCase(), raw: m[0] }); i += m[0].length; continue;
      }
      if ((m = /^(\$?[A-Za-z]{1,2}):(\$?[A-Za-z]{1,2})(?![A-Za-z0-9_(])/.exec(rest))) {
        toks.push({ t: 'colref', v: m[0].toUpperCase(), raw: m[0] }); i += m[0].length; continue;
      }
      if ((m = /^[A-Za-z_぀-ヿ一-鿿][A-Za-z0-9_.぀-ヿ一-鿿]*/.exec(rest))) {
        const word = m[0];
        let j = i + word.length;
        while (j < s.length && s[j] === ' ') j++;
        if (s[j] === '(') toks.push({ t: 'func', v: word.toUpperCase(), raw: word });
        else if (/^(TRUE|FALSE)$/i.test(word)) toks.push({ t: 'bool', v: word.toUpperCase() === 'TRUE', raw: word });
        else toks.push({ t: 'name', v: word, raw: word });
        i += word.length; continue;
      }
      if ((m = /^(<=|>=|<>)/.exec(rest))) { toks.push({ t: 'op', v: m[0], raw: m[0] }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(c)) { toks.push({ t: 'op', v: c, raw: c }); i++; continue; }
      if (c === '(' || c === ')' || c === ',') { toks.push({ t: c, v: c, raw: c }); i++; continue; }
      if (c === ';') throw new SyntaxError('区切りは「;」ではなく「,」（カンマ）です');
      throw new SyntaxError('使えない文字があります：「' + c + '」');
    }
    return toks;
  }

  // ---------- 構文解析 ----------
  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];
    const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;

    function parseCompare() {
      let left = parseConcat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
        const op = next().v;
        left = { k: 'bin', op, a: left, b: parseConcat() };
      }
      return left;
    }
    function parseConcat() {
      let left = parseAdd();
      while (isOp('&')) { next(); left = { k: 'bin', op: '&', a: left, b: parseAdd() }; }
      return left;
    }
    function parseAdd() {
      let left = parseMul();
      while (isOp('+') || isOp('-')) { const op = next().v; left = { k: 'bin', op, a: left, b: parseMul() }; }
      return left;
    }
    function parseMul() {
      let left = parsePow();
      while (isOp('*') || isOp('/')) { const op = next().v; left = { k: 'bin', op, a: left, b: parsePow() }; }
      return left;
    }
    function parsePow() {
      let left = parseUnary();
      while (isOp('^')) { next(); left = { k: 'bin', op: '^', a: left, b: parseUnary() }; }
      return left;
    }
    function parseUnary() {
      if (isOp('-')) { next(); return { k: 'neg', a: parseUnary() }; }
      if (isOp('+')) { next(); return parseUnary(); }
      let v = parsePrimary();
      while (isOp('%')) { next(); v = { k: 'bin', op: '/', a: v, b: { k: 'lit', v: 100 } }; }
      return v;
    }
    function parsePrimary() {
      const tk = next();
      if (!tk) throw new SyntaxError('数式が途中で終わっています');
      switch (tk.t) {
        case 'num': return { k: 'lit', v: tk.v };
        case 'str': return { k: 'lit', v: tk.v };
        case 'bool': return { k: 'lit', v: tk.v };
        case 'ref': return { k: 'ref', v: tk.v };
        case 'colref': return { k: 'colref', v: tk.v };
        case 'name': return { k: 'name', v: tk.v };
        case '(': {
          const e = parseCompare();
          if (!peek() || peek().t !== ')') throw new SyntaxError('カッコ「)」が足りません');
          next();
          return e;
        }
        case 'func': {
          next(); // (
          const args = [];
          if (peek() && peek().t === ')') { next(); return { k: 'call', f: tk.v, args }; }
          for (;;) {
            if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ k: 'empty' });
            else args.push(parseCompare());
            const t = next();
            if (!t) throw new SyntaxError('カッコ「)」が足りません');
            if (t.t === ')') break;
            if (t.t !== ',') throw new SyntaxError('引数の区切りは「,」です');
          }
          return { k: 'call', f: tk.v, args };
        }
        default:
          throw new SyntaxError('「' + tk.raw + '」の位置がおかしいです');
      }
    }
    const ast = parseCompare();
    if (p < toks.length) {
      const t = toks[p];
      if (t.t === ')') throw new SyntaxError('カッコ「)」が多すぎます');
      throw new SyntaxError('「' + t.raw + '」の前に記号や「,」が抜けていませんか？');
    }
    return ast;
  }

  // ---------- 評価 ----------
  function makeSheet(cells, names) {
    // cells: { A1: 値 or "=数式" or {d: 日数} }
    const sheet = { cells: cells || {}, names: names || {}, cache: {}, busy: {} };
    let maxRow = 1;
    for (const k of Object.keys(sheet.cells)) { const a = parseAddr(k); if (a && a.row > maxRow) maxRow = a.row; }
    sheet.maxRow = maxRow;
    return sheet;
  }

  function cellValue(sheet, a) {
    const key = a.toUpperCase().replace(/\$/g, '');
    if (key in sheet.cache) return sheet.cache[key];
    let raw = sheet.cells[key];
    let v;
    if (raw === undefined || raw === null || raw === '') v = null;
    else if (typeof raw === 'object' && 'd' in raw) v = todaySerial() + raw.d;
    else if (typeof raw === 'object' && 'v' in raw) v = raw.v;
    else if (typeof raw === 'string' && raw.startsWith('=')) {
      if (sheet.busy[key]) return ERR('#循環');
      sheet.busy[key] = true;
      try { v = evaluate(raw.slice(1), sheet); } catch (e) { v = ERR('#NAME?'); }
      sheet.busy[key] = false;
      if (isRange(v)) v = ERR('#VALUE!');
    } else v = raw;
    sheet.cache[key] = v;
    return v;
  }

  function rangeOf(sheet, a1, a2) {
    const p = parseAddr(a1), q = parseAddr(a2);
    const r1 = Math.min(p.row, q.row), r2 = Math.max(p.row, q.row);
    const c1 = Math.min(p.col, q.col), c2 = Math.max(p.col, q.col);
    const rows = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) row.push(cellValue(sheet, addr(c, r)));
      rows.push(row);
    }
    return { type: 'range', rows, r1, c1 };
  }

  function resolveRefText(sheet, text) {
    const parts = text.replace(/\$/g, '').split(':');
    if (parts.length === 1) return cellValue(sheet, parts[0]);
    return rangeOf(sheet, parts[0], parts[1]);
  }

  function flat(v) {
    if (isRange(v)) return v.rows.flat();
    return [v];
  }
  function scalar(v) {
    if (isRange(v)) {
      if (v.rows.length === 1 && v.rows[0].length === 1) return v.rows[0][0];
      return ERR('#VALUE!');
    }
    return v;
  }
  function toNum(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (v === null) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const t = v.trim();
      if (t !== '' && !isNaN(Number(t))) return Number(t);
      return ERR('#VALUE!');
    }
    return ERR('#VALUE!');
  }
  function toStr(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return String(clean(v));
    return v;
  }
  function toBool(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (v === null) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
      if (/^true$/i.test(v)) return true;
      if (/^false$/i.test(v)) return false;
      return ERR('#VALUE!');
    }
    return ERR('#VALUE!');
  }
  const clean = (x) => parseFloat(Number(x).toPrecision(15));

  function typeRank(v) {
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 0;
  }
  function compare(a, b) {
    if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = typeRank(a), rb = typeRank(b);
    if (ra !== rb) return ra - rb;
    if (typeof a === 'string') {
      const x = a.toLowerCase(), y = b.toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    }
    if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
    const d = clean(a) - clean(b);
    return d < 0 ? -1 : d > 0 ? 1 : 0;
  }

  // 条件（COUNTIF / SUMIF）
  function makeCriteria(crit) {
    crit = scalar(crit);
    if (typeof crit === 'number' || typeof crit === 'boolean') {
      return (v) => v !== null && typeof v === typeof crit && compare(v, crit) === 0;
    }
    if (crit === null) crit = '';
    let s = String(crit);
    let op = '=';
    const m = /^(<=|>=|<>|=|<|>)/.exec(s);
    if (m) { op = m[1]; s = s.slice(m[1].length); }
    const isNum = s.trim() !== '' && !isNaN(Number(s));
    if (isNum) {
      const n = Number(s);
      return (v) => {
        if (typeof v !== 'number') return op === '<>';
        const c = compare(v, n);
        return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
      };
    }
    if (/^(true|false)$/i.test(s)) {
      const b = /^true$/i.test(s);
      return (v) => {
        const eq = typeof v === 'boolean' && v === b;
        return op === '<>' ? !eq : op === '=' ? eq : false;
      };
    }
    if (op === '=' || op === '<>') {
      if (s === '') {
        return (v) => {
          const blank = v === null || v === '';
          return op === '=' ? blank : !blank;
        };
      }
      const re = new RegExp('^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      return (v) => {
        const hit = typeof v === 'string' && re.test(v);
        return op === '=' ? hit : !hit;
      };
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const c = compare(v, s);
      return { '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
    };
  }

  function roundCore(x, n, mode) {
    x = clean(x);
    const f = Math.pow(10, n);
    const sgn = x < 0 ? -1 : 1;
    const ax = Math.abs(x);
    let scaled = clean(ax * f);
    let r;
    if (mode === 'round') r = Math.floor(scaled + 0.5);
    else if (mode === 'up') r = Math.ceil(scaled);
    else r = Math.floor(scaled);
    return clean(sgn * r / f);
  }

  const FN = {
    SUM(args) {
      let s = 0;
      for (const a of args) {
        if (isRange(a)) { for (const v of flat(a)) { if (isErr(v)) return v; if (typeof v === 'number') s += v; } }
        else { const n = toNum(a); if (isErr(n)) return n; s += n; }
      }
      return clean(s);
    },
    MAX(args) { const ns = nums(args); if (isErr(ns)) return ns; return ns.length ? Math.max(...ns) : 0; },
    MIN(args) { const ns = nums(args); if (isErr(ns)) return ns; return ns.length ? Math.min(...ns) : 0; },
    COUNT(args) {
      let c = 0;
      for (const a of args) {
        if (isRange(a)) { for (const v of flat(a)) if (typeof v === 'number') c++; }
        else if (!isErr(toNum(a)) && a !== null) c++;
      }
      return c;
    },
    COUNTA(args) {
      let c = 0;
      for (const a of args) {
        if (isRange(a)) { for (const v of flat(a)) if (v !== null) c++; }
        else c++;
      }
      return c;
    },
    COUNTBLANK(args) {
      if (args.length !== 1) return ERR('#引数');
      let c = 0;
      for (const v of flat(args[0])) if (v === null || v === '') c++;
      return c;
    },
    AVERAGE(args) {
      const ns = nums(args); if (isErr(ns)) return ns;
      if (!ns.length) return ERR('#DIV/0!');
      return clean(ns.reduce((a, b) => a + b, 0) / ns.length);
    },
    AVERAGEA(args) {
      const ns = [];
      for (const a of args) {
        if (isRange(a)) {
          for (const v of flat(a)) {
            if (isErr(v)) return v;
            if (v === null) continue;
            if (typeof v === 'number') ns.push(v);
            else if (typeof v === 'boolean') ns.push(v ? 1 : 0);
            else ns.push(0);
          }
        } else { const n = toNum(a); if (isErr(n)) return n; ns.push(n); }
      }
      if (!ns.length) return ERR('#DIV/0!');
      return clean(ns.reduce((a, b) => a + b, 0) / ns.length);
    },
    ROUND(args) { return roundFn(args, 'round'); },
    ROUNDUP(args) { return roundFn(args, 'up'); },
    ROUNDDOWN(args) { return roundFn(args, 'down'); },
    INT(args) { const n = toNum(args[0]); if (isErr(n)) return n; return Math.floor(clean(n)); },
    ABS(args) { const n = toNum(args[0]); if (isErr(n)) return n; return Math.abs(n); },
    TODAY(args) { if (args.length) return ERR('#引数'); return todaySerial(); },
    IF(args) {
      if (args.length < 2 || args.length > 3) return ERR('#引数');
      const c = toBool(args[0]);
      if (isErr(c)) return c;
      const pick = c ? args[1] : (args.length === 3 ? args[2] : false);
      return pick === undefined ? false : scalar(pick);
    },
    AND(args) {
      if (!args.length) return ERR('#引数');
      let r = true;
      for (const a of args) for (const v of flat(a)) { const b = toBool(v); if (isErr(b)) return b; r = r && b; }
      return r;
    },
    OR(args) {
      if (!args.length) return ERR('#引数');
      let r = false;
      for (const a of args) for (const v of flat(a)) { const b = toBool(v); if (isErr(b)) return b; r = r || b; }
      return r;
    },
    NOT(args) { const b = toBool(args[0]); if (isErr(b)) return b; return !b; },
    COUNTIF(args) {
      if (args.length !== 2) return ERR('#引数');
      if (!isRange(args[0])) return ERR('#VALUE!');
      const f = makeCriteria(args[1]);
      return flat(args[0]).filter((v) => f(v)).length;
    },
    SUMIF(args) {
      if (args.length < 2 || args.length > 3) return ERR('#引数');
      if (!isRange(args[0])) return ERR('#VALUE!');
      const f = makeCriteria(args[1]);
      const crit = args[0].rows;
      const sumR = args.length === 3 ? args[2] : args[0];
      if (!isRange(sumR)) return ERR('#VALUE!');
      let s = 0;
      for (let r = 0; r < crit.length; r++) for (let c = 0; c < crit[r].length; c++) {
        if (f(crit[r][c])) {
          const v = sumR.rows[r] ? sumR.rows[r][c] : null;
          if (typeof v === 'number') s += v;
        }
      }
      return clean(s);
    },
    VLOOKUP(args) {
      if (args.length < 3 || args.length > 4) return ERR('#引数');
      const key = scalar(args[0]);
      if (isErr(key)) return key;
      const tbl = args[1];
      if (!isRange(tbl)) return ERR('#VALUE!');
      const col = toNum(args[2]);
      if (isErr(col)) return col;
      const ci = Math.floor(col);
      if (ci < 1) return ERR('#VALUE!');
      if (ci > tbl.rows[0].length) return ERR('#REF!');
      let approx = true;
      if (args.length === 4) { const b = toBool(args[3]); if (isErr(b)) return b; approx = b; }
      if (!approx) {
        for (const row of tbl.rows) {
          const v = row[0];
          if (v === null) continue;
          if (typeof key === 'string' && typeof v === 'string') {
            if (makeCriteria('=' + key)(v)) return row[ci - 1] === null ? 0 : row[ci - 1];
          } else if (typeof v === typeof key && compare(v, key) === 0) return row[ci - 1] === null ? 0 : row[ci - 1];
        }
        return ERR('#N/A');
      }
      let hit = -1;
      for (let i = 0; i < tbl.rows.length; i++) {
        const v = tbl.rows[i][0];
        if (v === null || typeof v !== typeof key) continue;
        if (compare(v, key) <= 0) hit = i; else break;
      }
      if (hit < 0) return ERR('#N/A');
      const out = tbl.rows[hit][ci - 1];
      return out === null ? 0 : out;
    },
  };
  function nums(args) {
    const out = [];
    for (const a of args) {
      if (isRange(a)) { for (const v of flat(a)) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); } }
      else { const n = toNum(a); if (isErr(n)) return n; out.push(n); }
    }
    return out;
  }
  function roundFn(args, mode) {
    if (args.length !== 2) return ERR('#引数');
    const x = toNum(args[0]); if (isErr(x)) return x;
    const n = toNum(args[1]); if (isErr(n)) return n;
    return roundCore(x, Math.trunc(n), mode);
  }

  function ev(node, sheet) {
    switch (node.k) {
      case 'lit': return node.v;
      case 'empty': return null;
      case 'ref': return resolveRefText(sheet, node.v);
      case 'colref': {
        const [a, b] = node.v.replace(/\$/g, '').split(':');
        return rangeOf(sheet, a + '1', b + sheet.maxRow);
      }
      case 'name': {
        const key = Object.keys(sheet.names).find((k) => k.toLowerCase() === node.v.toLowerCase());
        if (!key) return ERR('#NAME?');
        return resolveRefText(sheet, sheet.names[key]);
      }
      case 'neg': { const n = toNum(ev(node.a, sheet)); return isErr(n) ? n : -n; }
      case 'bin': {
        const a = ev(node.a, sheet), b = ev(node.b, sheet);
        if (node.op === '&') {
          const x = toStr(a), y = toStr(b);
          if (isErr(x)) return x; if (isErr(y)) return y;
          return x + y;
        }
        if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
          const x = scalar(a), y = scalar(b);
          if (isErr(x)) return x; if (isErr(y)) return y;
          const c = compare(x, y);
          return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[node.op];
        }
        const x = toNum(a), y = toNum(b);
        if (isErr(x)) return x; if (isErr(y)) return y;
        switch (node.op) {
          case '+': return clean(x + y);
          case '-': return clean(x - y);
          case '*': return clean(x * y);
          case '/': return y === 0 ? ERR('#DIV/0!') : clean(x / y);
          case '^': return clean(Math.pow(x, y));
        }
        return ERR('#VALUE!');
      }
      case 'call': {
        const fn = FN[node.f];
        if (!fn) return ERR('#NAME?');
        const args = node.args.map((a) => ev(a, sheet));
        // 範囲以外の引数にエラーがあれば伝える（IF は選ばれた側だけ）
        if (node.f !== 'IF') for (const a of args) if (isErr(a)) return a;
        return fn(args);
      }
    }
    return ERR('#VALUE!');
  }

  function evaluate(src, sheet) {
    const ast = parse(src);
    return ev(ast, sheet);
  }

  // 使われている関数名
  function functionsUsed(src) {
    try { return tokenize(src).filter((t) => t.t === 'func').map((t) => t.v); } catch (e) { return []; }
  }

  // 数式をコピーしたときのずらし（相対参照だけ動く）
  function shiftFormula(src, dRow, dCol) {
    const toks = tokenize(src);
    return toks.map((t) => {
      if (t.t === 'ref') {
        return t.v.split(':').map((part) => {
          const a = parseAddr(part);
          const c = a.absCol ? a.col : a.col + dCol;
          const r = a.absRow ? a.row : a.row + dRow;
          if (c < 1 || r < 1) return '#REF!';
          return (a.absCol ? '$' : '') + numToCol(c) + (a.absRow ? '$' : '') + r;
        }).join(':');
      }
      if (t.t === 'str') return t.raw;
      return t.raw;
    }).join('');
  }

  function format(v, fmt) {
    if (v === null || v === undefined) return '';
    if (isErr(v)) return v.err;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      if (fmt === 'date') return serialToDateStr(v);
      const c = clean(v);
      if (Math.abs(c) >= 1e11) return String(c);
      return String(parseFloat(c.toFixed(10)));
    }
    return String(v);
  }

  function sameValue(a, b) {
    if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.err === b.err;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
    if (a === null) a = '';
    if (b === null) b = '';
    return a === b;
  }

  global.XL = { parse, tokenize, evaluate, makeSheet, cellValue, shiftFormula, functionsUsed, format, sameValue,
    normalizeInput, parseAddr, colToNum, numToCol, addr, todaySerial, isErr };
})(typeof window !== 'undefined' ? window : globalThis);
