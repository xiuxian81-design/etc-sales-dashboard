/* brief_export.js — [v71] 전략브리핑 자료(Word 2쪽) · 목표 근거 대상 리스트(엑셀) 생성기. 폼(form.html)·성과보드(index.html) 공용.
   ★ 데이터 내장 없음 — 입력은 GAS doGet(mode=crm&gb=1&dir=지부장) 응답의 그 지부장 카드(d)와 meta뿐. 파일은 누르는 자리에서 만든다.
   형식 정본 = _workspace\v71_전략브리핑\견본\(전략브리핑_견본·목표근거_대상리스트_견본)_김종훈_2026-11 + build_sample.py. 1절 수치 정본 = goalBasis.kpi2.
   라이브러리(docx·ExcelJS)는 버튼을 누를 때만 **같은 출처**(Pages에 함께 올린 vendor\ 파일)에서 불러온다(QA🟡-4 — 외부 CDN 의존 없음). 공개 라이브러리 원본 그대로(데이터 아님):
     vendor/docx.iife.js   = docx 9.5.1 dist/index.iife.js (MIT) sha384-HWC8sFPgg/WYrr43jIR6e0YIKgPkS3VLdteGgh9MFlUet3xdAVEM8YS69xP4Puse
     vendor/exceljs.min.js = exceljs 4.4.0 dist/exceljs.min.js (MIT) sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz
     (11-15 jsdelivr·cdnjs 원본과 바이트 동일 확인). node에서도 같은 코드로 생성(라이브러리를 인자로 넘김) — 시험 = _workspace\v71_전략브리핑\gen_brief_node.js
   API: BriefExport.model(d, meta, opt) → M · docxDoc(M, docx) → Document · xlsxBook(M, ExcelJS) → Workbook · fileName('docx'|'xlsx', M)
        BriefExport.basisRows72(M) → [v72] 4절 두 기간 행 · BriefExport.coreRows(kpi2, dir) → 1절 두 줄 표 행(성과보드 표도 같은 함수) · BriefExport.download(kind, d, meta, opt) → Promise(브라우저) */
(function (root) {
  'use strict';
  var LIB = {
    docx: { url: 'vendor/docx.iife.js', global: 'docx', name: 'Word 생성 도구(docx 9.5.1)' },              // 페이지와 같은 출처(상대 경로)
    exceljs: { url: 'vendor/exceljs.min.js', global: 'ExcelJS', name: '엑셀 생성 도구(ExcelJS 4.4.0)' }
  };
  var WFONT = 'ChosunilboNM', XFONT = 'NanumSquare';
  var GRAY = 'EDEDED', HEAD = 'D6D6D6', BROWN = '8B4513', HINT = 'A0A0A0', NAVY = '1F3864', DIM = '595959';

  /* ── 숫자 표기(견본 build_sample.py amt·eok·pct와 같은 규칙) ── */
  function comma(n) { var s = String(n), neg = s.charAt(0) === '-'; if (neg) s = s.slice(1); var p = s.split('.'); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); return (neg ? '-' : '') + p.join('.'); }
  function num(v) { var n = Number(v); return (v === null || v === undefined || v === '' || isNaN(n)) ? null : n; }
  function amt(won, sign) {              // 억은 소수 둘째 자리, 1억 미만은 만 정수
    won = num(won); if (won === null) return '—';
    var s = (sign && won > 0) ? '+' : (won < 0 ? '−' : ''), a = Math.abs(won);
    if (a >= 1e8) return s + comma((a / 1e8).toFixed(2)) + '억';
    return s + comma(Math.round(a / 1e4)) + '만';
  }
  function eok(won) { won = num(won); return won === null ? '—' : comma((won / 1e8).toFixed(2)); }
  function pct(v, sign) { v = num(v); if (v === null) return '—'; var s = (sign && v > 0) ? '+' : (v < 0 ? '−' : ''); return s + (Math.abs(v) * 100).toFixed(1) + '%'; }
  function man(won) { won = num(won); return won === null ? null : Math.round(won / 1e4); }   // 엑셀 만원 정수
  function mon(ym) { return ym ? (parseInt(String(ym).slice(5, 7), 10) + '월') : '-'; }
  function wonTxt(won) { won = num(won); return won === null ? '—' : ((won < 0 ? '−' : '') + comma(Math.abs(Math.round(won))) + '원'); }
  function bizNm(co) { return String(co || '').replace(/주식회사|\(주\)|㈜/g, '').trim(); }

  /* ── 1절 핵심지표 두 줄(110대 = 품목 / 그 외 = 처) + 팀 줄 — 견본 core_rows와 같은 칸·문구 ── */
  var SMALL_RE = /\n(?=\((?:12개월|그중))/;   // 칸 안 작은 글씨 줄(12개월 후 참고치 · '그중 n처는 이전에 신청한 품목만') — Word·성과보드가 이 경계로 나눈다
  function over12(a, kind) {             // 진행율(3개월 후 MBO 대비) 100% 초과 줄만 12개월 후 MBO 참고치
    var rx = a.rx || {};
    if (!(num(rx.pg) > 1) || !num(rx.m12)) return '';
    return kind === 'p' ? ('\n(12개월 후 MBO ' + amt(rx.m12) + ' 대비 ' + pct(rx.pg12) + ')') : ('\n(12개월 후 기준 ' + amt(rx.left12) + ')');
  }
  function coreRows(k2, dir) {
    var out = [];
    if (!k2) return out;
    [['110', '110대병원\n(품목 기준)', '품목'], ['oth', '그 외 병원·로컬\n(처 기준)', '처']].forEach(function (G) {
      [['mine', k2.mine], ['team', k2.team]].forEach(function (W) {
        var a = W[1] && W[1][G[0]]; if (!a) return;
        var team = W[0] === 'team', u = G[2], n = Number(a.n) || 0, rx = a.rx || {}, wt = a.wait || {};
        var sep = team ? ' · ' : '\n';
        var of = function (x) { return comma(n) + u + ' 중 ' + comma(Number(x) || 0) + u; };
        var app = comma(n) + u + (G[0] === '110' ? ((team ? '' : '\n') + '(' + (Number(a.units) || 0) + '병원)') : '');
        var old = (!team && num(rx.old) > 0) ? ('\n(그중 ' + comma(rx.old) + u + '는 이전에 신청한 품목만 처방 중)') : '';   // [QA🟡-5] 최근 3개월 신청 품목 실적 0인 처방 중(n>0일 때만, 작은 글씨)
        out.push({ label: team ? '  팀' : G[1], team: team, g: G[0], cells: [
          app,
          of(rx.n) + sep + amt(rx.real) + old,
          amt(rx.m3),
          pct(rx.pg) + over12(a, 'p'),
          amt(rx.left) + over12(a, 'l'),
          of(a.cut),
          of(wt.n) + sep + amt(wt.m3),
          of(a.fut)] });
      });
    });
    return out;
  }
  var CORE_HEADS = function (cmn) { return ['구분', '신청', '처방 중\n(수 · ' + cmn + '월 실적)', '3개월 후\nMBO', '진행율(실적 ÷\n3개월 후 MBO)', '남은 금액(참고 —\n예측 자동 반영 안 됨)', '끊김', '미처방\n(수 · 3개월 후 MBO)', '시작 예정\n(수)']; };
  var CORE_NOTE = '그 외(처)의 실적·MBO·진행율은 최근 3개월에 신청한 품목만 계산합니다(옛 신청 품목 실적 제외). '
    + '읽는 법 — 110대는 석 달 안에 진행율 100%에 가까워지는지, 그 외는 60% 선을 지키는지 봅니다(실측: 110대 52→58→89→108%, 그 외 65→58→55→53%).';

  var TXT = {
    title: 'MS팀 지부 월간 전략 브리핑',
    guide: '회색 칸은 폼·실적에서 자동으로 채워진 값입니다. 숫자를 고치려면 일일보고 폼에서 고친 뒤 다시 받으세요.',
    check: '제출 전 체크    □ 평가실적 실매출 기준    □ 증감 거래처에 \'왜\'    □ 잘한 점·보완 3요소    □ GAP 극복 할 일 작성    □ 두 쪽 이내',
    why: '왜 늘고 줄었나 (지부장)'
  };

  /* ── 모델 — d = crm.directors[i](goalBasis 목록 포함), meta = crm.meta, opt = {ym: 대상월, today} ── */
  function model(d, meta, opt) {
    opt = opt || {}; meta = meta || {};
    var g = (d && d.goalBasis) || null;
    if (!g || g.error) throw new Error('목표 근거 데이터가 없습니다' + (g && g.error ? (' — ' + g.error) : ''));
    if (!g.kpi2 || !g.brief) throw new Error('서버가 아직 새 버전(v71)이 아니라 전략브리핑 자료를 만들 수 없습니다');
    var ms = (g.targetMonths && g.targetMonths.length) ? g.targetMonths : [meta.nowMonth || ''];
    var ym = (opt.ym && ms.indexOf(opt.ym) >= 0) ? opt.ym : ms[0];
    var sv = (g.saved || {})[ym] || null;
    var today = opt.today || meta.today || new Date().toISOString().slice(0, 10);
    var p = g.prm || {};
    return { dir: d.name, ym: ym, isNow: ym === ms[0], cm: g.kpi2.month || g.brief.month, made: String(today).slice(0, 10),
             k2: g.kpi2, brief: g.brief, base: g.base || {}, targets: g.targets || {}, summary: sv && sv.summary, rows: (sv && sv.rows) || [],
             prm: p, listOk: !g.listOmitted && Array.isArray(p.hw) && Array.isArray(p.hrx) && Array.isArray(p.wait) && Array.isArray(p.rx),
             dcPass110: g.dcPass110 || [], cutCands: g.cutCands || [],
             v72: (g.v72 && g.v72.saved && (g.v72.saved.rows || []).length) ? g.v72 : null };   // [v72] 집중업체 연결 저장분이 있으면 4절을 두 기간으로
  }
  function fileName(kind, M) { return kind === 'docx' ? ('전략브리핑_' + M.dir + '_' + M.ym + '.docx') : ('목표근거_대상리스트_' + M.dir + '_' + M.ym + '.xlsx'); }

  /* [QA🟡-3] Word 2쪽 보장 상한 — ② 종류별 상위 5 · ③④ 상위 6(나머지 '외 n건 · 합계') · 계획 내용 60자 · GAP 할 일 5줄 */
  var ROW2_MAX = 5, ROW34_MAX = 6, WHY_MAX = 60, GAP_MAX = 5, GAP_LINE_MAX = 120;
  function clip(t, n) { t = String(t == null ? '' : t).trim(); return t.length > n ? (t.slice(0, n) + '…') : t; }
  /* 4절 행 — [구분, 내용, 금액 표시, kind(real|sum|ex|tot|gap)] */
  function basisRows(M) {
    var s = M.summary; if (!s) return null;
    var tMis = !!(M.isNow && s.목표불일치 && num(s.목표확정값) > 0);          // 성과보드 crmGbHtml과 같은 규칙(확정 목표로 GAP 다시 계산)
    var tgt = tMis ? num(s.목표확정값) : num(s.목표), sum = num(s.근거합) || 0;
    var gap = tMis ? (tgt - sum) : num(s.GAP);
    var tm = mon(M.ym), sm = (M.prm && M.prm.sum) || {}, hwN = (sm.hw && sm.hw.n) || 0, bkN = (sm.bulk && sm.bulk.n) || 0, w110 = (sm.wait && sm.wait.n) || 0;
    var bulkBy = ((M.prm && M.prm.bulk) || []).map(function (b) { return bizNm(b.co) + ' ' + b.n + '처'; }).join('·');
    var k2 = s.s2kind || {}, kH = k2.hosp || {}, kO = k2.open || {}, kB = k2.bulk || {}, kOld = k2.old || {};
    var picked = [];
    if (num(kH.n)) picked.push('처 ' + kH.n + '곳');
    if (num(kB.n)) picked.push('업체 묶음 ' + kB.n + '개 업체 ' + (kB.units || 0) + '처');
    if (num(kO.n)) picked.push('110대 ' + kO.n + '품목');
    if (num(kOld.nPos)) picked.push('지난 방식 ' + kOld.nPos + '건');
    var R = [['목표', (s.목표구분 === '임시' && !tMis) ? (tm + ' 임시 목표(폼 다음달 근거에 적은 값)') : (tm + ' 목표(마감예측 탭 \'목표\', 지부장 입력)'), amt(tgt), 'real'],
             ['① 기반', mon(s.기반월) + ' 확정 실적', amt(s.기반액), 'real'],
             ['② PRM 미처방 개시', '미처방 처 ' + hwN + '처(' + (bkN ? ('거래처별 ' + (hwN - bkN) + '처 체크 + 업체 묶음 ' + bkN + '처 \'열 처 수\': ' + bulkBy) : '거래처별 체크') + '), '
               + '110대 미처방 품목 ' + w110 + '개 체크(3개월 후 MBO × ' + Math.round((num(M.prm.rate110) || 0.5) * 100) + '%)' + (picked.length ? (' — 이번 근거: ' + picked.join(' · ')) : ' — 이번 근거: 체크 없음'), amt(s.램프업합), 'sum']];
    var it = function (k) { return M.rows.filter(function (r) { return r.항목 === k && !r.제외; }); };
    var why = function (r) { var t = String(r.근거 || '').replace(/\s+/g, ' ').trim(); return t ? (' — ' + clip(t, WHY_MAX)) : ''; };   // [QA🟡-3] 계획 내용 60자에서 자름
    var byAmt = function (a, b) { return (num(b.금액) || 0) - (num(a.금액) || 0); };
    // [QA🟡-3] 2쪽 보장 — 상위 n행 + '외 n건 · 합계'(나머지 금액은 금액 칸에, 항목 합은 그대로)
    var top = function (list, n, line) {
      list.slice(0, n).forEach(function (r) { R.push(line(r)); });
      if (list.length > n) { var rest = list.slice(n), rs = rest.reduce(function (a, r) { return a + (num(r.금액) || 0); }, 0);
        R.push(['   외', '외 ' + rest.length + '건 · 합계 ' + amt(rs), amt(rs), 'ex']); }
    };
    var kind2 = function (r) { var md = r.방식; return md === '처' ? '처' : (md === '묶음' ? '묶음' : (md === '개시' ? '110대' : '지난 방식')); };
    var r2 = it('②램프업').filter(function (r) { return num(r.금액) > 0; });
    ['처', '묶음', '110대', '지난 방식'].forEach(function (kd) {   // 종류별(처 / 업체 묶음 / 110대 품목) 금액 큰 순 상위 5
      top(r2.filter(function (r) { return kind2(r) === kd; }).sort(byAmt), ROW2_MAX, function (r) {
        var md = r.방식;
        var desc = md === '묶음' ? (r.업체 + ' · 이번 달 열 처 수 ' + (r.예상처 || 0) + '처') : ([r.업체, r.거래처].filter(Boolean).join(' · ') + (md === '처' || !r.품목 ? '' : (' · ' + r.품목)));
        return ['   ' + kd, desc + why(r), amt(r.금액), 'ex']; });
    });
    R.push(['③ 추가 볼륨 예측', '처방 중 처·110대 처방 중 품목에서 더 나올 금액 — 거래처 또는 업체 단위 자유 추가', amt(s.볼륨업합), 'sum']);
    top(it('③볼륨업').filter(function (r) { return num(r.금액) > 0 || String(r.근거 || '').trim(); }).sort(byAmt), ROW34_MAX, function (r) {
      var up = r.대상종류 === '업체';
      return ['   ' + (up ? '업체' : '거래처'), (up ? (r.업체 + ' · 업체 단위(거래처 지정 없음)') : ([r.업체, r.거래처].filter(Boolean).join(' · ') + (r.품목 ? (' · ' + r.품목) : ''))) + why(r), amt(r.금액 || 0), 'ex']; });
    R.push(['④ 신규(신약 후보)', '처방 단계 진입 예상 신약 후보·새 처방', amt(s.신규합), 'sum']);
    top(it('④신규').filter(function (r) { return num(r.금액) > 0; }).sort(byAmt), ROW34_MAX, function (r) {
      return ['   ' + (r.신약키 ? '신약' : (String(r.카테고리 || '').replace(/^\S+\s/, '') || '신규')), [r.거래처, r.품목].filter(Boolean).join(' · ') + why(r), amt(r.금액), 'ex']; });
    R.push(['근거 합', '① + ② + ③ + ④ = ' + wonTxt(sum), amt(sum), 'tot']);
    R.push(['GAP', '목표 − 근거 합 = ' + wonTxt(gap), amt(gap), 'gap']);
    var all = (s.GAP할일 && s.GAP할일.length) ? s.GAP할일.slice() : (String(s.GAP계획 || '').trim() ? String(s.GAP계획).split(/\n/).map(function (x) { return x.trim(); }).filter(Boolean) : []);
    return { rows: R, gap: gap, sum: sum, target: tgt, lines: all.slice(0, GAP_MAX).map(function (x) { return clip(x, GAP_LINE_MAX); }), linesMore: Math.max(0, all.length - GAP_MAX), linesAll: all };
  }

  /* [v72] 4절 — 집중업체 연결 저장분(goalBasis.v72.saved). 표 = 기간 2열, ② 업체 상위 8 + '외 n개 · 합계', ③ 상위 6 + 같은 꼴.
     근거 합(기간) = round(기반액) + ② 합(제외 아님) + ③ 합 · GAP = 목표 − 근거 합(폼·서버·성과보드와 같은 식) */
  var ROW72_CO = 8, ROW72_O = 6;
  function basisRows72(M) {
    var v = M.v72; if (!v) return null;
    var sv = v.saved || {}, rows = sv.rows || [], P = (v.periods || []).slice(0, 2), keys = P.map(function (p) { return p.key; });
    var isEx = function (r) { var x = r.제외; return x === true || String(x == null ? '' : x).trim().toUpperCase() === 'Y'; };
    var base = Math.round(num(v.base && v.base.value) || 0);
    var co = {}, oth = {}, order = [], oOrder = [];
    rows.forEach(function (r) { var k = String(r.period || ''), hk = String(r.행키 || ''), a = Math.round(num(r.금액) || 0);
      if (r.항목 === '②집중업체') { var c = co[hk] || (co[hk] = { nm: r.업체 || '-', type: r.유형 || '', a: {}, why: '', ex: false }); if (order.indexOf(hk) < 0) order.push(hk);
        if (isEx(r)) c.ex = true; else if (keys.indexOf(k) >= 0) c.a[k] = (c.a[k] || 0) + a; if (String(r.근거 || '').trim()) c.why = r.근거; }
      else if (r.항목 === '③그외') { var o = oth[hk] || (oth[hk] = { nm: r.대상 || r.업체 || '-', a: {}, why: '' }); if (oOrder.indexOf(hk) < 0) oOrder.push(hk);
        if (keys.indexOf(k) >= 0) o.a[k] = (o.a[k] || 0) + a; if (String(r.근거 || '').trim()) o.why = r.근거; } });
    var tot = function (x) { return keys.reduce(function (s, k) { return s + (x.a[k] || 0); }, 0); };
    var live = order.filter(function (h) { return !co[h].ex; }).map(function (h) { return co[h]; }).filter(function (c) { return tot(c) > 0 || String(c.why).trim(); });
    var exN = order.filter(function (h) { return co[h].ex; }).length;
    var os = oOrder.map(function (h) { return oth[h]; });
    var T = {}; P.forEach(function (p) { var k = p.key, s2 = 0, s3 = 0;
      order.forEach(function (h) { if (!co[h].ex) s2 += (co[h].a[k] || 0); }); os.forEach(function (o) { s3 += (o.a[k] || 0); });
      var t = (v.targets || {})[k] || {}, tg = Math.round(num(t.value) || 0), sum = base + s2 + s3;
      T[k] = { s2: s2, s3: s3, sum: sum, tg: tg, kind: t.kind || '', gap: tg > 0 ? (tg - sum) : null }; });
    var why = function (x) { var t = String(x.why || '').replace(/\s+/g, ' ').trim(); return t ? (' — ' + clip(t, WHY_MAX)) : ''; };
    var cellA = function (x, k) { return x.a[k] ? amt(x.a[k]) : '—'; };
    var R = [];
    var push = function (l, d, f, kind) { R.push([l, d].concat(P.map(function (p) { return f(p); })).concat([kind])); };
    push('목표', P.map(function (p) { var x = T[p.key]; return (p.label || p.key) + (p.kind === '분기' ? (' = ' + mon(p.month) + ' 월매출') : '') + (x.kind ? ('(' + x.kind + ')') : ''); }).join(' · '), function (p) { return T[p.key].tg > 0 ? amt(T[p.key].tg) : '미입력'; }, 'real');
    push('① 기반', mon(v.base && v.base.month) + ' 확정 실적(성장금액의 기준)', function () { return amt(base); }, 'real');
    push('② 집중업체', '지부장이 낸 집중업체 ' + (((v.focus && v.focus.rows) || []).length || order.length) + '개 중 ' + live.length + '개 업체에 근거' + (exN ? (' · 제외 ' + exN + '개') : '') + ' — ' + mon(v.base && v.base.month) + ' 대비 성장금액', function (p) { return amt(T[p.key].s2); }, 'sum');
    var top = function (list, n, lbl) {
      list.sort(function (a, b) { return tot(b) - tot(a); });
      list.slice(0, n).forEach(function (x) { push('   ' + lbl(x), x.nm + why(x), function (p) { return cellA(x, p.key); }, 'ex'); });
      if (list.length > n) { var rest = list.slice(n); push('   외', '외 ' + rest.length + '개 · 합계', function (p) { return amt(rest.reduce(function (s, x) { return s + (x.a[p.key] || 0); }, 0)); }, 'ex'); }
    };
    top(live.slice(), ROW72_CO, function (x) { return x.type ? clip(x.type, 10) : '업체'; });
    push('③ 그 외', '집중업체 밖 — 병원신약 첫 처방·다른 업체 볼륨업 등', function (p) { return amt(T[p.key].s3); }, 'sum');
    top(os.filter(function (o) { return tot(o) > 0 || String(o.why).trim(); }), ROW72_O, function () { return '대상'; });
    push('근거 합', '① + ② + ③ = ' + P.map(function (p) { return (p.kind === '분기' ? String(p.label || '').replace(/^\d+년\s*/, '') : (p.label || p.key)) + ' ' + wonTxt(T[p.key].sum); }).join(' / '), function (p) { return amt(T[p.key].sum); }, 'tot');
    push('GAP', '목표 − 근거 합 = ' + P.map(function (p) { var g = T[p.key].gap; return (p.kind === '분기' ? String(p.label || '').replace(/^\d+년\s*/, '') : (p.label || p.key)) + ' ' + (g == null ? '목표 없음' : wonTxt(g)); }).join(' / '), function (p) { var g = T[p.key].gap; return g == null ? '—' : amt(g); }, 'gap');
    var all = (sv.gapLines || []).map(function (x) { return String(x == null ? '' : x).trim(); }).filter(Boolean);
    return { P: P, T: T, rows: R, lines: all.slice(0, GAP_MAX).map(function (x) { return clip(x, GAP_LINE_MAX); }), linesMore: Math.max(0, all.length - GAP_MAX), linesAll: all, saved: sv };
  }

  /* ════════ Word(2쪽) — docx 라이브러리(D)로 만든다. 기본 Heading 스타일은 쓰지 않는다(keepNext 등 0건) ════════ */
  function docxDoc(M, D) {
    var F = { ascii: WFONT, hAnsi: WFONT, eastAsia: WFONT, cs: WFONT };
    var MM = function (mm) { return Math.round(mm * 56.6929); };
    var run = function (text, size, o) { o = o || {};
      return String(text).split('\n').map(function (t, i) {
        return new D.TextRun({ text: t, break: i ? 1 : undefined, size: Math.round(size * 2), bold: !!o.b, italics: !!o.i, color: o.c || undefined, font: F }); }); };
    var para = function (runs, o) { o = o || {};
      var x = { children: [].concat.apply([], runs), spacing: { before: Math.round((o.before || 0) * 20), after: Math.round((o.after || 0) * 20), line: Math.round(240 * (o.line || 1)), lineRule: D.LineRuleType.AUTO } };
      if (o.align === 'c') x.alignment = D.AlignmentType.CENTER; else if (o.align === 'r') x.alignment = D.AlignmentType.RIGHT;
      if (o.shade) x.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: o.shade };
      if (o.line_) x.border = { bottom: { style: D.BorderStyle.DOTTED, size: 4, space: 1, color: 'BFBFBF' } };
      return new D.Paragraph(x); };
    var cell = function (runs, w, fill, align) {
      return new D.TableCell({ children: [para(runs, { align: align })], width: { size: MM(w), type: D.WidthType.DXA }, verticalAlign: D.VerticalAlign.CENTER,
        shading: { type: D.ShadingType.CLEAR, color: 'auto', fill: fill || 'FFFFFF' } }); };
    var table = function (rows, widths, mar) { mar = mar || [25, 60];
      return new D.Table({ rows: rows.map(function (cs) { return new D.TableRow({ children: cs, cantSplit: true }); }), width: { size: widths.reduce(function (a, b) { return a + MM(b); }, 0), type: D.WidthType.DXA },
        columnWidths: widths.map(MM), layout: D.TableLayoutType.FIXED, alignment: D.AlignmentType.CENTER,
        margins: { top: mar[0], bottom: mar[0], left: mar[1], right: mar[1], marginUnitType: D.WidthType.DXA } }); };
    var title = function (text, note) { return para(run(text, 10, { b: true, c: NAVY }).concat(note ? run('  ' + note, 7.5, { c: DIM }) : []), { before: 5, after: 2 }); };

    var C = [], cmn = parseInt(String(M.cm).slice(5, 7), 10), tY = M.ym.slice(0, 4), tM = parseInt(M.ym.slice(5, 7), 10);
    // 0. 머리
    C.push(para(run(TXT.title, 15, { b: true, c: NAVY }), { after: 2 }));
    var hv = ['지부장  ' + M.dir, '대상월  ' + tY + '년 ' + tM + '월', '자료 기준  확정월 ' + String(M.cm).slice(0, 4) + '년 ' + cmn + '월 · 생성 ' + M.made, '금액 = 평가실적 실매출(VAT 제외)'];
    C.push(table([hv.map(function (v, i) { return cell(run(v, 8.5, { b: i < 2 }), [34, 30, 72, 48][i], GRAY); })], [34, 30, 72, 48]));
    C.push(para(run('■ ', 7.5, { c: 'BFBFBF' }).concat(run(TXT.guide, 7.5, { c: '404040' }), run('   흰 칸 = 지부장 작성', 7.5, { b: true, c: '404040' })), { before: 2 }));

    // 1. 핵심지표 — 두 줄 표(kpi2)
    var k2 = M.k2, lom = parseInt(String(k2.from || '').slice(5, 7), 10);
    C.push(title('1. 핵심지표 — 최근 3개월 PRM 신청(' + lom + '~' + cmn + '월 시작 + 시작 예정) · ' + cmn + '월 확정 실적 기준'));
    var W1 = [25, 14, 27, 15, 17, 20, 20, 27, 19], rows1 = [CORE_HEADS(cmn).map(function (h, i) { return cell(run(h, 7, { b: true }), W1[i], HEAD, 'c'); })];
    coreRows(k2, M.dir).forEach(function (r) {
      var cs = [cell(run(r.label, r.team ? 7 : 8, { b: !r.team, c: r.team ? DIM : undefined }), W1[0], HEAD)];
      r.cells.forEach(function (v, j) { var i = j + 1, parts = v.split(SMALL_RE), main = parts[0];
        var rs = run(main, r.team ? 7 : 8, { b: !r.team && (i === 2 || i === 4), c: r.team ? DIM : (i === 4 ? BROWN : undefined) });
        parts.slice(1).forEach(function (p) { rs = rs.concat(run('\n' + p, 6, { c: DIM })); });
        cs.push(cell(rs, W1[i], GRAY, 'c')); });
      rows1.push(cs);
    });
    C.push(table(rows1, W1, [25, 50]));
    C.push(para(run(CORE_NOTE, 6.5, { c: DIM }), { before: 1 }));

    // 2. 현황분석 — 반기 + 확정월(직전 3개월 평균·전년 같은 달) + 거래처 상위·하위 3
    var B = M.brief, bm = B.mine || {}, H = bm.half || {}, Mo = bm.mon || {}, GL = B.goal || {}, GH = GL.half || {}, hf = B.half || {}, p3 = B.p3 || [];
    C.push(title('2. 현황분석', '실적·목표·증감 = 자동 / 원인 = 지부장'));
    var W2 = [27, 21, 21, 17, 56, 42];
    // [QA🟡-8] 반기 줄 = 실적·전년 같은 기간 대비만(목표·달성률 빼고 '—') — 마감예측 탭 7~8월 목표는 9월 담당 조정 전 기준이라 지금 귀속 실적과 맞지 않음(GH = brief.goal.half는 받기만, 쓰지 않음)
    var r1 = [(hf.label || '') + '(' + parseInt(String(hf.from).slice(5, 7), 10) + '~' + parseInt(String(hf.to).slice(5, 7), 10) + '월)', eok(H.cur), '—', '—', '—',
              pct(H.yoy, true) + ' (전년 같은 기간 ' + eok(H.ly) + ')'];
    var r2 = [cmn + '월(확정월)', eok(Mo.cur), GL.value ? eok(GL.value) : '—', pct(GL.ach),
              amt(Mo.d3, true) + ' (' + pct(Mo.r3, true) + ') · 평균 ' + amt(Mo.ref3), pct(Mo.rly, true) + ' (전년 ' + cmn + '월 ' + eok(Mo.ly) + ')'];
    var rows2 = [['구분', '실적(억)', '목표(억)', '달성률', '직전 3개월 평균(' + mon(p3[0]).replace('월', '') + '~' + mon(p3[p3.length - 1]) + ') 대비', '전년 대비'].map(function (h, i) { return cell(run(h, 7.5, { b: true }), W2[i], HEAD, 'c'); })];
    [r1, r2].forEach(function (row, ri) {
      rows2.push(row.map(function (v, i) { var neg = (i === 4 && ri === 1 && num(Mo.d3) !== null && Mo.d3 < 0);
        return cell(run(v, (i === 1 || i === 2) ? 8.5 : 8, { b: i <= 1, c: neg ? BROWN : undefined }), W2[i], GRAY, i ? 'c' : undefined); }));
    });
    C.push(table(rows2, W2));
    var h1 = bm.h1, note2 = h1 ? ('상반기(1~6월) 실적 ' + eok(h1.cur) + '억 · 전년 대비 ' + pct(h1.yoy, true) + ' (전년 ' + eok(h1.ly) + '억)   ') : '';
    C.push(para(run(note2 + '전년 대비 = 전년 같은 달(기간) — 성과보드 \'전년 같은 달 대비\'와 같은 산식. 직전 3개월 평균 = 실적 있는 달 평균(성과보드와 같음). 반기 줄은 실적·전년 대비만(월 목표 기준이 담당 조정 전후로 달라 합산하지 않음).', 6.5, { c: DIM }), { before: 1 }));
    [['up', '늘어난 거래처 상위 3'], ['down', '줄어든 거래처 하위 3']].forEach(function (K) {
      C.push(para(run(K[1] + ' — ' + cmn + '월 실매출 − 직전 3개월 월평균(거래처 사업자번호 단위)', 8.5, { b: true }), { before: 3, after: 1 }));
      (bm[K[0]] || []).forEach(function (x, i) {
        var neg = x.d < 0, rr = (x.r != null) ? (amt(x.d, true) + '(' + pct(x.r, true) + ')') : (amt(x.d, true) + '(신규)');
        C.push(para(run('①②③'.charAt(i) + ' ' + x.hosp + ' · ' + cmn + '월 ' + amt(x.cur) + ' · 직전 3개월 평균 대비 ', 8.5).concat(run(rr, 8.5, { b: true, c: neg ? BROWN : undefined })), { line: 1.2, shade: GRAY }));
        C.push(para(run('   원인: ', 8, { b: true, c: DIM }).concat(run(TXT.why, 7, { c: HINT, i: true })), { after: 2, line: 1.6, line_: true }));
      });
    });

    // 3. 잘한 점 · 보완·개선사항 — 번호 빈 줄
    C.push(title('3. 잘한 점 · 보완·개선사항', '지부장 작성 — 각 줄 [사실+숫자] → [원인] → [다음 행동]. 추상적 표현·외부 핑계 금지'));
    ['잘한 점', '보완·개선사항'].forEach(function (hd) {
      C.push(para(run(hd, 9, { b: true }), { before: 2 }));
      for (var i = 1; i <= 3; i++) C.push(para(run(i + '. ', 8.5, { b: true }).concat(i === 1 ? run('사실+숫자 → 원인 → 다음 행동', 7, { c: HINT, i: true }) : []), { after: 2, line: 1.9, line_: true }));
    });

    // 수동 쪽 나눔(<w:br w:type="page"/> — pageBreakBefore 아님)
    C.push(new D.Paragraph({ children: [new D.PageBreak()], spacing: { before: 0, after: 0 } }));

    // 4. 목표 근거 — 저장분(일일보고 폼 ②)
    var B72 = basisRows72(M);   // [v72] 집중업체 연결 저장분이 있으면 두 기간 표
    if (B72) {
      var P72 = B72.P, sh72 = function (p) { return p.kind === '분기' ? String(p.label || '').replace(/^\d+년\s*/, '') : (p.label || p.key); };
      C.push(title('4. 목표 근거 — ' + P72.map(function (p) { return p.label || p.key; }).join(' · '), '자동(일일보고 폼 \'목표 근거\') · 저장 ' + (B72.saved.갱신일 || B72.saved.작성일 || '')));
      var W72 = [30, 98, 28, 28];
      var hd72 = [cell(run('구분', 7.5, { b: true }), W72[0], HEAD, 'c'), cell(run('내용', 7.5, { b: true }), W72[1], HEAD, 'c')].concat(P72.map(function (p, i) {
        return cell(run((p.label || p.key) + (p.kind === '분기' ? ('\n(' + mon(p.month) + ' 월매출)') : ''), 7.5, { b: true }), W72[2 + i], HEAD, 'c'); }));
      C.push(table([hd72].concat(B72.rows.map(function (r) { var kd = r[r.length - 1], ex = kd === 'ex', bold = !ex;
        return [cell(run(r[0], ex ? 7.5 : 8, { b: bold }), W72[0], bold ? HEAD : GRAY), cell(run(r[1], ex ? 7 : 7.5), W72[1], GRAY)].concat(P72.map(function (p, i) {
          var gp = kd === 'gap' && B72.T[p.key].gap != null && B72.T[p.key].gap > 0;
          return cell(run(r[2 + i], ex ? 7.5 : 8, { b: bold, c: gp ? BROWN : undefined }), W72[2 + i], GRAY, 'r'); })); })), W72));
      var gt72 = P72.map(function (p) { var g = B72.T[p.key].gap; return sh72(p) + ' ' + (g == null ? '—' : amt(g)); }).join(' · ');
      C.push(para(run('GAP 극복을 위한 할 일 — GAP ' + gt72, 9, { b: true }).concat(run('  지부장 작성 · 누구에게·무엇을·언제까지', 6.5, { c: DIM })), { before: 4 }));
      for (var j = 1; j <= Math.max(3, B72.lines.length); j++) {
        var L72 = B72.lines[j - 1];
        C.push(para(run(j + '. ', 8.5, { b: true }).concat(L72 ? run(L72 + ' ', 8).concat(run('(일일보고 폼)', 7, { c: DIM })) : []), { after: 2, line: 1.9, line_: true, shade: L72 ? GRAY : undefined }));
      }
      if (B72.linesMore) C.push(para(run('외 ' + B72.linesMore + '건 — 일일보고 폼 참조', 7.5, { c: DIM }), { after: 2 }));
      C.push(para(run(TXT.check, 7.5, { b: true, c: '404040' }), { before: 6 }));
      return new D.Document({
        creator: 'MS팀 성과관리 CRM', title: TXT.title + ' — ' + M.dir + ' ' + M.ym,
        styles: { default: { document: { run: { font: F, size: 16 }, paragraph: { spacing: { before: 0, after: 0 } } } } },
        sections: [{ properties: { page: { size: { width: MM(210), height: MM(297) }, margin: { top: MM(11), bottom: MM(11), left: MM(13), right: MM(13), header: MM(6), footer: MM(6) } } }, children: C }]
      });
    }
    var BR = basisRows(M), s = M.summary;
    C.push(title('4. 목표 근거 — ' + tM + '월', '자동(일일보고 폼 \'' + (M.isNow ? '이번달' : '다음달') + ' 목표 근거\')' + (s ? (' · 저장 ' + (s.갱신일 || s.작성일 || '')) : '')));
    if (!BR) {
      C.push(para(run('근거 미작성 — 일일보고 폼 ②에서 작성', 9, { b: true, c: '404040' }), { before: 2, after: 2, line: 1.4, shade: GRAY }));
    } else {
      var W4 = [32, 124, 28];
      C.push(table(BR.rows.map(function (r) { var ex = r[3] === 'ex', bold = !ex;
        return [cell(run(r[0], ex ? 7.5 : 8, { b: bold }), W4[0], bold ? HEAD : GRAY),
                cell(run(r[1], ex ? 7 : 7.5), W4[1], GRAY),
                cell(run(r[2], ex ? 7.5 : 8, { b: bold, c: (r[3] === 'gap' && BR.gap > 0) ? BROWN : undefined }), W4[2], GRAY, 'r')]; }), W4));
    }
    // GAP 극복을 위한 할 일 — 저장된 줄(폼)을 채우고 남는 줄은 빈 줄(최소 3줄)
    var gapTxt = BR ? amt(BR.gap) : '—', lines = BR ? BR.lines : [];
    C.push(para(run('GAP 극복을 위한 할 일 — GAP ' + gapTxt, 9, { b: true }).concat(run('  지부장 작성 · 누가·언제까지·얼마', 6.5, { c: DIM })), { before: 4 }));
    for (var i = 1; i <= Math.max(3, lines.length); i++) {
      var L = lines[i - 1];
      C.push(para(run(i + '. ', 8.5, { b: true }).concat(L ? run(L + ' ', 8).concat(run('(일일보고 폼)', 7, { c: DIM })) : []), { after: 2, line: 1.9, line_: true, shade: L ? GRAY : undefined }));
    }
    if (BR && BR.linesMore) C.push(para(run('외 ' + BR.linesMore + '건 — 일일보고 폼 참조', 7.5, { c: DIM }), { after: 2 }));   // [QA🟡-3] 저장분 5줄 넘으면
    C.push(para(run(TXT.check, 7.5, { b: true, c: '404040' }), { before: 6 }));

    return new D.Document({
      creator: 'MS팀 성과관리 CRM', title: TXT.title + ' — ' + M.dir + ' ' + M.ym,
      styles: { default: { document: { run: { font: F, size: 16 }, paragraph: { spacing: { before: 0, after: 0 } } } } },
      sections: [{ properties: { page: { size: { width: MM(210), height: MM(297) }, margin: { top: MM(11), bottom: MM(11), left: MM(13), right: MM(13), header: MM(6), footer: MM(6) } } }, children: C }]
    });
  }

  /* ════════ 엑셀 — 견본 7시트(요약·미처방 처·업체 묶음·110대 품목·처방 중 처·끊긴 처·110대 DC 통과 처) ════════ */
  function xlsxBook(M, X) {
    if (!M.listOk) throw new Error('대상 목록을 받지 못했습니다 — 새로고침 후 다시 눌러 주세요');
    var wb = new X.Workbook(); wb.creator = 'MS팀 성과관리 CRM';
    var cmn = parseInt(String(M.cm).slice(5, 7), 10), lom = parseInt(String(M.k2.from || '').slice(5, 7), 10);
    var basis = '기준: 확정월 ' + M.cm + ' · 생성 ' + M.made + ' · 금액 = 평가실적 실매출(VAT 제외), 만원';
    var thin = { style: 'thin', color: { argb: 'FFBFBFBF' } }, BOX = { top: thin, left: thin, bottom: thin, right: thin };
    var counts = {};
    var dispW = function (s) { var n = 0; for (var i = 0; i < s.length; i++) n += s.charCodeAt(i) > 0x1100 ? 2 : 1; return n; };
    function sheet(ws, name, heads, rows, widths, o) {
      o = o || {}; var sr = o.startRow || 2;
      if (!o.keepLine1) ws.getCell(1, 1).value = basis + (o.note ? (' · ' + o.note) : '');
      heads.forEach(function (h, j) { var c = ws.getCell(sr, j + 1); c.value = h; c.border = BOX;
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + HEAD } };
        c.font = { name: XFONT, size: 9, bold: true }; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
      ws.getRow(sr).height = 36;
      rows.forEach(function (r, i) { r.forEach(function (v, j) { var c = ws.getCell(sr + 1 + i, j + 1); c.value = (v === undefined ? null : v); c.border = BOX;
        if ((o.pct || []).indexOf(j + 1) >= 0 && typeof v === 'number') c.numFmt = '0.0%'; else if (typeof v === 'number' && Math.round(v) === v) c.numFmt = '#,##0'; }); });
      widths.forEach(function (w, j) {               // 머리글이 두 줄 안에 들어가게 너비 보정(견본과 같은 식)
        var ls = String(heads[j] || '').split('\n'), nd = ls.map(function (l) { return dispW(l) * 1.1; });
        var ww = Math.max(w, ls.length > 1 ? Math.max.apply(null, nd) + 1 : nd[0] / 2 + 2);
        var col = ws.getColumn(j + 1); col.width = (sr > 3) ? Math.max(col.width || 0, ww) : ww; });
      if (o.filter !== false && rows.length) ws.autoFilter = { from: { row: sr, column: 1 }, to: { row: sr + rows.length, column: heads.length } };
      counts[name] = (counts[name] || 0) + rows.length;
    }
    var mk = function (name, frozen) { return wb.addWorksheet(name, { views: [frozen ? { state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', zoomScale: 90 } : { zoomScale: 90 }] }); };
    var ref12 = function (real, m3, m12) { return (num(m3) && real / m3 > 1 && num(m12)) ? [man(m12), real / m12] : [null, null]; };
    var P = M.prm, bulk = P.bulk || [], bulkCos = bulk.map(function (b) { return bizNm(b.co); }).join('·');

    // ① 요약 — 1절과 같은 두 줄 표(줄바꿈 → 공백) + 시트 안내
    var ws = mk('요약', false);
    ws.getCell(1, 1).value = basis;
    ws.getCell(2, 1).value = '핵심지표 — 최근 3개월 PRM 신청(' + lom + '~' + cmn + '월 시작 + 시작 예정) · 그 외 병원·로컬은 최근 3개월에 신청한 품목만 계산';
    ws.getCell(2, 1).font = { name: XFONT, size: 10, bold: true };
    var CR = coreRows(M.k2, M.dir).map(function (r) { return [r.team ? '팀' : r.label.replace(/\n/g, ' ').trim()].concat(r.cells.map(function (c) { return c.replace(/\n/g, ' '); })); });
    sheet(ws, '요약', ['구분', '신청', '처방 중(수 · ' + cmn + '월 실적)', '3개월 후 MBO', '진행율(실적 ÷ 3개월 후 MBO)', '남은 금액(참고)', '끊김', '미처방(수 · 3개월 후 MBO)', '시작 예정'],
          CR, [22, 16, 28, 14, 30, 24, 20, 28, 18], { startRow: 3, filter: false, keepLine1: true });
    var guide = ['시트 안내',
      '미처방 처 — 폼 ②에서 거래처별로 체크하는 처' + (bulk.length ? ('(' + bulkCos + ' 처는 \'업체 묶음\' 시트)') : ''),
      '업체 묶음 — ' + (bulk.length ? (bulkCos + ': 이번 달 열 처 수로 넣는 업체와 그 업체의 미처방 처') : '해당 없음'),
      '110대 품목 — 미처방·시작 예정·처방 중 품목(체크 기본 금액 = 3개월 후 MBO × ' + Math.round((num(P.rate110) || 0.5) * 100) + '%)',
      '처방 중 처 — 추가 볼륨 예측 후보(남은 금액은 참고, 예측에 자동으로 들어가지 않음)',
      '끊긴 처 — 처방이 나오다가 ' + cmn + '월에 0인 처(@ 후보)',
      '110대 DC 통과 처 — 신약 DC를 통과한 110대 병원(@ 후보)',
      '진행율이 100%를 넘은 줄에만 12개월 후 MBO를 참고로 함께 적었습니다.'];
    guide.forEach(function (t, i) { var c = ws.getCell(3 + CR.length + 2 + i, 1); c.value = t; c.font = { name: XFONT, size: 10, bold: i === 0 }; });

    // ② 미처방 처(묶음 업체 처 제외) — 폼 ② 목록 그대로
    var hwInd = P.hw.filter(function (x) { return !x.bulk; });
    sheet(mk('미처방 처', true), '미처방 처', ['구분', '업체', '거래처', '신청 품목 수', '시작월', '3개월 후 MBO(만)', '주요 품목'],
      hwInd.map(function (x) { return [x.fut ? '시작 예정' : '미처방', x.co, x.hosp, Number(x.n) || 0, x.start, man(x.m3), (x.p || []).slice(0, 4).join(', ')]; }),
      [11, 24, 30, 12, 10, 16, 44], { note: bulk.length ? (bulkCos + ' 처는 \'업체 묶음\' 시트') : null });

    // ③ 업체 묶음 — 업체 줄 + 그 업체의 미처방 처(묶음 업체가 없으면 '해당 없음')
    ws = mk('업체 묶음', true);
    if (!bulk.length) {
      ws.getCell(1, 1).value = basis;
      sheet(ws, '업체 묶음', ['업체 묶음'], [['해당 없음 — 이번 달 열 처 수로 넣는 업체가 없습니다']], [60], { filter: false, keepLine1: true });
    } else {
      var bRows = bulk.map(function (b) { var o = b.openRate || {}; return [b.co, Number(b.n) || 0, (o.den || 0) + '처 중 ' + (o.num || 0) + '처']; });
      sheet(ws, '업체 묶음', ['업체', '미처방 처 수', '최근 3개월 신청 처 중 ' + cmn + '월 처방 처'], bRows, [24, 14, 30], { filter: false });
      var r0 = 2 + bRows.length + 3;
      ws.getCell(r0 - 1, 1).value = '그 업체의 미처방 처'; ws.getCell(r0 - 1, 1).font = { name: XFONT, size: 10, bold: true };
      var hb = P.hw.filter(function (x) { return x.bulk; });
      sheet(ws, '업체 묶음', ['업체', '거래처', '신청 품목 수', '시작월', '3개월 후 MBO(만)'], hb.map(function (x) { return [x.co, x.hosp, Number(x.n) || 0, x.start, man(x.m3)]; }),
        [24, 30, 30, 10, 16], { startRow: r0, keepLine1: true });
      counts['업체 묶음'] = '업체 ' + bRows.length + ' · 처 ' + hb.length;
    }

    // ④ 110대 품목(1절과 같은 기준: 미처방·시작 예정 = 폼 목록, 처방 중·끊김 = 핵심지표 대상 in3)
    var ord = { '미처방': 0, '시작 예정': 1, '처방 중': 2, '끊김': 3 };
    var items = P.wait.map(function (b) { return { st: b.fut ? '시작 예정' : '미처방', b: b, real: 0 }; })
      .concat(P.rx.filter(function (b) { return b.in3; }).map(function (b) { return { st: b.st === '끊김' ? '끊김' : '처방 중', b: b, real: Number(b.aT) || 0 }; }));
    items.sort(function (a, b) { return (ord[a.st] - ord[b.st]) || String(a.b.start).localeCompare(String(b.b.start)) || String(a.b.hosp).localeCompare(String(b.b.hosp)); });
    sheet(mk('110대 품목', true), '110대 품목', ['구분', '병원', '품목', '업체', '시작월', '3개월 후 MBO(만)', cmn + '월 실적(만)', '진행율', '12개월 후 MBO(만)\n— 진행율 100% 넘은 품목만', '12개월 후 MBO 대비 진행율'],
      items.map(function (x) { var b = x.b, m3 = num(b.m3);
        return [x.st, b.hosp, b.prod, b.co, b.start, man(m3), x.real ? man(x.real) : null, (m3 && x.real) ? x.real / m3 : null].concat(ref12(x.real, m3, b.m12)); }),
      [11, 30, 28, 22, 10, 16, 13, 10, 22, 18], { pct: [8, 10] });

    // ⑤ 처방 중 처(핵심지표 '처방 중' = in3 · 실적 = 최근 3개월 신청 품목만)
    var rx = P.hrx.filter(function (h) { return h.in3 && h.st !== '끊김'; }).slice().sort(function (a, b) { return String(a.co).localeCompare(String(b.co)) || ((b.a3 || 0) - (a.a3 || 0)); });
    // [QA🟡-5] 최근 3개월 신청 품목 실적 0(이전에 신청한 품목만 처방 중) 행 = 실적·진행율·남은 금액 빈칸 + 비고
    var isOld = function (h) { return h.old != null ? !!h.old : !(Number(h.a3) > 0); };
    sheet(mk('처방 중 처', true), '처방 중 처', ['업체', '거래처', cmn + '월 실적(만)', '3개월 후 MBO(만)', '진행율', '남은 금액(만, 참고)', '12개월 후 MBO(만)\n— 진행율 100% 넘은 처만', '12개월 후 MBO 대비 진행율', '비고'],
      rx.map(function (h) { var a3 = Number(h.a3) || 0, m3 = num(h.m3);
        if (isOld(h)) return [h.co, h.hosp, null, man(m3), null, null, null, null, '이전 신청 품목만 처방 중'];
        return [h.co, h.hosp, man(a3), man(m3), m3 ? a3 / m3 : null, man(h.left != null ? h.left : Math.max((m3 || 0) - a3, 0))].concat(ref12(a3, m3, h.m12), [null]); }),
      [24, 30, 14, 16, 10, 18, 22, 18, 22], { pct: [5, 8], note: '최근 3개월에 신청한 품목만 계산' });

    // ⑥ 끊긴 처(@ 후보 ㉯ — cutCands 처)
    var cut = M.cutCands.filter(function (c) { return c.kind === '처'; }).slice().sort(function (a, b) { return String(a.co).localeCompare(String(b.co)) || String(a.lastM).localeCompare(String(b.lastM)); });
    sheet(mk('끊긴 처', true), '끊긴 처', ['업체', '거래처', '마지막 처방월', '마지막 처방 실적(만)'], cut.map(function (c) { return [c.co, c.hosp, c.lastM || '', man(c.lastAmt)]; }),
      [24, 30, 14, 20], { note: '최근 3개월에 신청한 품목 기준 · @ 후보' });

    // ⑦ 110대 DC 통과 처(@ 후보 ㉮)
    sheet(mk('110대 DC 통과 처', true), '110대 DC 통과 처', ['병원', '품목', '업체', 'MBO(만)'], M.dcPass110.map(function (x) { return [x.병원, x.품목, x.업체, man(x.MBO)]; }),
      [32, 30, 24, 12], { note: '@ 후보' });

    // 서체 — 쓴 범위 전 셀 나눔스퀘어(크기·굵기 유지)
    wb.eachSheet(function (s) {
      for (var r = 1; r <= s.rowCount; r++) for (var c = 1; c <= Math.max(s.columnCount, 1); c++) {
        var cl = s.getCell(r, c), f = cl.font || {};
        cl.font = { name: XFONT, size: f.size || 10, bold: !!f.bold };
      }
    });
    wb._counts = counts;
    return wb;
  }

  /* ════════ 브라우저 — CDN 지연 로드 · 내려받기 ════════ */
  var loading = {};
  function loadLib(key) {
    var L = LIB[key];
    if (root[L.global]) return Promise.resolve(root[L.global]);
    if (loading[key]) return loading[key];
    loading[key] = new Promise(function (ok, no) {
      var s = document.createElement('script'); s.src = L.url; s.async = true;
      var to = setTimeout(function () { no(new Error(L.name + '를 불러오지 못했습니다(시간 초과) — 연결을 확인하고 다시 눌러 주세요')); }, 30000);
      s.onload = function () { clearTimeout(to); root[L.global] ? ok(root[L.global]) : no(new Error(L.name + '를 불러오지 못했습니다 — 새로고침 후 다시 눌러 주세요')); };
      s.onerror = function () { clearTimeout(to); no(new Error(L.name + '를 불러오지 못했습니다(' + L.url + ') — 새로고침 후 다시 눌러 주세요. 계속되면 관리자에게 알려 주세요(파일 누락)')); };
      document.head.appendChild(s);
    }).catch(function (e) { loading[key] = null; throw e; });
    return loading[key];
  }
  function save(blob, name) {
    if (root.navigator && root.navigator.msSaveOrOpenBlob) { root.navigator.msSaveOrOpenBlob(blob, name); return; }
    var a = document.createElement('a'), u = URL.createObjectURL(blob);
    a.href = u; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(u); a.remove(); }, 1500);
  }
  function download(kind, d, meta, opt) {
    var M;
    try { M = model(d, meta, opt); } catch (e) { return Promise.reject(e); }
    if (kind === 'docx') return loadLib('docx').then(function (D) { return D.Packer.toBlob(docxDoc(M, D)); }).then(function (b) { save(b, fileName('docx', M)); return M; });
    if (!M.listOk) return Promise.reject(new Error('대상 목록을 받지 못했습니다 — 새로고침 후 다시 눌러 주세요'));
    return loadLib('exceljs').then(function (X) { return xlsxBook(M, X).xlsx.writeBuffer(); })
      .then(function (buf) { save(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName('xlsx', M)); return M; });
  }

  var API = { LIB: LIB, amt: amt, eok: eok, pct: pct, comma: comma, coreRows: coreRows, CORE_HEADS: CORE_HEADS, CORE_NOTE: CORE_NOTE, SMALL_RE: SMALL_RE,
              model: model, basisRows: basisRows, basisRows72: basisRows72, docxDoc: docxDoc, xlsxBook: xlsxBook, fileName: fileName, loadLib: loadLib, download: download };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else root.BriefExport = API;
})(typeof window !== 'undefined' ? window : this);
