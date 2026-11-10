/* 화면 계층: 파일 읽기(SheetJS) · 상태 보관 · 렌더링. 변환 규칙은 모두 core.js에 있다. */
(function () {
  'use strict';
  var C = window.OrderCore;
  var $ = function (s) { return document.querySelector(s); };

  var SAMPLE_FILES = [
    { url: 'samples/a-market-orders.xlsx', name: 'A마켓_주문_샘플.xlsx' },
    { url: 'samples/b-market-orders.xlsx', name: 'B마켓_주문_샘플.xlsx' },
    { url: 'samples/logistics-tracking-reply.xlsx', name: '물류사_송장회신_샘플.xlsx' }
  ];

  var app = {
    state: C.createState(),
    replies: [],
    orderLog: [],
    replyLog: [],
    view: null // { lines, built, match, summary }
  };

  // ------------------------------------------------------------------ 유틸
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function today() {
    var d = new Date();
    return d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
  }
  function toast(msg, ok) {
    var el = $('#toast');
    el.textContent = msg;
    el.className = 'toast' + (ok ? ' ok' : '');
    el.hidden = false;
    clearTimeout(toast.t);
    toast.t = setTimeout(function () { el.hidden = true; }, 6000);
  }
  function channelBadge(ch) {
    return '<span class="badge ' + (ch === 'A' ? 'b-a' : 'b-b') + '">' + esc(C.MARKETS[ch].label) + '</span>';
  }
  function table(headers, rows, emptyText) {
    if (!rows.length) return '<div class="empty">' + esc(emptyText || '표시할 데이터가 없습니다.') + '</div>';
    return '<div class="table-wrap"><table><thead><tr>' +
      headers.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') +
      '</tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
  }

  // ------------------------------------------------------------------ 파일 처리
  function readRows(buffer) {
    var wb = XLSX.read(buffer, { type: 'array' });
    var ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  }

  function handleFile(name, buffer) {
    var rows;
    try { rows = readRows(buffer); } catch (e) {
      app.orderLog.push({ name: name, type: null, error: 'Excel 파일을 읽을 수 없습니다' });
      return;
    }
    var fmt = C.detectFormat(rows);
    if (fmt.type === 'A' || fmt.type === 'B') {
      var parsed = C.parseOrderRows(rows, { fileName: name });
      var res = C.ingest(app.state, parsed.lines);
      app.state = res.state;
      app.orderLog.push({ name: name, type: fmt.type, added: res.added, duplicates: res.duplicates });
    } else if (fmt.type === 'REPLY') {
      var rep = C.parseReplyRows(rows);
      var merged = C.mergeReplies(app.replies, rep.records);
      app.replies = merged.records;
      app.replyLog.push({ name: name, type: 'REPLY', added: merged.added, duplicates: merged.duplicates });
    } else {
      app.orderLog.push({ name: name, type: null, error: '알 수 없는 양식 (A마켓·B마켓 주문 또는 송장 회신만 지원)' });
    }
  }

  function handleFileList(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return Promise.resolve();
    return Promise.all(files.map(function (f) {
      return f.arrayBuffer().then(function (buf) { return { name: f.name, buf: buf }; });
    })).then(function (items) {
      items.forEach(function (it) { handleFile(it.name, it.buf); });
      render();
    });
  }

  function loadSamples() {
    if (location.protocol === 'file:') {
      toast('file:// 로 열면 브라우저 보안 정책상 샘플을 불러올 수 없습니다. 로컬 서버(README 참고)로 열거나 samples 폴더의 파일을 직접 올려 주세요.');
      return Promise.resolve();
    }
    var btn = $('#btn-demo');
    btn.disabled = true;
    return Promise.all(SAMPLE_FILES.map(function (s) {
      return fetch(s.url).then(function (r) {
        if (!r.ok) throw new Error(s.url);
        return r.arrayBuffer();
      }).then(function (buf) { return { name: s.name, buf: buf }; });
    })).then(function (items) {
      items.forEach(function (it) { handleFile(it.name, it.buf); });
      render();
      toast('샘플 주문 2개 파일과 송장 회신 1개 파일을 불러왔습니다. 같은 버튼을 다시 눌러도 중복 집계되지 않습니다.', true);
    }).catch(function () {
      toast('샘플 파일을 불러오지 못했습니다.');
    }).then(function () { btn.disabled = false; });
  }

  function simulateReply() {
    if (!app.view || !app.view.built.shipments.length) return;
    var rows = C.simulateReplyRows(app.view.built.shipments, '가상택배');
    var merged = C.mergeReplies(app.replies, C.parseReplyRows(rows).records);
    app.replies = merged.records;
    app.replyLog.push({ name: '가상 회신 (현재 발주서 기준)', type: 'REPLY', added: merged.added, duplicates: merged.duplicates });
    render();
  }

  function reset() {
    app.state = C.createState();
    app.replies = [];
    app.orderLog = [];
    app.replyLog = [];
    $('#toast').hidden = true;
    render();
  }

  // ------------------------------------------------------------------ 내보내기
  function downloadSheets(fileName, sheets) {
    var wb = XLSX.utils.book_new();
    sheets.forEach(function (s) {
      var ws = XLSX.utils.aoa_to_sheet(s.rows);
      ws['!cols'] = s.rows[0].map(function (_, i) {
        var w = s.rows.reduce(function (m, r) { return Math.max(m, String(r[i] === undefined ? '' : r[i]).length); }, 4);
        return { wch: Math.min(Math.max(8, Math.ceil(w * 1.7)), 48) };
      });
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    });
    XLSX.writeFile(wb, fileName);
  }

  function exceptionRows(exceptions) {
    var rows = [['구분', '유형', '판매채널', '주문번호', '상품주문번호', '수취인(원본)', '내용', '원본 파일', '원본 행']];
    exceptions.forEach(function (e) {
      rows.push([e.level === 'error' ? '출고 제외' : '확인', e.label, e.channelLabel, e.orderNo, e.lineNo,
        e.recipient, e.message, e.source.file, e.source.row]);
    });
    return rows;
  }

  function downloadPurchaseOrder() {
    var v = app.view;
    downloadSheets('물류사_발주서_' + today() + '.xlsx', [
      { name: '발주서', rows: C.toPurchaseOrderRows(v.built.shipments) },
      { name: '예외', rows: exceptionRows(v.built.exceptions) }
    ]);
  }

  function downloadUpload(market) {
    var rows = C.toMarketUploadRows(app.view.match.matched, market);
    downloadSheets(C.MARKETS[market].label + '_송장업로드_' + today() + '.xlsx', [{ name: '송장업로드', rows: rows }]);
  }

  function downloadHold() {
    var m = app.view.match;
    var held = [['유형', '발주번호', '송장번호', '수취인(회신)', '내용', '회신 행']];
    m.held.forEach(function (h) { held.push([h.label, h.ref, h.trackingNo, h.recipient, h.message, h.row]); });
    var un = [['판매채널', '주문번호', '상품주문번호', '수취인', '출고번호']];
    m.unanswered.forEach(function (u) { un.push([u.line.channelLabel, u.line.orderNo, u.line.lineNo, u.line.recipient, u.shipmentId]); });
    downloadSheets('송장_확인필요_' + today() + '.xlsx', [
      { name: '보류', rows: held },
      { name: '미회신', rows: un },
      { name: '주문 예외', rows: exceptionRows(app.view.built.exceptions) }
    ]);
  }

  // ------------------------------------------------------------------ 렌더링
  function compute() {
    var lines = C.allLines(app.state);
    var built = C.buildShipments(lines);
    var match = app.replies.length ? C.matchTracking(built.shipments, app.replies) : null;
    return { lines: lines, built: built, match: match, summary: C.summarize(lines, built, match) };
  }

  function renderLogs() {
    function item(l) {
      var badge = l.type === 'REPLY' ? '<span class="badge b-reply">송장 회신</span>'
        : l.type ? channelBadge(l.type) : '<span class="badge b-err">인식 불가</span>';
      var detail = l.error ? '<span class="badge b-err">' + esc(l.error) + '</span>'
        : '<span class="badge b-ok">' + l.added + '건 추가</span>' +
          (l.duplicates ? '<span class="badge b-muted">' + l.duplicates + '건 중복 무시</span>' : '');
      return '<li>' + badge + '<span class="name">' + esc(l.name) + '</span>' + detail + '</li>';
    }
    $('#log-orders').innerHTML = app.orderLog.slice().reverse().map(item).join('');
    $('#log-reply').innerHTML = app.replyLog.slice().reverse().map(item).join('');
  }

  function renderKpis(s) {
    var t = s.tracking;
    var kpis = [
      { label: '주문', value: s.orderCount, sub: '주문 라인 ' + s.lineCount + '건' },
      { label: '출고 대상', value: s.shippableLines, sub: '제외 ' + s.blockedLines + '건' },
      { label: '출고(합포장 후)', value: s.shipmentCount, sub: '합포장 ' + s.mergedShipments + '건' },
      { label: '박스', value: s.boxCount, sub: '내품 ' + s.unitCount + '개' },
      { label: '예외', value: s.blockedLines + s.warningCount, sub: '제외 ' + s.blockedLines + ' · 확인 ' + s.warningCount, alert: s.blockedLines > 0 },
      { label: '송장 매칭', value: t ? t.matched : '-', sub: t ? '보류 ' + t.held + ' · 미회신 ' + t.unanswered : '회신 대기', alert: t && (t.held + t.unanswered) > 0 }
    ];
    $('#kpis').innerHTML = kpis.map(function (k) {
      return '<div class="kpi' + (k.alert ? ' alert' : '') + '"><div class="label">' + esc(k.label) +
        '</div><div class="value">' + esc(k.value) + '</div><div class="sub">' + esc(k.sub) + '</div></div>';
    }).join('');

    var bt = s.boxByType;
    var total = s.boxCount || 1;
    var seg = function (cls, n) { return n ? '<i class="' + cls + '" style="width:' + (n / total * 100) + '%"></i>' : ''; };
    $('#box-mix').innerHTML = s.boxCount
      ? '<span><b>박스 구성</b></span>' +
        '<span><i class="dot c-single"></i>싱글팩 ' + bt.SINGLE + '</span>' +
        '<span><i class="dot c-double"></i>더블팩 ' + bt.DOUBLE + '</span>' +
        '<span><i class="dot c-six"></i>6구 박스 ' + bt.SIX + '</span>' +
        '<span class="bar" role="img" aria-label="박스 유형 비율">' + seg('c-single', bt.SINGLE) + seg('c-double', bt.DOUBLE) + seg('c-six', bt.SIX) + '</span>'
      : '';

    var ch = s.byChannel;
    $('#dash-note').textContent = s.lineCount
      ? Object.keys(ch).sort().map(function (c) { return C.MARKETS[c].label + ' ' + ch[c].orders + '주문/' + ch[c].lines + '라인'; }).join(' · ')
      : '주문 파일을 올리면 집계됩니다.';
  }

  function renderShipments(v) {
    var trackingByRef = {};
    if (v.match) v.match.matched.forEach(function (m) { trackingByRef[m.ref] = m.trackingNo; });
    var rows = v.built.shipments.map(function (s) {
      var boxes = s.boxes.map(function (b) {
        var contents = b.contents.map(function (c) { return c.name + ' ' + c.units; }).join(', ');
        return '<span class="box-chip" title="' + esc(b.ref) + '">' + esc(b.typeLabel) + ' <em>' + esc(contents) + '</em></span>';
      }).join('');
      return '<tr><td class="mono">' + esc(s.id) + '</td><td>' + channelBadge(s.channel) + '</td><td>' + esc(s.recipient) +
        '</td><td class="addr">' + esc(s.address) + '</td><td>' + esc(s.orderNos.join(', ')) +
        (s.merged ? ' <span class="badge b-ok">합포장 ' + s.lines.length + '</span>' : '') +
        '</td><td class="num">' + s.totalUnits + '</td><td><div class="boxes">' + boxes + '</div></td></tr>';
    });
    $('#panel-shipments').innerHTML = table(['출고번호', '채널', '수취인', '주소', '주문', '수량', '박스 구성'], rows,
      '주문 파일을 올리거나 "샘플 데이터로 바로 체험"을 눌러 보세요.');
  }

  function renderExceptions(v) {
    var ex = v.built.exceptions.slice().sort(function (a, b) {
      return (a.level === b.level ? 0 : a.level === 'error' ? -1 : 1);
    });
    var rows = ex.map(function (e) {
      var lv = e.level === 'error' ? '<span class="badge b-err">출고 제외</span>' : '<span class="badge b-warn">확인</span>';
      return '<tr><td>' + lv + '</td><td>' + esc(e.label) + '</td><td>' + channelBadge(e.channel) + '</td><td class="mono">' +
        esc(e.orderNo) + '</td><td>' + esc(e.recipient) + '</td><td>' + esc(e.message) +
        '</td><td class="mono">' + esc(e.source.file) + ' : ' + esc(e.source.row) + '행</td></tr>';
    });
    $('#panel-exceptions').innerHTML = table(['구분', '유형', '채널', '주문번호', '수취인', '내용', '원본 위치'], rows, '예외가 없습니다.');
    $('#cnt-exceptions').textContent = ex.length;
  }

  function renderTracking(v) {
    var m = v.match;
    if (!m) {
      $('#panel-tracking').innerHTML = '<div class="empty">송장 회신 파일을 올리거나 "현재 발주서로 가상 회신 만들기"를 눌러 보세요.</div>';
      $('#cnt-tracking').textContent = 0;
      return;
    }
    var matched = m.matched.map(function (x) {
      return '<tr><td>' + channelBadge(x.line.channel) + '</td><td class="mono">' + esc(x.line.orderNo) + '</td><td class="mono">' +
        esc(x.line.lineNo) + '</td><td>' + esc(x.line.recipient) + '</td><td class="mono">' + esc(x.ref) + '</td><td>' +
        esc(x.carrier) + '</td><td class="mono">' + esc(x.trackingNo) + '</td></tr>';
    });
    var held = m.held.map(function (h) {
      return '<tr><td><span class="badge b-err">' + esc(h.label) + '</span></td><td class="mono">' + esc(h.ref) + '</td><td class="mono">' +
        esc(h.trackingNo) + '</td><td>' + esc(h.message) + '</td><td class="num">' + esc(h.row) + '</td></tr>';
    });
    var un = m.unanswered.map(function (u) {
      return '<tr><td>' + channelBadge(u.line.channel) + '</td><td class="mono">' + esc(u.line.orderNo) + '</td><td class="mono">' +
        esc(u.line.lineNo) + '</td><td>' + esc(u.line.recipient) + '</td><td class="mono">' + esc(u.shipmentId) + '</td></tr>';
    });
    var cA = m.matched.filter(function (x) { return x.line.channel === 'A'; }).length;
    $('#panel-tracking').innerHTML =
      '<div class="summary-row"><span class="badge b-ok">매칭 ' + m.matched.length + '건 (A마켓 ' + cA + ' · B마켓 ' + (m.matched.length - cA) + ')</span>' +
      '<span class="badge b-err">보류 ' + m.held.length + '건</span><span class="badge b-warn">미회신 주문 ' + m.unanswered.length + '건 · 박스 ' + m.unansweredBoxes.length + '개</span></div>' +
      '<h3 class="section-title">보류된 회신 — 업로드 파일에서 제외</h3>' +
      table(['사유', '발주번호', '송장번호', '내용', '회신 행'], held, '보류된 회신이 없습니다.') +
      '<h3 class="section-title">송장 미확정 주문</h3>' +
      table(['채널', '주문번호', '상품주문번호', '수취인', '출고번호'], un, '모든 출고 주문에 송장이 확정되었습니다.') +
      '<h3 class="section-title">매칭 완료 — 마켓 업로드 대상</h3>' +
      table(['채널', '주문번호', '상품주문번호', '수취인', '발주번호', '택배사', '송장번호'], matched, '매칭된 주문이 없습니다.');
    $('#cnt-tracking').textContent = m.matched.length;
  }

  function renderLines(v) {
    var rows = v.lines.map(function (l) {
      return '<tr><td class="mono">' + esc(l.uid) + '</td><td>' + channelBadge(l.channel) + '</td><td>' + esc(l.recipient) +
        '</td><td>' + esc(l.option) + '</td><td>' + esc(l.skuName || '-') + '</td><td class="num">' + esc(l.qtyRaw) +
        '</td><td class="num">' + (l.packSize || '-') + '</td><td class="num">' + l.units + '</td><td class="addr">' + esc(l.address) + '</td></tr>';
    });
    $('#panel-lines').innerHTML = table(['고유키', '채널', '수취인', '원본 옵션', '출고 상품', '주문수량', '구성', '출고수량', '주소'], rows,
      '정규화된 주문이 없습니다.');
  }

  function render() {
    var v = compute();
    app.view = v;
    renderLogs();
    renderKpis(v.summary);
    renderShipments(v);
    renderExceptions(v);
    renderTracking(v);
    renderLines(v);
    var hasShip = v.built.shipments.length > 0;
    $('#dl-po').disabled = !hasShip;
    $('#btn-simulate').disabled = !hasShip;
    $('#dl-a').disabled = !(v.match && v.match.matched.some(function (x) { return x.line.channel === 'A'; }));
    $('#dl-b').disabled = !(v.match && v.match.matched.some(function (x) { return x.line.channel === 'B'; }));
    $('#dl-hold').disabled = !(v.built.exceptions.length || (v.match && (v.match.held.length || v.match.unanswered.length)));
  }

  // ------------------------------------------------------------------ 이벤트
  function bindDrop(zoneSel, inputSel) {
    var zone = $(zoneSel);
    var input = $(inputSel);
    input.addEventListener('change', function () { handleFileList(input.files).then(function () { input.value = ''; }); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('dragover'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove('dragover'); });
    });
    zone.addEventListener('drop', function (e) { handleFileList(e.dataTransfer.files); });
  }

  function selectTab(name) {
    document.querySelectorAll('#tabs button').forEach(function (b) {
      b.setAttribute('aria-selected', String(b.dataset.tab === name));
    });
    ['shipments', 'exceptions', 'tracking', 'lines'].forEach(function (t) { $('#panel-' + t).hidden = t !== name; });
  }

  function init() {
    if (!window.XLSX) {
      toast('Excel 라이브러리(SheetJS)를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');
    }
    bindDrop('#drop-orders', '#input-orders');
    bindDrop('#drop-reply', '#input-reply');
    $('#btn-demo').addEventListener('click', loadSamples);
    $('#btn-reset').addEventListener('click', reset);
    $('#btn-simulate').addEventListener('click', simulateReply);
    $('#dl-po').addEventListener('click', downloadPurchaseOrder);
    $('#dl-a').addEventListener('click', function () { downloadUpload('A'); });
    $('#dl-b').addEventListener('click', function () { downloadUpload('B'); });
    $('#dl-hold').addEventListener('click', downloadHold);
    $('#tabs').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-tab]');
      if (b) selectTab(b.dataset.tab);
    });
    render();

    // README 캡처·공유용: ?demo=1 이면 샘플 자동 로드, &tab=exceptions 처럼 탭 지정
    var params = new URLSearchParams(location.search);
    if (params.get('tab')) selectTab(params.get('tab'));
    if (params.get('demo') === '1' && window.XLSX) loadSamples();
  }

  init();
})();
