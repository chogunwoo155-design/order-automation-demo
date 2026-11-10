/**
 * 주문 → 발주 → 송장 변환 핵심 로직 (순수 함수 모음)
 *
 * - DOM, 파일 I/O, SheetJS에 의존하지 않는다. 입력은 "행 배열(array of arrays)"이고
 *   출력도 행 배열 또는 일반 객체다. 그래서 브라우저(<script>)와 Node(require) 양쪽에서
 *   같은 코드를 그대로 실행하고 테스트할 수 있다.
 * - 원본 주문 데이터는 변경하지 않는다. 자동 보정(예: 수취인명 대체)은 출고 단계의
 *   값에만 반영하고, 원본 값은 그대로 남긴다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OrderCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // 설정: 마켓 양식, 상품 카탈로그, 박스 규칙
  // ---------------------------------------------------------------------------

  /** 가상 마켓 두 곳의 주문 Excel 양식. 필드명 → 원본 헤더명 */
  var MARKETS = {
    A: {
      code: 'A',
      label: 'A마켓',
      columns: {
        lineNo: '상품주문번호',
        orderNo: '주문번호',
        orderedAt: '주문일시',
        buyer: '구매자명',
        recipient: '수취인명',
        phone: '수취인연락처1',
        address: '통합배송지',
        zip: '우편번호',
        productName: '상품명',
        option: '옵션정보',
        qty: '수량',
        message: '배송메세지'
      },
      required: ['상품주문번호', '주문번호', '수취인명', '통합배송지', '옵션정보', '수량']
    },
    B: {
      code: 'B',
      label: 'B마켓',
      columns: {
        bundleNo: '묶음배송번호',
        orderNo: '주문번호',
        lineNo: '옵션ID',
        orderedAt: '주문일',
        buyer: '구매자',
        productName: '등록상품명',
        option: '등록옵션명',
        qty: '구매수(수량)',
        recipient: '수취인이름',
        phone: '수취인전화번호',
        zip: '우편번호',
        address1: '수취인 주소',
        address2: '수취인 상세주소',
        message: '배송메세지'
      },
      required: ['묶음배송번호', '주문번호', '옵션ID', '수취인이름', '수취인 주소', '구매수(수량)']
    }
  };

  /** 가상 물류사 송장 회신 양식 */
  var REPLY_COLUMNS = {
    ref: '발주번호',
    carrier: '택배사',
    trackingNo: '운송장번호',
    recipient: '수취인'
  };

  /** 옵션 문구 → 실제 출고 상품(SKU). 키워드는 위에서부터 검사한다. */
  var CATALOG = [
    { sku: 'JUICE-APL', name: '사과즙', keywords: ['사과'] },
    { sku: 'JUICE-GRP', name: '포도즙', keywords: ['포도'] },
    { sku: 'JUICE-PEAR', name: '배즙', keywords: ['배'] }
  ];

  var BOX_TYPES = {
    SINGLE: { code: 'SINGLE', label: '싱글팩', capacity: 1 },
    DOUBLE: { code: 'DOUBLE', label: '더블팩', capacity: 2 },
    SIX: { code: 'SIX', label: '6구 박스', capacity: 6 }
  };
  var MAX_UNITS_PER_BOX = 6;

  var ISSUE_LABELS = {
    MISSING_KEY: '주문키 누락',
    ADDRESS_MISSING: '주소 누락',
    RECIPIENT_MISSING: '수취인 누락',
    RECIPIENT_SHORT: '수취인 한 글자',
    RECIPIENT_REPLACED: '수취인명 대체',
    QTY_INVALID: '수량 오류',
    PRODUCT_UNMAPPED: '상품 매핑 실패',
    PHONE_INVALID: '연락처 확인',
    ZIP_INVALID: '우편번호 확인'
  };

  var HOLD_LABELS = {
    UNKNOWN_REF: '발주번호 없음',
    INVALID_TRACKING: '송장번호 형식 오류',
    RECIPIENT_MISMATCH: '수취인 불일치',
    CONFLICT: '송장번호 충돌',
    DUP_TRACKING: '송장번호 중복 사용'
  };

  // ---------------------------------------------------------------------------
  // 문자열 유틸
  // ---------------------------------------------------------------------------

  function str(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }
  function cleanSpaces(v) {
    return str(v).replace(/\s+/g, ' ');
  }
  function digits(v) {
    return str(v).replace(/\D/g, '');
  }
  function charLength(v) {
    return Array.from(str(v)).length;
  }
  /** 합포장 판단용 주소 정규화: 공백·쉼표 차이는 같은 주소로 본다 */
  function addressKey(v) {
    return str(v).replace(/[\s,]+/g, '');
  }

  /** FNV-1a 32bit. 같은 입력이면 항상 같은 출고번호가 나오도록 하는 결정적 해시 */
  function hash(text) {
    var h = 0x811c9dc5;
    var s = String(text);
    for (var i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 0x01000193) >>> 0;
    }
    return ('00000000' + h.toString(16).toUpperCase()).slice(-8);
  }

  // ---------------------------------------------------------------------------
  // 1) 양식 감지 · 파싱 · 정규화
  // ---------------------------------------------------------------------------

  function findHeader(rows, names, scanLimit) {
    var limit = Math.min(rows.length, scanLimit || 10);
    for (var i = 0; i < limit; i++) {
      var cells = (rows[i] || []).map(cleanSpaces);
      var ok = names.every(function (n) { return cells.indexOf(n) !== -1; });
      if (ok) return { index: i, cells: cells };
    }
    return null;
  }

  /**
   * 업로드된 시트가 어떤 양식인지 판별한다. 헤더가 첫 줄이 아닐 수 있으므로
   * 상단 10행을 훑어서 필수 헤더가 모두 있는 행을 찾는다.
   * @returns {{type: 'A'|'B'|'REPLY'|null, headerIndex: number}}
   */
  function detectFormat(rows) {
    var keys = Object.keys(MARKETS);
    for (var k = 0; k < keys.length; k++) {
      var h = findHeader(rows, MARKETS[keys[k]].required);
      if (h) return { type: keys[k], headerIndex: h.index };
    }
    var r = findHeader(rows, [REPLY_COLUMNS.ref, REPLY_COLUMNS.trackingNo]);
    if (r) return { type: 'REPLY', headerIndex: r.index };
    return { type: null, headerIndex: -1 };
  }

  /** 헤더 아래 행들을 { 필드: 값 } 레코드로 바꾼다. 빈 행은 건너뛴다. */
  function toRecords(rows, headerIndex, columns) {
    var header = (rows[headerIndex] || []).map(cleanSpaces);
    var idx = {};
    Object.keys(columns).forEach(function (field) {
      idx[field] = header.indexOf(columns[field]);
    });
    var out = [];
    for (var i = headerIndex + 1; i < rows.length; i++) {
      var row = rows[i] || [];
      if (!row.some(function (c) { return str(c) !== ''; })) continue;
      var rec = { _row: i + 1 }; // Excel 기준 행 번호
      Object.keys(idx).forEach(function (field) {
        rec[field] = idx[field] >= 0 ? str(row[idx[field]]) : '';
      });
      out.push(rec);
    }
    return out;
  }

  /**
   * 옵션 문구에서 실제 상품과 1구성당 개수를 뽑는다.
   *   "맛: 사과 / 구성: 2개 세트" → { sku: JUICE-APL, packSize: 2 }
   *   "배즙 x 3입"               → { sku: JUICE-PEAR, packSize: 3 }
   * 상품을 찾지 못하면 null.
   */
  function parseOption(text) {
    var t = cleanSpaces(text);
    var item = null;
    for (var i = 0; i < CATALOG.length && !item; i++) {
      var c = CATALOG[i];
      if (c.keywords.some(function (k) { return t.indexOf(k) !== -1; })) item = c;
    }
    if (!item) return null;
    var m = t.match(/(\d+)\s*(?:개\s*세트|개입|개|입)/);
    var packSize = m ? parseInt(m[1], 10) : 1;
    return { sku: item.sku, skuName: item.name, packSize: packSize > 0 ? packSize : 1 };
  }

  function parseQty(v) {
    var s = str(v).replace(/,/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
    return Number(s);
  }

  /** 마켓별 레코드 → 공통 주문 라인 구조 */
  function normalizeRecord(market, rec, fileName) {
    var opt = parseOption(rec.option);
    var qty = parseQty(rec.qty);
    var qtyOk = Number.isInteger(qty) && qty > 0;
    var address = market === 'B'
      ? cleanSpaces([rec.address1, rec.address2].filter(Boolean).join(' '))
      : cleanSpaces(rec.address);
    var orderNo = str(rec.orderNo);
    var lineNo = str(rec.lineNo);

    // 고유키: 판매채널 + 주문번호 + 상품(라인) 식별값.
    // 같은 파일을 다시 올려도 같은 키가 나오므로 중복 집계되지 않는다.
    var uid = orderNo && lineNo
      ? market + ':' + orderNo + ':' + lineNo
      : market + ':NOKEY:' + hash([rec._row, rec.recipient, rec.option, rec.qty, address].join('|'));

    return {
      uid: uid,
      channel: market,
      channelLabel: MARKETS[market].label,
      orderNo: orderNo,
      lineNo: lineNo,
      bundleNo: str(rec.bundleNo),
      orderedAt: str(rec.orderedAt),
      buyer: cleanSpaces(rec.buyer),
      recipient: cleanSpaces(rec.recipient),
      phone: str(rec.phone),
      zip: str(rec.zip),
      address: address,
      productName: cleanSpaces(rec.productName),
      option: cleanSpaces(rec.option),
      sku: opt ? opt.sku : '',
      skuName: opt ? opt.skuName : '',
      packSize: opt ? opt.packSize : 0,
      qty: qtyOk ? qty : 0,
      qtyRaw: str(rec.qty),
      units: opt && qtyOk ? qty * opt.packSize : 0,
      message: cleanSpaces(rec.message),
      source: { file: fileName || '', row: rec._row }
    };
  }

  /**
   * 주문 시트(행 배열)를 공통 주문 라인 배열로 변환한다.
   * @returns {{type: string|null, lines: object[], error?: string}}
   */
  function parseOrderRows(rows, opts) {
    var fileName = (opts && opts.fileName) || '';
    var fmt = detectFormat(rows);
    if (fmt.type !== 'A' && fmt.type !== 'B') {
      return { type: fmt.type, lines: [], error: fmt.type === 'REPLY' ? '송장 회신 파일입니다' : '알 수 없는 주문 양식입니다' };
    }
    var recs = toRecords(rows, fmt.headerIndex, MARKETS[fmt.type].columns);
    return {
      type: fmt.type,
      lines: recs.map(function (r) { return normalizeRecord(fmt.type, r, fileName); })
    };
  }

  // ---------------------------------------------------------------------------
  // 2) 누적 · 중복 제거
  // ---------------------------------------------------------------------------

  function createState() {
    return { byUid: {}, order: [] };
  }

  /** 기존 상태에 주문 라인을 더한다. 같은 고유키는 무시한다. (새 상태를 반환) */
  function ingest(state, lines) {
    var byUid = Object.assign({}, state.byUid);
    var order = state.order.slice();
    var added = 0;
    var duplicates = 0;
    lines.forEach(function (l) {
      if (byUid[l.uid]) { duplicates++; return; }
      byUid[l.uid] = l;
      order.push(l.uid);
      added++;
    });
    return { state: { byUid: byUid, order: order }, added: added, duplicates: duplicates };
  }

  function allLines(state) {
    return state.order.map(function (u) { return state.byUid[u]; });
  }

  // ---------------------------------------------------------------------------
  // 3) 검증 · 합포장 · 박스 규칙
  // ---------------------------------------------------------------------------

  function issue(code, level, message) {
    return { code: code, label: ISSUE_LABELS[code], level: level, message: message };
  }

  /**
   * 출고 가능 여부를 검사한다. error가 하나라도 있으면 출고에서 제외,
   * warning은 출고하되 확인 목록에 남긴다.
   */
  function validateLine(line) {
    var issues = [];
    var shipRecipient = line.recipient;

    if (!line.orderNo || !line.lineNo) {
      issues.push(issue('MISSING_KEY', 'error', '주문번호 또는 상품 식별값이 비어 있어 중복 검사를 할 수 없음'));
    }
    if (charLength(line.address) < 5) {
      issues.push(issue('ADDRESS_MISSING', 'error', '배송지 주소가 비어 있거나 너무 짧음'));
    }
    if (!line.recipient) {
      issues.push(issue('RECIPIENT_MISSING', 'error', '수취인명이 비어 있음'));
    } else if (charLength(line.recipient) < 2) {
      if (charLength(line.buyer) >= 2) {
        shipRecipient = line.buyer;
        issues.push(issue('RECIPIENT_REPLACED', 'warning',
          '수취인명 "' + line.recipient + '"(1글자) → 구매자명 "' + line.buyer + '"으로 발주 (원본 유지)'));
      } else {
        issues.push(issue('RECIPIENT_SHORT', 'error',
          '수취인명 "' + line.recipient + '"이 1글자이고 대체할 구매자명이 없음 (물류사 최소 2글자)'));
      }
    }
    if (!line.qty) {
      issues.push(issue('QTY_INVALID', 'error', '수량 "' + line.qtyRaw + '"이 1 이상의 정수가 아님'));
    }
    if (!line.sku) {
      issues.push(issue('PRODUCT_UNMAPPED', 'error', '옵션 "' + line.option + '"을 출고 상품으로 매핑할 수 없음'));
    }
    if (!/^0\d{9,10}$/.test(digits(line.phone))) {
      issues.push(issue('PHONE_INVALID', 'warning', '연락처 "' + line.phone + '" 형식 확인 필요'));
    }
    if (!/^\d{5}$/.test(line.zip)) {
      issues.push(issue('ZIP_INVALID', 'warning', '우편번호 "' + line.zip + '" 형식 확인 필요'));
    }

    var blocked = issues.some(function (i) { return i.level === 'error'; });
    return { ok: !blocked, issues: issues, shipRecipient: shipRecipient };
  }

  function boxTypeFor(units) {
    if (units <= 1) return BOX_TYPES.SINGLE;
    if (units === 2) return BOX_TYPES.DOUBLE;
    return BOX_TYPES.SIX;
  }

  /**
   * 수량별 박스 규칙.
   *   1개 → 싱글팩, 2개 → 더블팩, 3~6개 → 6구 박스,
   *   7개 이상 → 6개씩 6구 박스로 나누고 나머지는 같은 규칙으로 포장.
   * @returns {{type: object, units: number}[]}
   */
  function planBoxes(totalUnits) {
    var n = Math.floor(Number(totalUnits) || 0);
    if (n <= 0) return [];
    var sizes = [];
    while (n > MAX_UNITS_PER_BOX) {
      sizes.push(MAX_UNITS_PER_BOX);
      n -= MAX_UNITS_PER_BOX;
    }
    sizes.push(n);
    return sizes.map(function (u) { return { type: boxTypeFor(u), units: u }; });
  }

  /** 박스 계획에 따라 실제 내품을 채운다. 어떤 주문 라인이 어느 박스에 들어갔는지도 기록한다. */
  function packItems(lines, plan, shipmentId) {
    var unitsQueue = [];
    lines.forEach(function (l) {
      for (var i = 0; i < l.units; i++) unitsQueue.push(l);
    });
    var cursor = 0;
    return plan.map(function (p, idx) {
      var slice = unitsQueue.slice(cursor, cursor + p.units);
      cursor += p.units;
      var bySku = {};
      var skuOrder = [];
      var uids = [];
      slice.forEach(function (l) {
        if (!bySku[l.sku]) { bySku[l.sku] = { sku: l.sku, name: l.skuName, units: 0 }; skuOrder.push(l.sku); }
        bySku[l.sku].units++;
        if (uids.indexOf(l.uid) === -1) uids.push(l.uid);
      });
      return {
        ref: shipmentId + '-' + (idx + 1),
        seq: idx + 1,
        of: plan.length,
        type: p.type.code,
        typeLabel: p.type.label,
        units: p.units,
        contents: skuOrder.map(function (s) { return bySku[s]; }),
        lineUids: uids
      };
    });
  }

  /**
   * 주문 라인 → 출고(합포장) 단위로 묶고 박스를 계획한다.
   * 합포장 기준: 같은 판매채널 · 같은 수취인 · 같은 주소(공백/쉼표 무시).
   */
  function buildShipments(lines) {
    var exceptions = [];
    var groups = {};
    var groupOrder = [];

    lines.forEach(function (line) {
      var v = validateLine(line);
      v.issues.forEach(function (i) {
        exceptions.push({
          uid: line.uid,
          channel: line.channel,
          channelLabel: line.channelLabel,
          orderNo: line.orderNo,
          lineNo: line.lineNo,
          recipient: line.recipient,
          option: line.option,
          code: i.code,
          label: i.label,
          level: i.level,
          message: i.message,
          source: line.source
        });
      });
      if (!v.ok) return;

      var key = [line.channel, v.shipRecipient, addressKey(line.address)].join('|');
      if (!groups[key]) {
        groups[key] = {
          id: 'SH-' + hash(key),
          key: key,
          channel: line.channel,
          channelLabel: line.channelLabel,
          recipient: v.shipRecipient,
          phone: line.phone,
          zip: line.zip,
          address: line.address,
          messages: [],
          lines: []
        };
        groupOrder.push(key);
      }
      var g = groups[key];
      g.lines.push(line);
      if (line.message && g.messages.indexOf(line.message) === -1) g.messages.push(line.message);
    });

    var shipments = groupOrder.map(function (key) {
      var g = groups[key];
      var totalUnits = g.lines.reduce(function (s, l) { return s + l.units; }, 0);
      var plan = planBoxes(totalUnits);
      var orderNos = [];
      g.lines.forEach(function (l) { if (orderNos.indexOf(l.orderNo) === -1) orderNos.push(l.orderNo); });
      return {
        id: g.id,
        key: g.key,
        channel: g.channel,
        channelLabel: g.channelLabel,
        recipient: g.recipient,
        phone: g.phone,
        zip: g.zip,
        address: g.address,
        message: g.messages.join(' / '),
        lines: g.lines,
        orderNos: orderNos,
        merged: g.lines.length > 1,
        totalUnits: totalUnits,
        boxes: packItems(g.lines, plan, g.id)
      };
    });

    return { shipments: shipments, exceptions: exceptions };
  }

  // ---------------------------------------------------------------------------
  // 4) 물류사 발주서
  // ---------------------------------------------------------------------------

  var PURCHASE_ORDER_HEADER = ['발주번호', '박스', '박스유형', '수취인', '연락처', '우편번호', '주소',
    '배송메시지', '내품', '내품수량', '판매채널', '원주문번호'];

  function contentsText(contents) {
    return contents.map(function (c) { return c.name + ' x' + c.units; }).join(', ');
  }

  /** 박스 1개 = 발주서 1행 */
  function toPurchaseOrderRows(shipments) {
    var rows = [PURCHASE_ORDER_HEADER.slice()];
    shipments.forEach(function (s) {
      s.boxes.forEach(function (b) {
        rows.push([
          b.ref, b.seq + '/' + b.of, b.typeLabel, s.recipient, s.phone, s.zip, s.address,
          s.message, contentsText(b.contents), b.units, s.channelLabel, s.orderNos.join(', ')
        ]);
      });
    });
    return rows;
  }

  // ---------------------------------------------------------------------------
  // 5) 송장 회신 매칭 · 마켓 업로드 파일
  // ---------------------------------------------------------------------------

  function parseReplyRows(rows) {
    var fmt = detectFormat(rows);
    if (fmt.type !== 'REPLY') return { records: [], error: '송장 회신 양식이 아닙니다' };
    var recs = toRecords(rows, fmt.headerIndex, REPLY_COLUMNS).map(function (r) {
      return {
        ref: str(r.ref),
        carrier: cleanSpaces(r.carrier),
        trackingNo: str(r.trackingNo),
        recipient: cleanSpaces(r.recipient),
        row: r._row
      };
    });
    return { records: recs };
  }

  function replyKey(r) {
    return [r.ref, digits(r.trackingNo), r.recipient].join('|');
  }

  /** 회신 레코드 누적. 같은 회신 파일을 다시 올려도 중복되지 않는다. */
  function mergeReplies(existing, incoming) {
    var seen = {};
    existing.forEach(function (r) { seen[replyKey(r)] = true; });
    var out = existing.slice();
    var added = 0;
    var duplicates = 0;
    incoming.forEach(function (r) {
      var k = replyKey(r);
      if (seen[k]) { duplicates++; return; }
      seen[k] = true;
      out.push(r);
      added++;
    });
    return { records: out, added: added, duplicates: duplicates };
  }

  function hold(r, reason, message) {
    return { ref: r.ref, trackingNo: r.trackingNo, recipient: r.recipient, row: r.row,
      reason: reason, label: HOLD_LABELS[reason], message: message };
  }

  /**
   * 송장 회신을 발주 박스와 맞춰 보고, 주문 라인별 송장번호를 정한다.
   * 발주번호·수취인이 모두 맞는 회신만 채택하고, 애매한 건은 보류 목록으로 보낸다.
   */
  function matchTracking(shipments, replies) {
    var boxIndex = {};
    shipments.forEach(function (s) {
      s.boxes.forEach(function (b) { boxIndex[b.ref] = { shipment: s, box: b }; });
    });

    var held = [];
    var accepted = {};      // ref → { trackingNo, carrier }
    var refByTracking = {}; // trackingNo → ref
    var conflicted = {};

    replies.forEach(function (r) {
      var hit = boxIndex[r.ref];
      if (!hit) { held.push(hold(r, 'UNKNOWN_REF', '발주서에 없는 발주번호')); return; }
      var tn = digits(r.trackingNo);
      if (!/^\d{10,14}$/.test(tn) || digits(r.trackingNo).length !== str(r.trackingNo).replace(/-/g, '').length) {
        held.push(hold(r, 'INVALID_TRACKING', '송장번호 "' + r.trackingNo + '" 형식 오류 (숫자 10~14자리)'));
        return;
      }
      if (r.recipient && addressKey(r.recipient) !== addressKey(hit.shipment.recipient)) {
        held.push(hold(r, 'RECIPIENT_MISMATCH',
          '수취인 불일치 (발주: ' + hit.shipment.recipient + ' / 회신: ' + r.recipient + ')'));
        return;
      }
      var prev = accepted[r.ref];
      if (prev && prev.trackingNo !== tn) {
        conflicted[r.ref] = true;
        held.push(hold(r, 'CONFLICT', '같은 발주번호에 서로 다른 송장번호가 회신됨 (' + prev.trackingNo + ')'));
        return;
      }
      var otherRef = refByTracking[tn];
      if (otherRef && otherRef !== r.ref) {
        held.push(hold(r, 'DUP_TRACKING', '이미 ' + otherRef + '에 쓰인 송장번호'));
        return;
      }
      accepted[r.ref] = { trackingNo: tn, carrier: r.carrier };
      refByTracking[tn] = r.ref;
    });
    Object.keys(conflicted).forEach(function (ref) { delete accepted[ref]; });

    var matched = [];
    var unanswered = [];
    var unansweredBoxes = [];
    shipments.forEach(function (s) {
      s.boxes.forEach(function (b) {
        if (!accepted[b.ref]) unansweredBoxes.push({ ref: b.ref, shipmentId: s.id, recipient: s.recipient });
      });
      s.lines.forEach(function (l) {
        var box = null;
        for (var i = 0; i < s.boxes.length && !box; i++) {
          var b = s.boxes[i];
          if (accepted[b.ref] && b.lineUids.indexOf(l.uid) !== -1) box = b;
        }
        if (box) {
          matched.push({ line: l, shipmentId: s.id, ref: box.ref,
            carrier: accepted[box.ref].carrier, trackingNo: accepted[box.ref].trackingNo });
        } else {
          unanswered.push({ line: l, shipmentId: s.id });
        }
      });
    });

    return { matched: matched, held: held, unanswered: unanswered, unansweredBoxes: unansweredBoxes };
  }

  var UPLOAD_HEADERS = {
    A: ['상품주문번호', '배송방법', '택배사', '송장번호'],
    B: ['묶음배송번호', '주문번호', '옵션ID', '택배사', '운송장번호']
  };

  /** 매칭 결과 → 마켓별 송장 업로드 파일(행 배열). 양식이 서로 다르다. */
  function toMarketUploadRows(matched, market) {
    var rows = [UPLOAD_HEADERS[market].slice()];
    matched.forEach(function (m) {
      if (m.line.channel !== market) return;
      if (market === 'A') rows.push([m.line.lineNo, '택배', m.carrier, m.trackingNo]);
      else rows.push([m.line.bundleNo, m.line.orderNo, m.line.lineNo, m.carrier, m.trackingNo]);
    });
    return rows;
  }

  /** 데모용: 현재 발주서 기준으로 가상 물류사 회신을 만든다. (실제 물류사 연동 대체) */
  function simulateReplyRows(shipments, carrier) {
    var rows = [[REPLY_COLUMNS.ref, REPLY_COLUMNS.carrier, REPLY_COLUMNS.trackingNo, REPLY_COLUMNS.recipient]];
    shipments.forEach(function (s) {
      s.boxes.forEach(function (b) {
        var n = (parseInt(hash('trk:' + b.ref), 16) % 1e9).toString();
        rows.push([b.ref, carrier || '가상택배', '700' + ('000000000' + n).slice(-9), s.recipient]);
      });
    });
    return rows;
  }

  // ---------------------------------------------------------------------------
  // 6) 대시보드 집계
  // ---------------------------------------------------------------------------

  function summarize(lines, built, match) {
    var orderKeys = {};
    var byChannel = {};
    lines.forEach(function (l) {
      orderKeys[l.channel + ':' + l.orderNo] = true;
      byChannel[l.channel] = byChannel[l.channel] || { lines: 0, orders: {} };
      byChannel[l.channel].lines++;
      byChannel[l.channel].orders[l.orderNo] = true;
    });
    Object.keys(byChannel).forEach(function (c) {
      byChannel[c] = { lines: byChannel[c].lines, orders: Object.keys(byChannel[c].orders).length };
    });

    var boxByType = { SINGLE: 0, DOUBLE: 0, SIX: 0 };
    var boxCount = 0;
    var unitCount = 0;
    var mergedShipments = 0;
    var shippableLines = 0;
    built.shipments.forEach(function (s) {
      if (s.merged) mergedShipments++;
      shippableLines += s.lines.length;
      unitCount += s.totalUnits;
      s.boxes.forEach(function (b) { boxByType[b.type]++; boxCount++; });
    });

    var blockedUids = {};
    var warningCount = 0;
    built.exceptions.forEach(function (e) {
      if (e.level === 'error') blockedUids[e.uid] = true;
      else warningCount++;
    });

    var result = {
      orderCount: Object.keys(orderKeys).length,
      lineCount: lines.length,
      shippableLines: shippableLines,
      blockedLines: Object.keys(blockedUids).length,
      warningCount: warningCount,
      shipmentCount: built.shipments.length,
      mergedShipments: mergedShipments,
      boxCount: boxCount,
      boxByType: boxByType,
      unitCount: unitCount,
      byChannel: byChannel
    };
    if (match) {
      result.tracking = {
        matched: match.matched.length,
        held: match.held.length,
        unanswered: match.unanswered.length,
        unansweredBoxes: match.unansweredBoxes.length
      };
    }
    return result;
  }

  return {
    MARKETS: MARKETS,
    REPLY_COLUMNS: REPLY_COLUMNS,
    CATALOG: CATALOG,
    BOX_TYPES: BOX_TYPES,
    ISSUE_LABELS: ISSUE_LABELS,
    HOLD_LABELS: HOLD_LABELS,
    PURCHASE_ORDER_HEADER: PURCHASE_ORDER_HEADER,
    UPLOAD_HEADERS: UPLOAD_HEADERS,
    hash: hash,
    detectFormat: detectFormat,
    parseOption: parseOption,
    parseOrderRows: parseOrderRows,
    normalizeRecord: normalizeRecord,
    createState: createState,
    ingest: ingest,
    allLines: allLines,
    validateLine: validateLine,
    planBoxes: planBoxes,
    buildShipments: buildShipments,
    toPurchaseOrderRows: toPurchaseOrderRows,
    parseReplyRows: parseReplyRows,
    mergeReplies: mergeReplies,
    matchTracking: matchTracking,
    toMarketUploadRows: toMarketUploadRows,
    simulateReplyRows: simulateReplyRows,
    summarize: summarize
  };
});
