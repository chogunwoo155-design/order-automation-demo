'use strict';
// 실행: node --test tests/
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const C = require('../js/core.js');

const fixture = (name) =>
  JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', name), 'utf8'));
const A_ROWS = fixture('a-market.json');
const B_ROWS = fixture('b-market.json');
const REPLY_ROWS = fixture('reply.json');

function loadAll(order = ['A', 'B']) {
  let state = C.createState();
  for (const m of order) {
    const rows = m === 'A' ? A_ROWS : B_ROWS;
    state = C.ingest(state, C.parseOrderRows(rows, { fileName: m }).lines).state;
  }
  const lines = C.allLines(state);
  return { state, lines, built: C.buildShipments(lines) };
}
const byRecipient = (shipments, channel, name) =>
  shipments.find((s) => s.channel === channel && s.recipient === name);

test('박스 규칙: 1개 싱글팩, 2개 더블팩, 3~6개 6구, 7개 이상 분할', () => {
  const types = (n) => C.planBoxes(n).map((b) => `${b.type.code}:${b.units}`);
  assert.deepEqual(types(0), []);
  assert.deepEqual(types(1), ['SINGLE:1']);
  assert.deepEqual(types(2), ['DOUBLE:2']);
  assert.deepEqual(types(3), ['SIX:3']);
  assert.deepEqual(types(6), ['SIX:6']);
  assert.deepEqual(types(7), ['SIX:6', 'SINGLE:1']);
  assert.deepEqual(types(8), ['SIX:6', 'DOUBLE:2']);
  assert.deepEqual(types(12), ['SIX:6', 'SIX:6']);
  assert.deepEqual(types(14), ['SIX:6', 'SIX:6', 'DOUBLE:2']);
});

test('옵션 문구 → 상품·구성 수량 분리 (마켓별 표기 차이 흡수)', () => {
  assert.deepEqual(C.parseOption('맛: 포도 / 구성: 2개 세트'), { sku: 'JUICE-GRP', skuName: '포도즙', packSize: 2 });
  assert.deepEqual(C.parseOption('배즙 x 3입'), { sku: 'JUICE-PEAR', skuName: '배즙', packSize: 3 });
  assert.deepEqual(C.parseOption('맛: 사과 / 구성: 1개'), { sku: 'JUICE-APL', skuName: '사과즙', packSize: 1 });
  assert.equal(C.parseOption('맛: 망고 / 구성: 1개'), null);
});

test('양식 자동 감지: 헤더가 첫 줄이 아니어도 찾는다', () => {
  assert.deepEqual(C.detectFormat(A_ROWS), { type: 'A', headerIndex: 1 });
  assert.deepEqual(C.detectFormat(B_ROWS), { type: 'B', headerIndex: 0 });
  assert.equal(C.detectFormat(REPLY_ROWS).type, 'REPLY');
  assert.equal(C.detectFormat([['아무', '헤더']]).type, null);
});

test('두 마켓 양식이 같은 공통 구조로 정규화된다', () => {
  const a = C.parseOrderRows(A_ROWS).lines;
  const b = C.parseOrderRows(B_ROWS).lines;
  assert.equal(a.length, 13);
  assert.equal(b.length, 9);
  const keys = Object.keys(a[0]).sort();
  assert.deepEqual(Object.keys(b[0]).sort(), keys);

  const pa5 = a.find((l) => l.lineNo === 'PA-0005');
  assert.equal(pa5.uid, 'A:A-1004:PA-0005');
  assert.equal(pa5.units, 9); // 3개 세트 x 3
  assert.equal(pa5.source.row, 7); // 안내행 + 헤더 다음 5번째 데이터 → Excel 7행

  const b3 = b.find((l) => l.orderNo === 'B-5003');
  assert.equal(b3.units, 14); // 2입 x 7
  assert.equal(b3.address, '가상시 가짜구 샘플길 2 B동 5층'); // 주소 + 상세주소 결합
});

test('같은 파일을 다시 올려도 중복 집계되지 않는다 (채널+주문번호+상품 고유키)', () => {
  let state = C.createState();
  const lines = C.parseOrderRows(A_ROWS).lines;
  const first = C.ingest(state, lines);
  assert.equal(first.added, 13);
  assert.equal(first.duplicates, 0);
  const second = C.ingest(first.state, C.parseOrderRows(A_ROWS).lines);
  assert.equal(second.added, 0);
  assert.equal(second.duplicates, 13);
  assert.equal(C.allLines(second.state).length, 13);
  // ingest는 원래 상태를 바꾸지 않는다
  assert.equal(C.allLines(first.state).length, 13);
  assert.equal(state.order.length, 0);
});

test('합포장: 같은 채널·수취인·주소는 주문이 달라도 한 출고로 묶는다', () => {
  const { built } = loadAll();
  const choi = byRecipient(built.shipments, 'A', '최테스트');
  assert.deepEqual(choi.orderNos, ['A-1005', 'A-1006']);
  assert.equal(choi.totalUnits, 3);
  assert.deepEqual(choi.boxes.map((b) => b.type), ['SIX']);

  // 주소 공백·쉼표 차이("가상시  예제구 모의로 7,")는 같은 주소로 본다
  const seo = byRecipient(built.shipments, 'B', '서예제');
  assert.equal(seo.lines.length, 2);
  assert.deepEqual(seo.boxes.map((b) => b.type), ['DOUBLE']);

  // 같은 사람·주소라도 판매채널이 다르면 송장 업로드 대상이 달라 따로 출고
  const hongs = built.shipments.filter((s) => s.recipient === '홍길동');
  assert.equal(hongs.length, 2);
});

test('박스 분할: 내품 합계가 출고 수량과 같고, 라인이 어느 박스에 들어갔는지 추적된다', () => {
  const { built } = loadAll();
  const han = byRecipient(built.shipments, 'B', '한가짜');
  assert.deepEqual(han.boxes.map((b) => `${b.typeLabel}:${b.units}`), ['6구 박스:6', '6구 박스:6', '더블팩:2']);
  assert.deepEqual(han.boxes.map((b) => b.ref), [`${han.id}-1`, `${han.id}-2`, `${han.id}-3`]);
  for (const s of built.shipments) {
    const sum = s.boxes.reduce((n, b) => n + b.contents.reduce((m, c) => m + c.units, 0), 0);
    assert.equal(sum, s.totalUnits, s.id);
  }
  const jo = byRecipient(built.shipments, 'B', '조허구');
  assert.deepEqual(jo.boxes[0].contents, [
    { sku: 'JUICE-PEAR', name: '배즙', units: 3 },
    { sku: 'JUICE-GRP', name: '포도즙', units: 1 }
  ]);
});

test('예외: 주소 누락·수취인 한 글자·수량 오류·상품 매핑 실패는 출고 제외, 경고는 출고하되 표시', () => {
  const { built } = loadAll();
  const codes = (orderNo) => built.exceptions.filter((e) => e.orderNo === orderNo).map((e) => e.code).sort();
  assert.deepEqual(codes('A-1008'), ['RECIPIENT_SHORT']);
  assert.deepEqual(codes('A-1009'), ['ADDRESS_MISSING', 'ZIP_INVALID']);
  assert.deepEqual(codes('A-1010'), ['PRODUCT_UNMAPPED']);
  assert.deepEqual(codes('A-1011'), ['PHONE_INVALID']);
  assert.deepEqual(codes('B-5006'), ['QTY_INVALID']);
  assert.deepEqual(codes('B-5007'), ['ADDRESS_MISSING']);

  // 수취인 "정" → 구매자명 "정가상"으로 발주, 원본 라인은 그대로
  assert.deepEqual(codes('A-1007'), ['RECIPIENT_REPLACED']);
  const jung = built.shipments.find((s) => s.orderNos.includes('A-1007'));
  assert.equal(jung.recipient, '정가상');
  assert.equal(jung.lines[0].recipient, '정');

  const blocked = new Set(built.exceptions.filter((e) => e.level === 'error').map((e) => e.orderNo));
  for (const s of built.shipments) for (const o of s.orderNos) assert.ok(!blocked.has(o), o);
});

test('출고번호는 업로드 순서와 무관하게 같다 (결정적 키)', () => {
  const ab = loadAll(['A', 'B']).built.shipments.map((s) => s.id).sort();
  const ba = loadAll(['B', 'A']).built.shipments.map((s) => s.id).sort();
  assert.deepEqual(ab, ba);
});

test('물류사 발주서: 박스 1개당 1행', () => {
  const { built } = loadAll();
  const rows = C.toPurchaseOrderRows(built.shipments);
  assert.deepEqual(rows[0], C.PURCHASE_ORDER_HEADER);
  assert.equal(rows.length - 1, 17);
  const bae = rows.filter((r) => r[3] === '배가짜');
  assert.deepEqual(bae.map((r) => [r[1], r[2], r[8], r[9]]), [
    ['1/2', '6구 박스', '배즙 x6', 6],
    ['2/2', '싱글팩', '배즙 x1', 1]
  ]);
});

test('송장 매칭: 발주번호·수취인이 맞는 회신만 채택, 나머지는 보류/미회신', () => {
  const { built } = loadAll();
  const replies = C.parseReplyRows(REPLY_ROWS).records;
  const m = C.matchTracking(built.shipments, replies);

  assert.deepEqual(m.held.map((h) => h.reason).sort(), ['INVALID_TRACKING', 'RECIPIENT_MISMATCH', 'UNKNOWN_REF']);
  assert.equal(m.matched.length, 13);
  assert.deepEqual(m.unanswered.map((u) => u.line.recipient).sort(), ['강더미', '김가상', '김가상', '이샘플']);
  assert.equal(m.unansweredBoxes.length, 3);

  // 7개 주문(2박스)은 첫 박스 송장으로 업로드
  const bae = m.matched.find((x) => x.line.recipient === '배가짜');
  assert.ok(bae.ref.endsWith('-1'));
});

test('마켓별 송장 업로드 파일은 양식이 다르고 해당 마켓 주문만 담긴다', () => {
  const { built } = loadAll();
  const m = C.matchTracking(built.shipments, C.parseReplyRows(REPLY_ROWS).records);
  const a = C.toMarketUploadRows(m.matched, 'A');
  const b = C.toMarketUploadRows(m.matched, 'B');
  assert.deepEqual(a[0], ['상품주문번호', '배송방법', '택배사', '송장번호']);
  assert.deepEqual(b[0], ['묶음배송번호', '주문번호', '옵션ID', '택배사', '운송장번호']);
  assert.equal(a.length - 1, 7);
  assert.equal(b.length - 1, 6);
  assert.ok(a.slice(1).every((r) => r[0].startsWith('PA-')));
  assert.ok(b.slice(1).every((r) => r[0].startsWith('BD-')));
});

test('송장 회신 판정: 충돌·중복 송장번호는 보류, 같은 회신 재업로드는 무시', () => {
  const { built } = loadAll();
  const [s1, s2] = built.shipments;
  const r1 = s1.boxes[0].ref;
  const r2 = s2.boxes[0].ref;
  const m = C.matchTracking(built.shipments, [
    { ref: r1, carrier: 'X', trackingNo: '111111111111', recipient: s1.recipient, row: 2 },
    { ref: r1, carrier: 'X', trackingNo: '222222222222', recipient: s1.recipient, row: 3 },
    { ref: r2, carrier: 'X', trackingNo: '111111111111', recipient: s2.recipient, row: 4 }
  ]);
  assert.deepEqual(m.held.map((h) => h.reason), ['CONFLICT', 'DUP_TRACKING']);
  assert.ok(!m.matched.some((x) => x.ref === r1)); // 충돌 난 발주번호는 채택하지 않음

  const replies = C.parseReplyRows(REPLY_ROWS).records;
  const once = C.mergeReplies([], replies);
  const twice = C.mergeReplies(once.records, replies);
  assert.equal(twice.added, 0);
  assert.equal(twice.records.length, once.records.length);
});

test('대시보드 집계', () => {
  const { lines, built } = loadAll();
  const m = C.matchTracking(built.shipments, C.parseReplyRows(REPLY_ROWS).records);
  const s = C.summarize(lines, built, m);
  assert.deepEqual(
    {
      orderCount: s.orderCount, lineCount: s.lineCount, shippableLines: s.shippableLines,
      blockedLines: s.blockedLines, shipmentCount: s.shipmentCount, mergedShipments: s.mergedShipments,
      boxCount: s.boxCount, boxByType: s.boxByType
    },
    {
      orderCount: 20, lineCount: 22, shippableLines: 17, blockedLines: 5, shipmentCount: 13,
      mergedShipments: 4, boxCount: 17, boxByType: { SINGLE: 5, DOUBLE: 4, SIX: 8 }
    }
  );
  assert.deepEqual(s.tracking, { matched: 13, held: 3, unanswered: 4, unansweredBoxes: 3 });
});

test('데모용 가상 회신은 모든 박스를 매칭시킨다', () => {
  const { built } = loadAll();
  const reply = C.parseReplyRows(C.simulateReplyRows(built.shipments)).records;
  const m = C.matchTracking(built.shipments, reply);
  assert.equal(m.held.length, 0);
  assert.equal(m.unanswered.length, 0);
  assert.equal(m.matched.length, 17);
});
