// 샘플 주문으로 발주서를 만든 뒤, 그 발주번호에 대한 "가상 물류사 송장 회신"을 만든다.
// 매칭 예외를 보여주기 위해 일부러 몇 가지 오류를 섞는다.
'use strict';
const fs = require('fs');
const path = require('path');
const C = require('../js/core.js');

const fx = path.join(__dirname, '..', 'tests', 'fixtures');
const a = JSON.parse(fs.readFileSync(path.join(fx, 'a-market.json'), 'utf8'));
const b = JSON.parse(fs.readFileSync(path.join(fx, 'b-market.json'), 'utf8'));

let state = C.createState();
state = C.ingest(state, C.parseOrderRows(a, { fileName: 'a-market-orders.xlsx' }).lines).state;
state = C.ingest(state, C.parseOrderRows(b, { fileName: 'b-market-orders.xlsx' }).lines).state;
const { shipments } = C.buildShipments(C.allLines(state));

const rows = C.simulateReplyRows(shipments, '가상택배');
const header = rows[0];
let body = rows.slice(1);

// 1) 회신 누락: 강더미 출고분은 물류사가 아직 회신하지 않음
body = body.filter((r) => r[3] !== '강더미');
// 2) 수취인 불일치: 물류사 시스템에 오타로 입력됨
body.forEach((r) => { if (r[3] === '이샘플') r[3] = '이샘풀'; });
// 3) 송장번호 형식 오류
body.forEach((r) => { if (r[3] === '김가상') r[2] = '12-34'; });
// 4) 발주서에 없는 발주번호
body.push(['SH-00000000-1', '가상택배', '700000000000', '없는사람']);
// 5) 같은 행이 두 번 들어온 경우(무해한 중복)
const dup = body.find((r) => r[3] === '배가짜');
if (dup) body.push(dup.slice());

fs.writeFileSync(path.join(fx, 'reply.json'), JSON.stringify([header, ...body], null, 1), 'utf8');
console.log('reply rows:', body.length);
