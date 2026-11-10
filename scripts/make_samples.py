"""가상 샘플 데이터 생성기.

모든 이름·주소·연락처는 가상의 값이다(010-0000-xxxx, '가상시 샘플구' 등).
실제 회사 양식을 복제하지 않은, 이 데모 전용 'A마켓' / 'B마켓' / '물류사 회신' 양식이다.

생성물
  samples/a-market-orders.xlsx         A마켓 주문 (상단 안내행 + 헤더)
  samples/b-market-orders.xlsx         B마켓 주문 (주소/상세주소 분리)
  samples/logistics-tracking-reply.xlsx 가상 물류사 송장 회신
  tests/fixtures/*.json                 테스트용 동일 데이터(행 배열)

실행: python scripts/make_samples.py [node 실행 파일 경로]
"""
import json
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"

A_HEADER = ["상품주문번호", "주문번호", "주문일시", "구매자명", "수취인명", "수취인연락처1",
            "통합배송지", "우편번호", "상품명", "옵션정보", "수량", "배송메세지"]
A_PRODUCT = "가상농장 착즙주스"

# (상품주문번호, 주문번호, 주문일시, 구매자, 수취인, 연락처, 주소, 우편번호, 옵션, 수량, 메모)
A_ROWS = [
    ("PA-0001", "A-1001", "2026-11-10 09:12", "홍길동", "홍길동", "010-0000-0001",
     "가상시 샘플구 예시로 12, 101동 1001호", "00001", "맛: 사과 / 구성: 1개", "1", "문 앞에 놓아주세요"),
    ("PA-0002", "A-1002", "2026-11-10 09:30", "김가상", "김가상", "010-0000-0002",
     "가상시 테스트구 더미길 3", "00002", "맛: 사과 / 구성: 1개", "1", ""),
    ("PA-0003", "A-1002", "2026-11-10 09:30", "김가상", "김가상", "010-0000-0002",
     "가상시 테스트구 더미길 3", "00002", "맛: 배 / 구성: 1개", "1", ""),
    ("PA-0004", "A-1003", "2026-11-10 10:05", "이샘플", "이샘플", "010-0000-0003",
     "가상시 예제구 견본로 45", "00003", "맛: 포도 / 구성: 2개 세트", "2", "경비실 보관"),
    ("PA-0005", "A-1004", "2026-11-10 10:40", "박예시", "박예시", "010-0000-0004",
     "가상시 모의구 시험대로 7, 2층", "00004", "맛: 사과 / 구성: 3개 세트", "3", ""),
    ("PA-0006", "A-1005", "2026-11-10 11:02", "최테스트", "최테스트", "010-0000-0005",
     "가상시 샘플구 연습로 88", "00005", "맛: 사과 / 구성: 1개", "1", ""),
    ("PA-0007", "A-1006", "2026-11-10 13:20", "최테스트", "최테스트", "010-0000-0005",
     "가상시 샘플구 연습로 88", "00005", "맛: 배 / 구성: 2개 세트", "1", "부재 시 연락"),
    ("PA-0008", "A-1007", "2026-11-10 13:45", "정가상", "정", "010-0000-0006",
     "가상시 견본구 임시로 21", "00006", "맛: 포도 / 구성: 1개", "1", ""),
    ("PA-0009", "A-1008", "2026-11-10 14:10", "윤", "윤", "010-0000-0007",
     "가상시 견본구 임시로 22", "00007", "맛: 사과 / 구성: 1개", "2", ""),
    ("PA-0010", "A-1009", "2026-11-10 14:30", "임더미", "임더미", "010-0000-0008",
     "", "", "맛: 배 / 구성: 1개", "1", ""),
    ("PA-0011", "A-1010", "2026-11-10 15:00", "장모의", "장모의", "010-0000-0009",
     "가상시 테스트구 가짜길 9", "00009", "맛: 망고 / 구성: 1개", "1", ""),
    ("PA-0012", "A-1011", "2026-11-10 15:20", "문예시", "문예시", "010-12",
     "가상시 예제구 샘플로 31", "00010", "맛: 포도 / 구성: 1개", "1", ""),
    ("PA-0013", "A-1012", "2026-11-10 16:00", "배가짜", "배가짜", "010-0000-0011",
     "가상시 모의구 연습대로 100", "00011", "맛: 배 / 구성: 1개", "7", "선물 포장 불필요"),
]

B_HEADER = ["묶음배송번호", "주문번호", "옵션ID", "주문일", "구매자", "등록상품명", "등록옵션명",
            "구매수(수량)", "수취인이름", "수취인전화번호", "우편번호", "수취인 주소", "수취인 상세주소", "배송메세지"]
B_PRODUCT = "[가상농장] 착즙주스 모음"

# (묶음배송번호, 주문번호, 옵션ID, 주문일, 구매자, 옵션, 수량, 수취인, 연락처, 우편번호, 주소, 상세주소, 메모)
B_ROWS = [
    ("BD-9001", "B-5001", "OPT-11", "2026-11-10", "강더미", "사과즙 x 1입", "2", "강더미",
     "01000000101", "00101", "가상시 더미구 견본로 5", "301호", ""),
    ("BD-9002", "B-5002", "OPT-23", "2026-11-10", "조허구", "배즙 x 3입", "1", "조허구",
     "01000000102", "00102", "가상시 허구구 예시대로 17", "", "직접 수령"),
    ("BD-9002", "B-5002", "OPT-31", "2026-11-10", "조허구", "포도즙 x 1입", "1", "조허구",
     "01000000102", "00102", "가상시 허구구 예시대로 17", "", "직접 수령"),
    ("BD-9003", "B-5003", "OPT-32", "2026-11-10", "한가짜", "포도즙 x 2입", "7", "한가짜",
     "01000000103", "00103", "가상시 가짜구 샘플길 2", "B동 5층", ""),
    ("BD-9004", "B-5004", "OPT-11", "2026-11-10", "서예제", "사과즙 x 1입", "1", "서예제",
     "01000000104", "00104", "가상시 예제구 모의로 7", "102동 304호", ""),
    ("BD-9005", "B-5005", "OPT-21", "2026-11-10", "서예제", "배즙 x 1입", "1", "서예제",
     "01000000104", "00104", "가상시  예제구 모의로 7,", "102동  304호", ""),
    ("BD-9006", "B-5006", "OPT-11", "2026-11-10", "오목업", "사과즙 x 1입", "0", "오목업",
     "01000000105", "00105", "가상시 목업구 테스트로 11", "", ""),
    ("BD-9007", "B-5007", "OPT-21", "2026-11-10", "신모의", "배즙 x 1입", "1", "신모의",
     "01000000106", "00106", "  ", "", ""),
    ("BD-9008", "B-5008", "OPT-11", "2026-11-10", "홍길동", "사과즙 x 1입", "1", "홍길동",
     "010-0000-0001", "00001", "가상시 샘플구 예시로 12", "101동 1001호", "문 앞에 놓아주세요"),
]


def a_aoa():
    rows = [["A마켓 주문내역 (가상 양식 · 데모용)"] + [""] * (len(A_HEADER) - 1), A_HEADER]
    for r in A_ROWS:
        line_no, order_no, at, buyer, rcpt, phone, addr, zip_, opt, qty, msg = r
        rows.append([line_no, order_no, at, buyer, rcpt, phone, addr, zip_, A_PRODUCT, opt, qty, msg])
    return rows


def b_aoa():
    rows = [B_HEADER]
    for r in B_ROWS:
        bundle, order_no, opt_id, at, buyer, opt, qty, rcpt, phone, zip_, a1, a2, msg = r
        rows.append([bundle, order_no, opt_id, at, buyer, B_PRODUCT, opt, qty, rcpt, phone, zip_, a1, a2, msg])
    return rows


def write_xlsx(path, rows, header_index, sheet_title):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    for r in rows:
        ws.append([str(c) for c in r])  # 전화번호·우편번호 앞자리 0 보존을 위해 모두 문자열
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="E8EEF4")
    for cell in ws[header_index + 1]:
        cell.font = bold
        cell.fill = fill
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, width * 1.6), 45)
    wb.save(path)


def main():
    node = sys.argv[1] if len(sys.argv) > 1 else "node"
    SAMPLES.mkdir(exist_ok=True)
    FIXTURES.mkdir(parents=True, exist_ok=True)

    a, b = a_aoa(), b_aoa()
    (FIXTURES / "a-market.json").write_text(json.dumps(a, ensure_ascii=False, indent=1), encoding="utf-8")
    (FIXTURES / "b-market.json").write_text(json.dumps(b, ensure_ascii=False, indent=1), encoding="utf-8")
    write_xlsx(SAMPLES / "a-market-orders.xlsx", a, 1, "주문내역")
    write_xlsx(SAMPLES / "b-market-orders.xlsx", b, 0, "Sheet1")

    # 송장 회신은 발주번호(출고번호 해시)에 의존하므로 핵심 로직으로 계산한다.
    subprocess.run([node, str(ROOT / "scripts" / "make-reply.js")], check=True)
    reply = json.loads((FIXTURES / "reply.json").read_text(encoding="utf-8"))
    write_xlsx(SAMPLES / "logistics-tracking-reply.xlsx", reply, 0, "송장회신")
    print("samples written:", ", ".join(p.name for p in sorted(SAMPLES.glob("*.xlsx"))))


if __name__ == "__main__":
    main()
