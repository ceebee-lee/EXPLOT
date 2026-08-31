# PO·FOC LOT 번호 자동 배정

PO/FOC 주문 파일과 글로벌 재고 파일을 업로드하여 LOT 배정 결과 엑셀을 생성하는 Streamlit 앱입니다.

## 기본 배정 규칙

- PO 배정 후 남은 재고로 FOC 배정
- GD 단독 → DU 단독 → 필요 시 GD+DU 합산
- GD+DU 합산 재고도 부족하면 해당 주문 전체 미배정
- 판정이 `통과`인 GD/DU 재고만 사용하고 GL은 제외
- 빠른 EXP → 잔여수량이 적은 LOT → LOT 번호순
- EXP 미확인 재고는 제외하지 않고 정상 EXP 뒤에 배정

웹 화면의 **고급 배정 설정**에서 창고 우선순위, PO/FOC 우선순위, 창고 합산 허용 여부를 변경할 수 있습니다.

## 로컬 실행

Python 3.11 또는 3.12를 권장합니다.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Streamlit Community Cloud 배포

1. 이 폴더의 파일을 GitHub **비공개 저장소**에 올립니다.
2. Streamlit Community Cloud에서 저장소를 선택합니다.
3. Main file path를 `app.py`로 지정합니다.
4. Python 3.11 또는 3.12를 선택하고 배포합니다.
5. Sharing 설정을 `Only specific people can view this app`으로 설정합니다.

> 회사 주문·재고 정보를 처리하므로 공개 앱으로 배포하지 마세요. 회사 보안 정책상 외부 클라우드 업로드가 불가하면 사내 서버에 배포해야 합니다.

## 입력 파일

### PO/FOC 주문 파일

- `PO` 또는 `FOC` 시트
- 1행 헤더: `Barcode`, `Product`, `Quantity`

### 글로벌 재고 파일

- `글로벌` 또는 `글로벌 유통기한` 시트
- 1행 헤더: `저장 위치`, `자재`, `자재내역`, `가용재고(ASN)`, `유통기한`, `로트번호(생산)`, `판정`

## 결과 엑셀

- `LOT 배정`
- `주문 요약`
- `검증 요약`

## 명령줄 기능

기존 PC 실행 방식도 유지됩니다.

```powershell
python ".\PO_FOC_LOT_배정_자동화.py" --order ".\PO_qty.xlsx" --inventory ".\글로벌재고.xlsx" --out ".\배정결과.xlsx"
```
