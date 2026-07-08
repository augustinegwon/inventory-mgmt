# 스프레드시트 구조 (현재 배포본 기준)

`260707_Inventory.xlsx`에서 추출한 **실제 시트 상태**입니다.
스크립트는 `Setup.gs` / `Transaction.gs` 두 파일로 구성됩니다(아래 참고).

## 시트 목록
| 시트 | 역할 |
|---|---|
| `Input_Transaction` | 입력 폼 UI |
| `Ledger` | 거래 원장 (append-only, 재고의 단일 진실 공급원) |
| `Dashboard` | 아이템 × 위치 재고 매트릭스 |
| `Settings` | 마스터 데이터 |

## 이름 있는 범위 (Named Ranges)
| 이름 | 정의 | 상태 |
|---|---|---|
| `LIST_CATEGORY` | `Settings!$A$2:$A$1000` | OK |
| `LIST_ITEM` | `Settings!$B$2:$B$1000` | OK |
| `LIST_BUCKET` | `Settings!$E$2:$F$1000` | OK (물리 위치 E + 투어 패키지 F 를 하나의 "버킷"으로 통합) |
| `LIST_LOCATION_PHYSICAL` | `#REF!` | ⚠️ 깨짐 (미사용) |
| `LIST_LOCATION_TOUR` | `#REF!` | ⚠️ 깨짐 (미사용) |

## Ledger (원장) — A~J
`Timestamp | Type | Category | Item | Serial Number | From | To | Quantity | Worker | Note`

- 재고 = `SUMIFS(수량, To=위치) − SUMIFS(수량, From=위치)` (아이템·위치별)
- `ADD` → From=`EXTERNAL (VENDOR)`, `REMOVE` → To=`EXTERNAL (SCRAP)`

## Settings (마스터)
- A: Category / B: Item / C: Manage Serial(YES·NO)
- E: Location(물리 창고) / F: Tour Package(투어)
- H: USER(이메일)

## Input_Transaction — 실제 셀/수식
검색 영역(B~C)과 트랜잭션 폼(E~F), 그리고 숨김 헬퍼 열(W/X/Y/Z).

| 셀 | 내용 |
|---|---|
| `C3` | 검색 Category (드롭다운: LIST_CATEGORY) |
| `C4` | 검색 Item (드롭다운: Z열 필터 결과) |
| `B7:C8` | 선택 아이템의 위치별 현재고 요약 (LET/BYROW 수식) |
| `F3` | Type (ADD/MOVE/REMOVE) |
| `F4` | Item = `=$C$4` (검색 Item 미러링) |
| `F5` | Serial No. |
| `F6` | FROM |
| `F7` | TO |
| `F8` | Quantity |
| `F9` | USER |
| `F10` | Note |

### 헬퍼 열 (현재 배포본 — LET/LAMBDA/TOCOL 사용)
- `W1` FROM 후보(재고>0 버킷):
  `=IFERROR(UNIQUE(FILTER(TOCOL(LIST_BUCKET,1,TRUE), BYROW(TOCOL(LIST_BUCKET,1,TRUE), LAMBDA(b, SUMIFS(Ledger!$H:$H,Ledger!$D:$D,$C$4,Ledger!$G:$G,b)-SUMIFS(Ledger!$H:$H,Ledger!$D:$D,$C$4,Ledger!$F:$F,b)))>0)),"")`
- `X1` 시리얼 후보(FROM 위치에 재고 있는 시리얼):
  `=IFERROR(LET(i,$C$4, loc,$F$6, sers,UNIQUE(FILTER(Ledger!E:E,(Ledger!D:D=i)*(Ledger!E:E<>"")*(Ledger!E:E<>"N/A"))), bals,BYROW(sers,LAMBDA(s, SUMIFS(Ledger!H:H,Ledger!D:D,i,Ledger!E:E,s,Ledger!G:G,loc)-SUMIFS(Ledger!H:H,Ledger!D:D,i,Ledger!E:E,s,Ledger!F:F,loc))), FILTER(sers,bals>0)),"")`
- `Y1` TO 후보(전체 버킷): `=IFERROR(TOCOL(LIST_BUCKET,1,TRUE),"")`
- `Z1` 검색 Item 후보(Category로 필터):
  `=IFERROR(IF(C3="",FILTER(LIST_ITEM,LIST_ITEM<>""),FILTER(LIST_ITEM,LIST_CATEGORY=C3,LIST_ITEM<>"")),"")`

## Dashboard — 재고 매트릭스
- `A2` `Item` / `B2` `Total` / `C2~` 버킷 헤더
- `A3` 아래로 아이템 목록: `=FILTER(LIST_ITEM, LIST_ITEM<>"")`
- `B3` 품목별 **총 수량**(spill): `BYROW(items, LAMBDA(i, SUMIFS(H,D=i,B="ADD") - SUMIFS(H,D=i,B="REMOVE")))` — MOVE는 총량 불변이라 ADD−REMOVE = 매트릭스 행 합계
- `C2` 가로 헤더(버킷): `=IFERROR(TRANSPOSE(TOCOL(LIST_BUCKET,1,TRUE)),"")`
- `C3` 매트릭스 본문(spill): `MAKEARRAY(아이템수, 버킷수, LAMBDA(r,c, SUMIFS(To)-SUMIFS(From)))`
- Freeze: 2행 + **2열(A Item, B Total)**

## ✅ 스크립트 파일 구성
- `Setup.gs` — 공용 상수 + `onOpen`(메뉴), `initializeSystem`(빈 시트 부트스트랩), `setupInputSheet`, `createTriggers`, `migrateSerialsToText`, 시트 헬퍼(`getRequiredSheet_`/`getOrCreateSheet_`/`setupLedgerSheet_`/`setupSettingsSheet_`/`defineNamedRanges_`/`setupDashboardSheet_`)
- `Transaction.gs` — `updateDynamicUI`, `submitTransaction` + 재고 계산 헬퍼

## 🚀 빈 스프레드시트에서 처음 세팅하는 순서
1. Apps Script 편집기에 `Setup.gs` / `Transaction.gs` 붙여넣고 저장
2. 시트 새로고침 → **📦 Inventory → 🚀 전체 초기화 (Initialize)** 실행 (권한 승인)
   - Ledger/Settings/Dashboard/Input_Transaction 시트 + 헤더 + 마스터 데이터 + 이름범위 + 대시보드 수식 생성, 빈 `Sheet1` 제거
3. **📦 Inventory → 편집 트리거 등록** 실행 (동적 UI 동작에 필요)
4. 제출 버튼(도형)에 `submitTransaction` 함수 연결 (또는 메뉴의 "트랜잭션 제출" 사용)

> (해결됨) 과거 `setupInputSheet()`는 구버전 헬퍼 수식(단순 FILTER)을 써서 재실행 시 검색 UI·투어 버킷이 깨질 위험이 있었으나,
> 현재 `Setup.gs`의 `setupInputSheet()`는 위에 정리된 **현재 배포 수식(LET/LAMBDA, LIST_BUCKET)**을 그대로 재현하도록 갱신되었습니다.
