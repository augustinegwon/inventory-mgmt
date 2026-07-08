# 스프레드시트 구조 (현재 배포본 기준)

`260707_Inventory.xlsx`에서 추출한 **실제 시트 상태**입니다.
스크립트는 `Setup.gs` / `Transaction.gs` 두 파일로 구성됩니다(아래 참고).

## 시트 목록 / 데이터 흐름
```
Ledger ──(연산)──▶ Inventory ──(참조)──▶ Dashboard
```
| 시트 | 역할 |
|---|---|
| `Transaction` | 입력 폼 UI |
| `Ledger` | 거래 원장 (append-only, 재고의 단일 진실 공급원) |
| `Inventory` | Ledger를 연산해 정리한 **정규화 재고 목록** (Category/Item/Location/Serial/Qty) |
| `Dashboard` | Inventory를 **피벗**한 아이템 × 위치 매트릭스 (표시 계층) |
| `Opening` | 기초재고 일괄입력 스테이징 (→ Ledger OPENING 거래로 등록) |
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

## Transaction — 실제 셀/수식
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

## Inventory — 정규화된 재고 목록 (스크립트가 Ledger 연산 후 기록)
헤더: `Category | Item | Location | Serial Number | Quantity` (A1:E1), 데이터는 2행부터.
- 한 행 = 하나의 재고 항목, **수량 > 0** 인 것만
  - 시리얼 관리 품목: **시리얼별**로 한 행 (수량 보통 1)
  - 비시리얼 품목: **(품목, 위치)별**로 한 행, Serial 칸 공백
- 정렬: Category → Item → Location → Serial
- **라이브 수식이 아니라 스크립트 값**: `rebuildInv_()`가 계산해 기록.
  - 거래 제출(`submitTransaction`) 때마다 자동 재계산
  - 수동: 메뉴 📦 Inventory → **재고 새로고침 (Rebuild Inventory)**
- Freeze: 1행

## Opening — 기초재고 일괄입력
파일 사용 전, 초기 실사 재고를 한 번에 넣기 위한 스테이징 시트.
- 입력 열(3행부터): `Item`(드롭다운) / `Location`(드롭다운) / `Serial Number` / `Quantity`
- 메뉴 📦 Inventory → **기초재고 일괄등록 (Import Opening)** 실행 시:
  - 각 행 검증(품목 존재·수량 양의 정수·시리얼 품목은 시리얼 필수·중복 차단)
  - 통과 시 Ledger에 `Type=ADD, From=EXTERNAL (OPENING), To=위치, Worker=OPENING, Note=Opening balance` 로 기록
  - 하나라도 오류면 **전체 미기록**(원자적), 성공 시 스테이징 표는 비워짐 + Inventory 재계산
- 초기재고는 이렇게 **Ledger 단일 원천**으로 관리되며, `From = EXTERNAL (OPENING)` 으로 필터하면 초기 실사 목록만 조회 가능

## Dashboard — 표시 계층 (Inventory 피벗)
- `A2` `Item` / `B2` `Total` / `C2~` 버킷 헤더
- `A3` 아이템 목록: `=FILTER(LIST_ITEM, LIST_ITEM<>"")`
- `B3` 총 수량: `BYROW(items, LAMBDA(i, SUMIFS(Inventory!E, Inventory!B, i)))`
- `C3` 매트릭스: `MAKEARRAY(..., LAMBDA(r,c, SUMIFS(Inventory!E, Inventory!B, item, Inventory!C, bucket)))`
- 계산 원천은 **Ledger 가 아니라 Inventory** (Ledger ▶ Inventory ▶ Dashboard)
- Freeze: 2행 + 2열(A Item, B Total)

## ✅ 스크립트 파일 구성
- `Setup.gs` — 공용 상수 + `onOpen`(메뉴), `initializeSystem`(빈 시트 부트스트랩), `setupInputSheet`, `rebuildInv`/`rebuildInv_`(Inventory 재계산), `createTriggers`, `migrateSerialsToText`, 시트 헬퍼(`getRequiredSheet_`/`getOrCreateSheet_`/`setupLedgerSheet_`/`setupSettingsSheet_`/`defineNamedRanges_`/`setupDashboardSheet_`/`addInvQty_`)
- `Transaction.gs` — `updateDynamicUI`, `submitTransaction` + 재고 계산 헬퍼

## 🚀 빈 스프레드시트에서 처음 세팅하는 순서
1. Apps Script 편집기에 `Setup.gs` / `Transaction.gs` 붙여넣고 저장
2. 시트 새로고침 → **📦 Inventory → 🚀 전체 초기화 (Initialize)** 실행 (권한 승인)
   - Ledger/Settings/Dashboard/Transaction 시트 + 헤더 + 마스터 데이터 + 이름범위 + 대시보드 수식 생성, 빈 `Sheet1` 제거
3. **📦 Inventory → 편집 트리거 등록** 실행 (동적 UI 동작에 필요)
4. 제출 버튼(도형)에 `submitTransaction` 함수 연결 (또는 메뉴의 "트랜잭션 제출" 사용)

> (해결됨) 과거 `setupInputSheet()`는 구버전 헬퍼 수식(단순 FILTER)을 써서 재실행 시 검색 UI·투어 버킷이 깨질 위험이 있었으나,
> 현재 `Setup.gs`의 `setupInputSheet()`는 위에 정리된 **현재 배포 수식(LET/LAMBDA, LIST_BUCKET)**을 그대로 재현하도록 갱신되었습니다.
