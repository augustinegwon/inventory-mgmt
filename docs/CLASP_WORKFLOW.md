# clasp 작업 워크플로우 (라이브 시트 기준)

이 저장소는 **라이브 Google 스프레드시트에 배포된 Apps Script 코드**를 소스 오브 트루스로 관리한다.

- 대상 스프레드시트: `1PMEgTQNoXFRjTimq6BCsX_w4R9PTvwedpvhlbInVXBM`
- Apps Script Script ID: `1gqZZlemsn6194skACvBECQTxpUFEdXYPUKuRuNhKodvmesX2omFsjGDu`
- 소유/관리 계정: **auggie@amazevr.com**

## 파일 구성 (루트 = clasp 프로젝트)
| 파일 | 역할 |
|---|---|
| `.clasp.json` | clasp 프로젝트 설정 (scriptId 포함) |
| `appsscript.json` | Apps Script 매니페스트 (시간대 Asia/Seoul, V8) |
| `Setup.js` | 초기화·시트 생성·Inventory/Dashboard 빌드·입력폼 |
| `Transaction.js` | 동적 UI + 트랜잭션 제출 로직 |
| `Migration.js` | (1회성) ID 체계 마이그레이션 |
| `Origin Migration.js` | (1회성) 기초재고 → Ledger 이관 |
| `archive/redesign-v1/` | 이전 세션의 재설계본(.gs). 참고용 보존, 배포 안 함 |

> clasp는 서버의 `.gs` 파일을 로컬에 `.js`로 내려받는다. (`.clasp.json`의 `scriptExtensions`에 `.js`, `.gs` 모두 등록)

## 인증 (컨테이너는 세션마다 초기화됨 — 매 세션 재로그인 필요)
```bash
clasp login --no-localhost   # auggie@amazevr.com 으로 승인, 리디렉트 URL 붙여넣기
```
자격증명은 `~/.clasprc.json` 에 저장된다(저장소에 커밋 금지 — `.gitignore` 처리됨).

## 코드 내려받기 / 배포
```bash
clasp pull    # 라이브 → 로컬 (배포본과 동기화 확인)
clasp push    # 로컬 → 라이브 (변경 배포)
```

## 주의
- Sheets API(raw)는 clasp 기본 GCP 프로젝트에서 비활성화되어 있어 `values.get` 등이 403.
  시트 **데이터 읽기**는 Google Drive 커넥터 또는 Apps Script 함수 실행으로 처리한다.
- 배포 전 반드시 `clasp pull` 로 라이브가 로컬보다 앞서 있지 않은지 확인(다른 사람이 편집기에서 직접 수정했을 수 있음).
