# AGENTS.md — AI 에이전트 셋업 가이드

이 저장소를 받은 **AI 에이전트가 사람 대신 셋업을 자동화**하기 위한 지침입니다.
원칙: **사람이 꼭 해야 하는 것(아래 🔴)만 사람에게 요청하고, 나머지(🟢)는 에이전트가 직접 처리**한다.

---

## 0. 시스템 한 줄 이해
- Google Sheets(컨테이너) + Apps Script. 코드는 `gas/`에 있고, `Setup.gs`가 시트·양식·예시데이터·검증·수식을 **전부 자동 생성**한다.
- 출력 양식(파란 거래명세서)은 `docs/거래명세서양식.xlsx` 를 시트로 가져와야 존재한다. 코드가 그 양식의 입력칸을 채워 PDF로 내보낸다.

## 1. 역할 분담 (반드시 지킬 것)

| | 작업 | 누가 |
|---|------|------|
| 🔴 | Google 계정 로그인 / `clasp login` (OAuth) | **사람** (보안상 위임 불가) |
| 🔴 | Apps Script 권한 승인 (첫 실행 OAuth 동의) | **사람** (브라우저 클릭) |
| 🔴 | `docs/거래명세서양식.xlsx` 를 시트로 가져오기 | **사람** (Sheets UI, 30초) |
| 🔴 | 실제 회사/거래처 데이터 입력, `발송_활성화=TRUE` | **사람** (본인 데이터) |
| 🟢 | Apps Script 프로젝트 생성·코드 push (`clasp`) | **에이전트** |
| 🟢 | 시트 생성·양식 통합·예시데이터·드롭다운·검증·수식 | **에이전트가 트리거**(Setup.gs가 수행) |
| 🟢 | 권한 매니페스트(`appsscript.json`) 적용 | **에이전트** |
| 🟢 | 사람에게 남은 4가지를 **명확한 체크리스트로 안내** | **에이전트** |

## 2. 자동화 경로 (clasp)

> clasp = Apps Script CLI. 이걸로 코드 배포를 자동화한다.

### 사전 (🔴 사람에게 1회 요청)
```bash
npm install -g @google/clasp     # (에이전트가 실행 가능)
clasp login                      # 🔴 사람: 브라우저에서 Google 로그인·동의
```
또한 https://script.google.com/home/usersettings 에서 **Apps Script API: 사용** 토글 (🔴 사람, 1회).

### 에이전트 실행 (🟢)
```bash
# 1) 시트+바인딩 스크립트 새로 생성
clasp create --type sheets --title "거래명세서_자동화"
#   → .clasp.json(scriptId) 생성, 새 구글시트 URL 출력 → 사람에게 전달

# 2) 코드 배포
cp gas/Code.gs gas/Setup.gs gas/appsscript.json .   # clasp 루트로 복사(또는 rootDir 설정)
clasp push -f
```

### 배포 후 (🔴 사람 — 에이전트가 안내)
1. 출력된 시트 URL 열기
2. `파일 → 가져오기` → `docs/거래명세서양식.xlsx` → **"새 시트 삽입"** (양식 2탭 생성)
3. 새로고침 → 메뉴 `📨 거래명세서 → ⚙ 관리 → 설치 / 갱신` → 권한 승인
   - (이 한 번의 클릭이 시트·양식통합·예시데이터·검증·수식을 모두 자동 생성)
4. `설정`·`거래처`·`정기항목`에 실제 데이터 입력
5. 테스트(`👀 미리보기` / `테스트 발송`) → 이상 없으면 `발송_활성화=TRUE`

> `clasp run setupSheets` 로 3번까지 자동화하려면 API 실행 배포·스코프 설정이 추가로 필요(고급). 보통은 사람이 메뉴 1번 클릭이 가장 간단·안전하다.

## 3. 붙여넣기 경로 (clasp 없이, 더 간단)
clasp/Node 없이도 됨. 에이전트는 사람에게 아래를 안내한다(코드 내용은 에이전트가 클립보드/메시지로 제공):
1. 새 구글시트 → 양식 xlsx 가져오기
2. `확장 → Apps Script` → `appsscript.json`·`Code.gs`·`Setup.gs` 붙여넣기
3. 메뉴 `⚙ 관리 → 설치 / 갱신` → 권한 승인
4. 데이터 입력 → `발송_활성화=TRUE`

## 4. 코드 구조 (에이전트 참고)
- `gas/Setup.gs`
  - `setupSheets()` 전체 설치(데이터 초기화 — 신규용) / `setupUpgrades()` 데이터 보존 갱신
  - 시트: `설정·거래처·정기항목·거래내역·세금계산서양식·발송로그`
  - `unifyToOneSheet()` 자료입력→세금계산서양식 한 시트 통합(수식 재연결)
  - 입력 형식 검증(드롭다운), 금액 자동수식, 거래처ID 드롭다운
- `gas/Code.gs`
  - 진입점: `previewMonthly()` / `runMonthlyConfirm()` / `runMonthly()`(트리거) / `runPOC()`
  - 정기항목→거래내역 `generateMonthlyCore_()`, 발송 `processCustomer()`, PDF export, Drive 보관, Gmail
  - 자동발송 트리거: `createMonthlyTrigger()` / `deleteTriggers()` / 상태 `showAutoSendStatus()`
- `gas/appsscript.json` — 필요한 OAuth 스코프(spreadsheets, drive, gmail.send, script.external_request, script.scriptapp, userinfo.email)

## 5. 안전 규칙
- 처음엔 반드시 `발송_활성화=FALSE` 로 두고 `테스트_수신메일`로만 검증한다. 실고객 발송은 사람 확인 후 `TRUE`.
- 코드/문서의 모든 값은 예시다. 실데이터는 사람이 시트에 입력한다(저장소에 커밋 금지 — `.gitignore`로 `capture/`, `*.pdf` 등 제외).

## 6. 검증
- 순수 로직은 GAS 없이 Node로 회귀 테스트 가능(개발 시). 함수 의존성은 `Setup.gs`/`Code.gs` 상호 호출.
- 셋업 성공 기준: `📨 거래명세서` 메뉴 표시 + `⚙ 관리 → 설치/갱신` 후 시트 6개 생성 + `👀 미리보기`로 PDF 1건 생성.
