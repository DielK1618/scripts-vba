# HANDOFF.claude.md — corelib 인수인계 (Claude)

> 대표 문서 `HANDOFF.md`(2026-10-16까지 보관)와 함께 읽는다. 규칙: `Git/CLAUDE.md` 3절.

## 2026-10-16 13:45 KST · Claude · am_Inspect 모듈 신설 · 진행 중

**현재 상태**
- ✅ 소스 작성: `xlam/am_Inspect.bas`, `cwb/ref_Inspect.bas` (UTF-8, BOM 없음)
- ✅ 문서 반영: `CLAUDE.md`, `INDEX.md`, `LOG.claude.md`, `../TROUBLESHOOTING.claude.md`
- ✅ 이 PC(Excel 16 64bit) 임시 통합문서에서 컴파일·실행 검증 — 세부는 `LOG.claude.md`
- ⬜ 커밋 안 함 (마무리 절차 때 이 세션에서 바꾼 파일만 커밋)

**다음 할 일 (사용자 PC에서 실행)**
1. ⬜ `corelib.xlam` VBE 에서 `am_Inspect.bas` 가져오기 (UIAutomationClient 참조는 이미 체크됨 — 기존 HANDOFF.md 2026-07-24 항목)
2. ⬜ CWB 에서 `ref_Inspect.bas` 가져오기 (CWB 에는 참조 추가 불필요)
3. ⬜ xlam 에서 `디버그 > VBAProject 컴파일` — am_Automation 의 미완료 컴파일 체크(HANDOFF.md)도 함께 확인됨
4. ⬜ CWB 직접 실행 창에서 동작 확인
   ```vba
   ?DumpWindows                  ' 창 목록 시트 생성, 창 수 반환
   ?DumpElementTree("계산기")    ' 계산기 요소 트리 — 숫자 버튼이 Button/Invoke 로 나오면 정상
   ?DumpElementAtCursor(5)       ' 5초 안에 마우스를 계산기 버튼 위에 → 마지막 행이 그 버튼
   ```
   ※ `?` 로 실행할 때 인수가 있으면 괄호 필수 — 괄호 없이 쓰면 "선택적 인수가 아닙니다" 컴파일 오류 (2026-10-16 회사 PC 실측)
5. ⬜ 회사 PC 사내 메신저로 같은 순서 실행 → 결과 시트로 제어 가능 여부 판단 → 필요 시 am_Automation 에 제어 함수(예: UIA 값 입력) 추가

**확인 필요 지점**
- 회사 PC = Windows 64bit + **Excel 32bit** 확인(2026-10-16) → 커서 조회는 32bit 경로(`WindowFromPoint(x, y)`)를 탐. 이 PC(64bit Excel)에서는 실행되지 않은 경로라 회사 PC 에서 `DumpElementAtCursor` 로 확인 필요
- 커서 조회는 화면 위치로 안쪽 요소를 찾아 내려가는 방식 — 겹친 요소는 면적이 가장 작은 쪽을 고름. 계산기 버튼 위에서 마지막 행이 그 버튼인지 확인
- Chromium/Electron 계열 앱은 접근성이 꺼져 있으면 창 틀만 보임 — 코드 문제가 아님

## 2026-10-16 15:10 KST · Claude · WMC Messenger 트리 항목 선택 · 진행 중

**현재 상태**
- ✅ `am_Inspect`(커서 조회 수정), `am_Automation`(`SelectItemByUIA`/`ExpandItemByUIA` 추가), `ref_Automation`(래퍼) 소스 수정
- ✅ 이 PC 에서 WMC 와 같은 구조의 테스트 프로그램으로 검증 완료 (`LOG.claude.md`)
- ⬜ 커밋 안 함

**다음 할 일 (회사 PC)**
1. ⬜ xlam: `am_Inspect`, `am_Automation` 제거 후 다시 가져오기 → `디버그 > VBAProject 컴파일`
2. ⬜ CWB: `ref_Automation` 제거 후 다시 가져오기
3. ⬜ CWB 직접 실행 창
   ```vba
   ?SelectItemByUIA("WMC Messenger", "Target", True, "TreeItem")   ' True + 메신저에서 Target 선택·펼침
   ?DumpElementAtCursor(5)                                          ' 그룹 안 사용자 위 → 마지막 행이 그 사용자
   ```
4. ⬜ 선택 다음에 하려는 동작(쪽지·대화창 열기 등) 확인 → 필요한 제어 함수 추가. 사용자 항목 더블클릭이 필요하면 UIA 에 더블클릭이 없으므로 항목 화면 위치 기준 좌표 더블클릭 함수 검토

## 2026-10-16 15:50 KST · Claude · WMC Messenger 우클릭 메뉴 실행 · 진행 중

**현재 상태**
- ✅ `SelectItemByUIA` 회사 PC 동작 확인 (사용자)
- ✅ `ClickContextMenuByUIA` 추가·이 PC 검증 (`LOG.claude.md`)
- ⬜ 커밋 안 함

**다음 할 일 (회사 PC)**
1. ⬜ xlam: `am_Automation` 제거 후 다시 가져오기 → 컴파일 / CWB: `ref_Automation` 제거 후 다시 가져오기
2. ⬜ 직접 실행 창
   ```vba
   ?ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem", False)   ' 마우스 없이 되는지 먼저
   ?ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem")          ' False 였다면 마우스 대체 허용
   ```
   False 면 직접 실행 창의 `am_Automation.ClickContextMenuByUIA: …` 줄로 실패 단계 확인
3. ⬜ 친구추가 뒤에 뜨는 창(대상 선택·확인 버튼 등)이 있으면 `DumpWindows` → `DumpElementTree` 로 구조 확인 후 다음 단계 자동화

## 2026-10-16 16:30 KST · Claude · WMC 자동화 — 실행부터 친구추가까지 · 진행 중

**현재 상태**
- ✅ 회사 PC: `ClickContextMenuByUIA` 마우스 대체 경로로 친구추가 실행 확인 (WMC 는 마우스 우클릭 방식)
- ✅ `OpenAppWindow` 추가·이 PC 검증, `DumpWindows` 실행 파일 경로 열 추가
- ⬜ 커밋 안 함

**다음 할 일 (회사 PC)**
1. ⬜ xlam: `am_Automation`, `am_Inspect` 다시 가져오기 → 컴파일 / CWB: `ref_Automation` 다시 가져오기
2. ⬜ `?DumpWindows` → WMC 행의 '실행 파일 경로' 확인
3. ⬜ 전체 흐름 확인
   ```vba
   ?OpenAppWindow("WMC Messenger", "<실행 파일 경로>", "Target")
   ?ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem")
   ```
4. ⬜ WMC 를 완전히 종료한 상태에서 1번 줄 실행 — 로그인 창이 먼저 뜨는지 확인 (뜨면 로그인 단계 자동화 필요)
5. ⬜ 친구추가 뒤에 뜨는 창 구조 확인 → 다음 단계 자동화

## 2026-10-16 16:00 KST · Claude · WMC 로그인 자동화 · 진행 중 (사용자 결정 대기)

**현재 상태**
- ✅ 범용 함수 `IsWindowOpen` / `WaitElementByUIA` / `SetValueByUIA` / `InvokeByUIA`, `#아이디`·`=이름`·`A|B` 규칙 추가, 이 PC 검증
- ⬜ 사용자 결정: 비밀번호 보관 방식(자격 증명 관리자 권장 / 매번 입력), `txLifeDispNo` 칸 용도·필수 여부
- ⬜ 결정 후 CWB `tpl_Messenger.bas`(가칭)에 WMC 전용 흐름 작성: 실행 → 로그인 → 목록 대기 → 친구추가
- ⬜ 회사 PC 검증: WMC 실제 입력칸(외부 라이브러리 편집기, `[Editor] Edit Area` 자식 있음)에 값 입력이 되는지 — 안 되면 키보드 대체 경로가 동작
- ⬜ 커밋 안 함

## 2026-10-16 16:20 KST · Claude · WMC 실행~로그인~그룹 메뉴 · 진행 중 (회사 PC 검증 대기)

**현재 상태**
- ✅ `am_Credential`/`ref_Credential`, `tpl_Messenger` 작성, 이 PC 에서 실제 구조(진짜 corelib.xlam + CWB)로 전체 흐름 검증
- ⬜ 세 번째 칸(`txLifeDispNo`) 값이 고정값인지 매번 바뀌는지 미확인 → 지금은 비워 두어 매번 입력
- ⬜ 커밋 안 함

**다음 할 일 (회사 PC)**
1. ⬜ xlam: `am_Automation`, `am_Inspect` 다시 가져오기 + `am_Credential` 새로 가져오기 → 컴파일
2. ⬜ CWB: `ref_Automation` 다시 가져오기 + `ref_Credential`, `tpl_Messenger` 새로 가져오기 → 컴파일
3. ⬜ WMC 를 완전히 종료한 뒤 `?WmcAddFriend("Target")` — 입력 창에 아이디·비밀번호, 세 번째 칸 입력 → True 확인
4. ⬜ WMC 실제 입력칸(Infragistics 계열로 보임, `[Editor] Edit Area` 자식)에 값 입력이 되는지 — 직접 실행 창에 "키보드 입력으로 대체" 가 찍히면 1차 방식이 안 된 것
5. ⬜ 친구추가 뒤에 뜨는 창 구조 확인 → 다음 단계

## 2026-10-16 16:35 KST · Claude · 모듈 버전 줄 · 완료

- 회사 PC 에 반영할 때 각 모듈 3번째 줄 `' 버전 : …` 을 `INDEX.md` "모듈 버전" 표와 비교
- 이번 세션 변경 7개 모듈은 모두 `2026-10-16 16:32` — 회사 PC 에 이 시각으로 다시 붙여 넣어야 최신 (`—` → `-` 교체 포함)
- ⬜ 선택: "한 번에 확인" 매크로(VBE 접근 신뢰 필요) — 사용자 판단 대기
- ⬜ 선택: `am_Utils`·`tpl_Test` 기존 문자열의 `—` 정리

## 2026-10-16 17:16 KST · Claude · 2026-10-16 마무리 — 다음 세션 시작점 · 완료

**현재 상태**
- 소스는 모두 커밋·푸시됨 (이 항목과 같은 커밋)
- 회사 PC 에는 일부 모듈만 반영된 상태 — 각 모듈 3번째 줄 버전을 `INDEX.md` "모듈 버전" 표와 비교해 다른 것만 다시 붙여 넣기
  - 이번 세션 변경 7개 모듈 최신 버전: `2026-10-16 16:32` (am_Automation, am_Inspect, am_Credential, ref_Automation, ref_Inspect, ref_Credential, tpl_Messenger)
  - 나머지 모듈은 버전 줄만 추가됨 (코드 변화 없음 — 급하지 않음)

**다음 할 일**
1. ⬜ 회사 PC: xlam 에 `am_Credential` 추가, `am_Automation`·`am_Inspect` 갱신 / CWB 에 `ref_Credential`·`tpl_Messenger` 추가, `ref_Automation`·`ref_Inspect` 갱신 → 컴파일
2. ⬜ WMC 완전 종료 후 `?WmcAddFriend("Target")` — 로그인 입력 창 → 세 번째 칸 입력 → 친구추가까지 True 확인
3. ⬜ WMC 실제 입력칸에 값 입력이 되는지 (직접 실행 창에 "키보드 입력으로 대체" 가 찍히는지)
4. ⬜ 세 번째 칸(`txLifeDispNo`) 값 성격 확인 → 고정값이면 `WMC_LIFEDISPNO` 상수에 기록
5. ⬜ 친구추가 뒤 이어지는 창 구조 확인 → 다음 단계 자동화
6. ⬜ (선택) "한 번에 확인/자동 교체" 매크로 — 회사 PC 의 VBA 프로젝트 개체 모델 액세스 신뢰 설정·소스 폴더 접근 가능 여부 확인 필요
7. ⬜ (선택) `am_Utils`·`tpl_Test` 기존 문자열의 `—` 정리
8. ⬜ (실사용 전) 자격 증명 관리자 방식은 IT 보안 정책 확인 후 사용 — 지금은 "prompt"(매번 입력)
