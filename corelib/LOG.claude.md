# LOG.claude.md — corelib 작업 기록 (Claude)

> 대표 문서 `LOG.md`(2026-10-16까지 보관)와 함께 읽는다. 규칙: `Git/CLAUDE.md` 3절.

## 2026-10-16 13:45 KST · Claude · am_Inspect 모듈 신설 (외부 프로그램 UI 요소 인식) · 진행 중

**배경**: 사내 메신저 등 Windows 데스크톱 앱을 마우스·키보드 흉내 없이 직접 제어하려면, 먼저 대상 앱의 UI 요소를 인식할 수 있는지 확인해야 한다. 일회성 진단 코드 대신 다른 프로그램에도 다시 쓸 수 있도록 정식 모듈로 만들기로 결정(사용자 요청).

**추가 파일**
- `xlam/am_Inspect.bas` — UI Automation 조회 전용 모듈
  - `ListWindows` / `DumpWindows` — 최상위 창 목록 (이름, 프로세스, 프로세스ID, 클래스, 프레임워크, 창 핸들, 화면 위치)
  - `GetElementTree` / `DumpElementTree` — 창 안의 요소 트리 (깊이, 위치번호, 이름, 종류, AutomationId, 클래스, 프레임워크, 지원 패턴, 값, 사용 가능, 화면 위치, 창 핸들)
  - `GetElementAtCursor` / `DumpElementAtCursor` — 대기 후 마우스 커서 아래 요소와 최상위 창까지의 상위 경로
  - `Get*` = 2차원 배열 반환(1행 머리글, 실패 시 Empty), `Dump*` = CWB 에 새 시트로 기록(실패 시 MsgBox, 반환 0)
- `cwb/ref_Inspect.bas` — `Application.Run` 래퍼, `Dump*` 는 `ThisWorkbook` 자동 입력

**설계 결정**
| 항목 | 결정 |
|---|---|
| 모듈 경계 | am_Inspect = 인식(조회 전용), am_Automation = 제어(클릭·입력). 제어 함수는 계속 am_Automation 에 추가 |
| 바인딩 | Early Binding(`New CUIAutomation`) — TS-V16 과 같은 이유의 Late Binding 예외. 참조는 xlam 에만 필요 |
| 창 찾기 | 창 이름 부분 일치 → 없으면 프로세스 이름(WMI) 부분 일치로 재탐색 |
| 대형 앱 보호 | `lngMaxDepth`(기본 10), `lngMaxRows`(기본 3000) 상한 |
| 결과 시트 | 이름 열을 깊이만큼 들여쓰기, 문자열은 `'` 접두로 텍스트 고정(숫자·수식 변환 방지), 머리글 필터·틀 고정 |
| 위치번호 | 형제 순번을 점으로 이은 경로(예: `1.3.2`) — 같은 이름 요소 구분용 |
| 결과 쓰기 대상 | 통합문서를 `wb` 인수(맨 끝 Optional)로 명시 전달, 비어 있으면 MsgBox 후 종료 (`ActiveWorkbook` 추정 금지 원칙 준수) |

**문서 반영**: `CLAUDE.md`(폴더 구조·모듈 표·Late Binding 예외), `INDEX.md`(모듈 표)

**검증 중 수정 (이 PC: Excel 16 64bit, 임시 통합문서에 am_Core 대체 모듈 + am_Inspect 가져오기 후 실행)**
| 문제 | 원인 | 수정 |
|---|---|---|
| `GetElementAtCursor` 첫 호출 시 VBE 중단 | `ElementFromPoint(tagPOINT)` → 컴파일 오류 "사용자 정의 형식은 ByVal로 전달될 수 없습니다" (필요할 때 컴파일되는 방식이라 그 함수를 처음 부를 때 발생) | Win32 `WindowFromPoint` 로 커서 아래 창 → `ElementFromHandle` → 화면 위치로 가장 안쪽 요소를 찾아 내려가는 `prv_HitTest` 방식으로 교체 |
| `ElementFromHandle(hWnd)` 실패 | 형식 라이브러리 인수가 `void*` | `ElementFromHandle(ByVal hWnd)` |
| 커서가 Excel 위에 있을 때 멈춤 위험 | 실행 중인 자기 UI 를 UI Automation 으로 조회 | 커서 아래 창 프로세스가 Excel 자신이면 중단, 창 찾기에서도 Excel 자기 창 제외 |
| `Get*` 결과 문자열 앞에 `'` 가 붙음 | 텍스트 고정용 `'` 를 배열 단계에서 붙임 | 시트에 쓰는 단계(`prv_WriteSheet`)에서만 붙이도록 이동 |

**검증 결과 (64bit)**: VBAProject 컴파일 통과, `ListWindows`(창 10개·프로세스 이름), `GetElementTree`(VS Code 창, 창 이름·프로세스 이름 `code.exe` 둘 다로 찾기), `GetElementAtCursor`(AI Curator 창 위), `DumpWindows`/`DumpElementTree`(시트 생성·제목·데이터) 모두 정상. Electron 앱(VS Code, AI Curator)은 예상대로 창 틀만 보이고 본문 요소는 안 보임.

**미검증**: 32bit 전용 코드 경로(`WindowFromPoint(x, y)`), 표준 Windows 앱 위에서 `prv_HitTest` 가 안쪽 요소까지 내려가는지(테스트 중 사용자 마우스를 움직이지 않으려고 생략)

## 2026-10-16 14:30 KST · Claude · am_Inspect 32bit 구문 오류 수정 · 완료

- 회사 PC 확인 결과: Windows 64bit + Excel 32bit (`Application.OperatingSystem` = `Windows (32-bit) NT 6.02`) → `CLAUDE.md` 11절 반영
- 32bit Excel 가져오기 시 `prv_WindowAtPoint` 의 64bit 줄이 빨간 줄로 표시(사용자 보고) → LongLong 리터럴 제거, Double 계산으로 교체 (`../TROUBLESHOOTING.claude.md` TS-V-20261016-1430)
- 이 PC(64bit)에서 재검증: 컴파일 통과, 음수 x 좌표(왼쪽 보조 모니터)에서 커서 조회 정상

## 2026-10-16 15:10 KST · Claude · WMC Messenger 트리 항목 선택 기능 + 커서 조회 수정 · 완료

**회사 PC 결과(사용자 제공)**: WMC Messenger(`WMC_Msg.exe`, WinForms) `사용자` 탭의 `tree_user`(`WindowsForms10.SysTreeView32`) 아래 그룹이 `TreeItem` 으로 노출됨 — 지원 패턴 `SelectionItem, ExpandCollapse`. 단 `DumpElementAtCursor` 는 목록 어디서나 `tree_user` 에서 멈춤

**원인**: 트리 항목의 화면 영역은 자기 글자 줄뿐이고 자식(그룹)은 그 아래 줄에 있음 → "부모 영역 안의 자식만 따라 내려가는" 탐색이 `WMC USER`(첫 줄)에서 끊김

**수정·추가**
- `am_Inspect.prv_HitTest` — 하위 전체를 탐색해 좌표를 포함하는 가장 작은 요소 선택(상한 5000개, 같은 면적이면 더 깊은 쪽)
- `am_Automation.SelectItemByUIA(창, 항목, [펼치기], [종류])` — 이름으로 찾아 선택(SelectionItem, 없으면 Invoke), 화면 밖이면 스크롤 후 선택
- `am_Automation.ExpandItemByUIA(창, 항목, [펼치기/접기], [종류])`
- 이름 일치 규칙: 앞뒤 공백 무시 앞부분 일치(`"Target"` → `"Target (52/212)"`), `*`·`?` 있으면 Like 패턴. 창은 이름 부분 일치, Excel 자기 창 제외
- `ref_Automation` 래퍼 2개 추가, 기존 `ClickButtonByUIA` 래퍼 주석의 "Late Binding" 오기 수정
- `INDEX.md`, `CLAUDE.md`(Early Binding 예외 목록) 반영

**검증 (이 PC, .NET 4.8 WinForms 트리 테스트 프로그램 — WMC 와 같은 노출 구조)**: 그룹 선택+펼치기, 그룹 안 사용자 선택, 접기, Like 패턴, 없는 항목·잘못된 종류·없는 창(False), 커서 조회(그룹 안 사용자 위 → `WMC USER > Target > 이영희`) 모두 정상

## 2026-10-16 15:50 KST · Claude · 우클릭 메뉴 실행 함수 ClickContextMenuByUIA · 완료

**요청**: WMC Messenger 에서 `Target` 그룹 선택 → 우클릭 메뉴 → `친구추가` 실행 (사용자 확인: `SelectItemByUIA` 회사 PC 정상 동작)

**추가**: `am_Automation.ClickContextMenuByUIA(창, 항목, 메뉴, [종류], [마우스 허용=True])` + `ref_Automation` 래퍼
1. 항목 찾기·선택 (`SelectItemByUIA` 와 같은 규칙)
2. 메뉴 열기 — 대상 창 활성화 → 항목을 담은 컨트롤(트리) 창에 `WM_CONTEXTMENU`(항목 중심 화면 좌표) 전송. 1.5초 안에 안 열리면 마우스 허용 시 항목 위치를 매번 계산해 실제 우클릭 후 커서 원위치
3. 메뉴 항목 실행 — 열기 전 메뉴 항목(RuntimeId)을 기억하고 새로 나타난 항목 중 이름이 맞는 것을 Invoke. 사용 불가(흐린) 항목·Invoke 미지원은 실행 안 하고 Esc 로 메뉴 닫음
- 실패 단계는 `Debug.Print` 로 직접 실행 창에 출력
- `ForceActivateWindow` 의 활성화 코드를 `prv_ActivateHwnd` 로 분리해 공용화 (동작 변화 없음)
- `prv_UIA_FindItem` 이 찾은 최상위 창도 돌려주도록 인수 추가 (내부 함수)

**발견**: WinForms 우클릭 메뉴(ContextMenuStrip)는 Win32 로는 별도 최상위 창이지만 UIA 에서는 메인 창 아래에 붙어 보임 → "팝업 창에서만 찾기" 첫 구현이 모두 실패. 메인 창 상단 메뉴(친구찾기 등)와 섞이지 않도록 열기 전/후 비교 방식으로 수정 (`../TROUBLESHOOTING.claude.md` TS-V-20261016-1550)

**검증 (이 PC, .NET 4.8 WinForms 테스트 프로그램, 메뉴 연결 방식 2가지)**
| 경우 | 결과 |
|---|---|
| ContextMenuStrip 속성 연결 + 마우스 금지: 친구추가 / 쪽지(앞부분 일치) | True, 실행됨 |
| 사용 불가 메뉴 / 없는 메뉴 | False, 메뉴 닫힘 |
| 마우스 우클릭 이벤트에서만 메뉴 표시 + 마우스 금지 | False (예상대로) |
| 같은 앱 + 마우스 허용 | True, 실행됨 |

**미확인**: WMC 실제 앱이 어느 방식으로 메뉴를 띄우는지, 회사 PC(Excel 2010 32bit) 화면 배율이 100%가 아닐 때 마우스 대체 경로의 좌표 정확도

## 2026-10-16 16:30 KST · Claude · 프로그램 실행 함수 OpenAppWindow + 실행 파일 경로 열 · 완료

**회사 PC 결과(사용자 제공)**: `ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem")` → 메시지로는 메뉴가 안 열려 마우스 우클릭으로 대체 → True. WMC 는 실제 마우스 우클릭에서만 메뉴를 띄우는 방식으로 판단. (첫 시도의 False 는 창 이름 오타 `WMC Messnger`)

**추가**
- `am_Automation.OpenAppWindow(창, 실행 파일, [준비 항목], [대기 초=60], [실행 인수])` + `ref_Automation` 래퍼
  - 창이 있으면 앞으로(최소화 복원, 숨겨진 창은 제목 표시줄 있는 창만 다시 표시), 없으면 실행 파일 폴더를 작업 폴더로 실행 후 창 대기
  - 준비 항목을 주면 UIA 로 그 항목이 나타날 때까지 대기 (WMC "Making a list of users..." 로딩 대비)
  - 창 찾기는 Win32 `EnumWindows`(숨겨진 창 포함, 제목 부분 일치, Excel 자기 창 제외)
- `am_Inspect.ListWindows`/`DumpWindows` 에 '실행 파일 경로' 열 추가 (WMI `ExecutablePath`, 권한 없는 프로세스는 빈 값)

**검증 (이 PC, WmcTest.exe)**: 미실행→실행+Target 대기, 실행 중 재호출(프로세스 1개 유지), 최소화→복원, 숨김→다시 표시, 잘못된 경로·경로 비움·없는 준비 항목(False) 모두 정상. 경로 열에 실행 파일 전체 경로 표시 확인

**작업 중 실수**: 패치 스크립트(Python 삼중 따옴표 문자열) 안의 VBA `""""` 가 깨져 실행 줄의 따옴표가 사라짐 → 컴파일 오류로 발견, `Chr(34)` 로 수정

## 2026-10-16 16:00 KST · Claude · 로그인 자동화용 범용 함수 · 완료

**회사 PC 결과(사용자 제공)**: `OpenAppWindow(...,"Target")` True, `ClickContextMenuByUIA` 마우스 대체로 True. WMC 로그인 창 구조 — 제목 `Login`(frmLogin), 입력칸·버튼 이름이 비어 있고 AutomationId 로만 구분: `teUserID`, `tePassword`, `txLifeDispNo`(용도 미확인), `btnLogIn` / 입력칸 `Value`, 버튼 `Invoke` 지원 / 자동 로그인 체크 상자 없음

**추가·변경 (am_Automation + ref_Automation 래퍼)**
- 요소 일치 규칙에 `"#아이디"`(AutomationId 완전 일치) 추가 — 모든 UIA 함수 공통
- 창 이름 규칙에 `"=이름"`(완전 일치), `"A|B"`(여러 이름 중 하나) 추가 — `Login` 처럼 짧은 제목이 브라우저 탭 등과 섞이지 않게
- `IsWindowOpen(창)`, `WaitElementByUIA(창, 요소, [초])`, `SetValueByUIA(창, 입력칸, 값, [키보드 허용])`, `InvokeByUIA(창, 요소, [종류])`
- `SetValueByUIA` 순서: 요소 자신 Value 입력 → 안쪽 입력 요소(외부 라이브러리 입력칸 대비, RawView 3단계) → 키보드(SendKeys, 특수문자 이스케이프). 값을 다시 읽어 확인(비밀번호 칸 제외)

**검증 (이 PC, 로그인 창 → 2초 뒤 목록이 채워지는 메인 창 테스트 앱)**: `OpenAppWindow("…|=Login")` → 로그인 창 인식, `=Log` 완전 일치 아님 False, 틀린 비밀번호 거부(앱 로그 확인), 특수문자 비밀번호 `p+w{1}%` 정확히 입력·로그인, `WaitElementByUIA` 로 목록 대기, 친구추가 실행, 없는 요소·입력칸 False

**참고**: PowerShell(관리형) UI Automation 은 이 앱의 AutomationId 를 숫자로만 보여 검증에 못 씀 → 테스트 앱이 자체 로그 파일에 기록하는 방식으로 확인

**결정 대기**: 비밀번호 보관 방식, `txLifeDispNo` 입력칸 용도

## 2026-10-16 16:20 KST · Claude · 로그인 정보 모듈 am_Credential + WMC 흐름 tpl_Messenger · 완료

**사용자 결정 (2026-10-16)**: 기능은 모두 구현하되 실사용은 "실행할 때마다 입력". 자격 증명 관리자 방식은 만들어 두고, 보안 문제가 생기면 사용자가 알려 주기로 함. 사용법은 주석으로

**추가**
- `xlam/am_Credential.bas` (외부 참조 없음, Windows API) + `cwb/ref_Credential.bas`
  - `PromptCredential` — Windows 자격 증명 입력 창(`CredUIPromptForCredentialsW`, 비밀번호 ● 표시, 저장 안 함) → `Array(아이디, 비밀번호)` / 취소 시 Empty
  - `SaveCredential` / `ReadCredential` / `HasCredential` / `DeleteCredential` — 일반 자격 증명, 이 PC·이 사용자 범위(`CRED_PERSIST_LOCAL_MACHINE`)
- `cwb/tpl_Messenger.bas` — WMC 전용 흐름, 상단 주석에 사용법·로그인 방식·화면 구조 기록
  - `WmcOpen` (실행 → 로그인 창이면 로그인 → 목록 대기), `WmcAddFriend(그룹)`, `WmcRunGroupMenu(그룹, 메뉴)`, `WmcSaveCredential` / `WmcDeleteCredential` / `WmcHasCredential`
  - 로그인 방식 상수 `WMC_LOGIN_MODE` = "prompt"(기본) / "stored" / "manual", 세 번째 칸 `WMC_LIFEDISPNO`(비우면 매번 InputBox)
- `CLAUDE.md`, `INDEX.md` 모듈 표 반영

**검증 (이 PC, 실제 구조와 같게 — am_ 모듈을 `corelib.xlam` 추가 기능으로 저장·`Workbooks.Open` 로드, CWB 에 ref_·tpl_ 모듈, tpl 상수만 테스트 앱으로 교체)**
- 자격 증명: 저장·확인·읽기(특수문자 `p+w{1}%`, 한글·유니코드 아이디/비밀번호)·삭제·없는 이름 Empty — 모두 정상, 테스트 항목 삭제 확인
- 입력 창: 표시·기본 아이디 채움·확인 시 배열 반환 확인. 비밀번호 칸 자동 입력은 테스트 도구로 못 넣어 비밀번호 값 자체는 미검증
- 전체 흐름("stored" 방식, 앱 꺼진 상태에서): 실행 → `Login` 창 인식 → 세 칸 입력(앱 로그 `id=user pw=p+w{1}% no=1234`) → 로그인 → 2초 뒤 목록 대기 → 친구추가 → 재호출 시 재로그인 없이 쪽지 보내기 — 모두 True

**테스트 중 발견(하네스 쪽)**: Excel COM 의 `SaveAs(…, 55)` 직후 통합 문서 이름이 바뀌지 않아 `corelib.xlam!` 호출이 실패 → 저장 후 닫고 `Workbooks.Open` 으로 다시 여는 방식으로 해결. 셀 주소와 같은 매크로 이름(`T1`)은 `Application.Run` 이 실행하지 못함

## 2026-10-16 16:35 KST · Claude · 모듈 버전 줄 도입 + VBE 깨짐 문자 정리 · 완료

**사용자 결정 (2026-10-16)**: corelib 전체 모듈에 "주석 한 줄 + 버전 표" 방식으로 버전 표시. (사용자는 VBE 에 코드를 복사·붙여넣기로 반영 중)

- 모든 모듈(xlam 15, cwb 19 — ThisWorkbook 포함) `Option Explicit` 다음 줄에 `' 버전 : YYYY-MM-DD HH:mm · 변경 요약` 추가
  - 이번 세션에 고친 7개 모듈은 16:32, 나머지는 git 마지막 커밋 시각
- `INDEX.md` 에 "모듈 버전" 표, `CLAUDE.md` §5 에 "버전 줄" 규칙과 "VBE 에서 깨지는 문자" 규칙 추가
- 발견: 실행 문자열의 `—`(U+2014)는 CP949 에 없어 VBE 에서 `?` 로 바뀜 (회사 PC 출력 "항목을 찾지 못함 ? WMC Messnger" 의 원인) → 이번 세션 모듈 7개(am_Automation, am_Inspect, am_Credential, ref_Automation, ref_Inspect, ref_Credential, tpl_Messenger)에서 `-` 로 교체
- 미처리(기존 코드): `am_Utils.bas:545` MsgBox, `tpl_Test.bas` 의 Debug.Print·PrintBool 문자열 다수에 `—` 남음 — 출력에 `?` 로 보이는 것 외 동작 영향 없음
- 검증: xlam 16개·cwb 19개 모듈 전체 가져오기 후 VBAProject 컴파일 통과

## 2026-10-16 17:16 KST · Claude · 2026-10-16 작업 마무리 · 완료

**오늘 결과 요약**
- 외부 프로그램 UI 인식 모듈 `am_Inspect` 신설 (창 목록·요소 트리·커서 위치 요소, 실행 파일 경로)
- `am_Automation` 에 UI Automation 제어 함수 추가: 항목 선택·펼치기, 우클릭 메뉴 실행, 프로그램 실행·창 대기, 요소 대기, 입력칸 값 입력, 버튼 실행 (`#아이디`·`=이름`·`A|B` 규칙)
- 로그인 정보 모듈 `am_Credential` 신설 (입력 창, 자격 증명 관리자)
- CWB: `ref_Inspect`, `ref_Credential` 신설, `ref_Automation` 래퍼 추가, WMC 흐름 `tpl_Messenger` 신설
- 전체 모듈 버전 줄 도입 + `INDEX.md` 모듈 버전 표
- 회사 PC 확인 완료: `DumpWindows`/`DumpElementTree`/`DumpElementAtCursor`, `SelectItemByUIA`, `OpenAppWindow`, `ClickContextMenuByUIA`(마우스 대체로 친구추가 성공)
- 이 PC 검증(테스트 앱): 로그인 → 목록 대기 → 친구추가 전체 흐름, 자격 증명 저장·읽기·삭제
