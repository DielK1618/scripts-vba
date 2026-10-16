# TROUBLESHOOTING.claude.md — scripts-vba (Claude)

> 대표 문서 `TROUBLESHOOTING.md`(2026-10-16까지 보관)와 함께 읽는다. 규칙: `Git/CLAUDE.md` 3절.

## TS-V-20261016-1400 · UI Automation `ElementFromPoint` 호출 시 컴파일 오류
작성: 2026-10-16 14:00 KST · Claude

**증상**: `objUIA.ElementFromPoint(udtPt)`(`udtPt As tagPOINT`)가 들어 있는 함수를 처음 호출할 때 VBE 가 중단 상태로 열리고 "컴파일 오류입니다: 사용자 정의 형식은 ByVal로 전달될 수 없습니다" 표시. 같은 모듈의 다른 함수는 이미 정상 실행된 뒤라 원인을 놓치기 쉬움
**원인**: `ElementFromPoint` 는 POINT 구조체를 값으로 받는데 VBA 는 구조체를 값으로 넘길 수 없음. VBA 가 필요할 때 컴파일하는 방식이라 그 함수를 처음 부를 때 오류가 남
**해결**: Win32 `WindowFromPoint` 로 커서 아래 창 핸들 → `ElementFromHandle(ByVal hWnd)` → 하위 요소의 화면 위치(BoundingRectangle)로 좌표를 포함하는 가장 안쪽 요소까지 내려감 (`am_Inspect.prv_HitTest`)
**재발 방지**: 새 모듈은 가져온 뒤 `디버그 > VBAProject 컴파일` 로 전체 컴파일 확인. 구조체를 값으로 받는 COM 메서드(`ElementFromPoint`, `ElementFromPointBuildCache` 등)는 VBA 에서 쓰지 않음
**적용 프로젝트**: corelib (`am_Inspect`)

## TS-V-20261016-1401 · UI Automation `ElementFromHandle(hWnd)` 실패
작성: 2026-10-16 14:01 KST · Claude

**증상**: `objUIA.ElementFromHandle(hWnd)` 가 런타임 오류로 실패
**원인**: 형식 라이브러리에서 인수 형식이 `void*` 라 VBA 가 변수 값이 아니라 변수 주소를 넘김
**해결**: 호출할 때 `ElementFromHandle(ByVal hWnd)` 처럼 `ByVal` 을 명시
**재발 방지**: UIA 의 창 핸들 인수(`UIA_HWND`)에는 항상 `ByVal` 을 붙임
**적용 프로젝트**: corelib (`am_Inspect`)

## TS-V-20261016-1402 · Excel 자기 창을 UI Automation 으로 조회하면 멈춤
작성: 2026-10-16 14:02 KST · Claude

**증상**: 마우스 커서가 Excel 창 위에 있을 때 커서 위치 요소 조회가 응답 없이 멈춤
**원인**: VBA 는 Excel 화면 스레드에서 실행됨. 그 스레드가 매크로를 실행하는 동안 자기 UI 를 UI Automation 으로 조회하면 응답을 처리할 스레드가 없어 멈출 수 있음
**해결**: 조회 전에 Win32 로 대상 창의 프로세스 ID 를 확인해 `GetCurrentProcessId()` 와 같으면 중단. 창 이름으로 찾을 때도 Excel 자기 창은 제외(통합문서 이름에 키워드가 들어가는 경우 대비)
**재발 방지**: UI Automation 조회 대상에서 Excel 자신을 항상 제외
**적용 프로젝트**: corelib (`am_Inspect`)

## TS-V-20261016-1430 · 32bit Excel 에서 `#If Win64` 블록 안 LongLong 리터럴이 빨간 줄
작성: 2026-10-16 14:30 KST · Claude

**증상**: 32bit Excel(2010)에 모듈을 가져오면 `#If Win64 Then` 블록 안의 `&H100000000^` 가 들어간 줄이 빨간색(구문 오류)으로 표시됨
**원인**: `^`(LongLong 리터럴 접미사)는 64bit VBA 에만 있음. 편집기는 실행되지 않는 조건부 블록도 문법 검사를 하므로 32bit 에서 오류로 표시
**해결**: 리터럴 대신 Double 로 계산 후 `CLngLng` 로 변환 (`CLngLng(lngY * 4294967296# + dblLow)`, 음수 x 는 `+ 4294967296#` 로 부호 없는 값 변환). 64bit 에서 왼쪽 보조 모니터(음수 x) 좌표로 동작 확인
**재발 방지**: `#If Win64` 블록 안에서도 `^` 접미사 리터럴은 쓰지 않음. `LongLong` 형식 이름·`CLngLng` 는 32bit 편집기에서도 오류 표시 없음
**적용 프로젝트**: corelib (`am_Inspect.prv_WindowAtPoint`)

## TS-V-20261016-1510 · 커서 위치 요소 조회가 트리 컨트롤에서 멈추고 항목까지 내려가지 않음
작성: 2026-10-16 15:10 KST · Claude

**증상**: `DumpElementAtCursor` 를 트리 목록의 어느 줄 위에서 실행해도 결과 마지막 행이 `Tree` 컨트롤 자체 (`DumpElementTree` 에는 `TreeItem` 들이 정상으로 나옴)
**원인**: 트리 항목의 화면 영역(BoundingRectangle)은 자기 글자 줄뿐이고, 자식 항목은 부모 영역 밖(아래 줄)에 있음. 부모 영역 안의 자식만 따라 내려가는 탐색이 첫 줄 항목에서 끊김
**해결**: 시작 요소 하위 전체를 탐색해 좌표를 포함하는 가장 작은 요소를 고름 (`am_Inspect.prv_HitTest`, 탐색 상한 5000)
**재발 방지**: 화면 위치로 요소를 찾을 때 "부모 영역이 자식을 포함한다"고 가정하지 않음
**적용 프로젝트**: corelib (`am_Inspect`)

## TS-V-20261016-1511 · WinForms 트리 항목 이름이 비어 있고 종류가 `Tree` 로만 나옴 (참고)
작성: 2026-10-16 15:11 KST · Claude

**증상**: 어떤 WinForms 앱은 트리 항목이 이름 없이 `Tree` 로 나와 `SelectItemByUIA` 로 찾을 수 없음 (같은 `SysTreeView32` 라도 WMC Messenger 는 이름·`TreeItem` 으로 정상 노출)
**원인**: .NET Framework 4.7.2 미만을 대상으로 만든 앱은 구형 접근성 방식이 기본 — 항목 이름·종류를 제대로 노출하지 않음 (PowerShell 5 로 띄운 테스트 창에서 확인. 실행 중 `AppContext` 스위치로는 바뀌지 않음)
**해결**: 앱 쪽 설정이라 VBA 에서 바꿀 수 없음. 이런 앱은 `DumpElementTree` 로 항목 영역을 확인한 뒤 좌표 클릭(`ClickAtPosition`)이나 키보드 이동으로 대체
**재발 방지**: 테스트용 WinForms 창은 .NET 4.8 대상으로 컴파일해 실제 앱과 같은 조건에서 검증
**적용 프로젝트**: corelib (`am_Automation`, `am_Inspect`)

## TS-V-20261016-1550 · WinForms 우클릭 메뉴 항목을 UI Automation 으로 찾지 못함
작성: 2026-10-16 15:50 KST · Claude

**증상**: `WM_CONTEXTMENU` 로 메뉴는 화면에 열렸는데, 대상 프로세스의 "메인 창이 아닌 최상위 창"에서 메뉴 항목을 찾으면 없음
**원인**: WinForms `ContextMenuStrip` 은 Win32 로는 별도 최상위 창(`WindowsForms10.Window.20808…` + `SysShadow`)이지만, UI Automation 트리에서는 소유자인 메인 창 아래(`DropDown` 도구 모음)에 붙어 보임
**해결**: 메인 창을 포함해 프로세스의 모든 최상위 창을 찾되, 메뉴를 열기 전 메뉴 항목의 RuntimeId 를 기억해 두고 연 뒤 새로 나타난 항목만 대상 (상단 메뉴 `친구찾기` 등 이름이 비슷한 기존 메뉴와 구분) — `am_Automation.ClickContextMenuByUIA`
**재발 방지**: 팝업·메뉴는 UIA 트리에서 소유자 창 아래에 나타날 수 있음. 위치로 가정하지 말고 열기 전/후 비교로 찾음
**적용 프로젝트**: corelib (`am_Automation`)

## TS-V-20261016-1635 · VBE 출력·코드에서 긴 줄표(—)가 ? 로 바뀜
작성: 2026-10-16 17:16 KST · Claude

**증상**: 직접 실행 창 출력이 `항목을 찾지 못함 ? WMC Messnger` 처럼 `—` 자리에 `?` 로 표시됨
**원인**: VBE 는 한글 Windows 에서 CP949 로 코드를 다룸. `—`(U+2014) 는 CP949 에 없어 붙여 넣기·가져오기 때 `?` 로 바뀜 (`═` U+2550 도 같음)
**해결**: 실행 문자열과 버전 줄에는 `-` 사용. 이번 세션 모듈 7개에서 `—` → `-` 교체
**재발 방지**: `corelib/CLAUDE.md` §5 "VBE 에서 깨지는 문자" 규칙. 새 코드 작성 후 CP949 로 인코딩 안 되는 문자가 문자열에 있는지 확인
**적용 프로젝트**: corelib (기존 `am_Utils`·`tpl_Test` 문자열에는 남아 있음)
