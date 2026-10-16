Attribute VB_Name = "ref_Automation"
Option Explicit
' 버전 : 2026-10-16 16:32 · am_Automation 신규 함수 래퍼 추가

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Automation                                          │
' │  역할 : am_Automation Application.Run 래퍼               │
' └─────────────────────────────────────────────────────────┘

' 참고 : 아래 프로시저는 래핑 불가 → 직접 Application.Run 사용
'        - ExecuteKeySequence          : ParamArray 파라미터
'        - ProcessRangeWithKeySequence : ParamArray 파라미터
'        - GetMousePosition            : ByRef Long 파라미터

Private Const REF As String = "corelib.xlam!am_Automation."

' ── Enum 재선언 (CWB 직접 참조용) ────────────────────────────
' xlam 에 정의된 Public Enum 은 CWB 에서 이름으로 접근 불가
' → ref_Automation 에 동일 값으로 재선언하여 CWB 코드에서 상수처럼 사용

Public Enum VirtualKeys
    VK_TAB     = &H9
    VK_RETURN  = &HD
    VK_CONTROL = &H11
    VK_MENU    = &H12
    VK_ESCAPE  = &H1B
    VK_SPACE   = &H20
    VK_LEFT    = &H25
    VK_UP      = &H26
    VK_RIGHT   = &H27
    VK_DOWN    = &H28
End Enum

Public Enum KeyActions
    ACTION_COPY        = 1
    ACTION_PASTE       = 2
    ACTION_TAB         = 3
    ACTION_ENTER       = 4
    ACTION_ALT_TAB     = 5
    ACTION_ESCAPE      = 6
    ACTION_ARROW_DOWN  = 7
    ACTION_ARROW_UP    = 8
    ACTION_ARROW_LEFT  = 9
    ACTION_ARROW_RIGHT = 10
End Enum

' ── 키보드 ───────────────────────────────────────────────────

' 목적   : 단일 키보드 액션 실행 (KeyActions Enum 사용)
' 인수   : action    - 실행할 액션 (KeyActions Enum: ACTION_COPY 등)
'          waitAfter - 액션 후 대기 시간(ms, 기본: 100)
' 예시   : ExecuteKeyAction ACTION_COPY
'          ExecuteKeyAction ACTION_PASTE, 200
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Application.Run REF & "ExecuteKeyAction", CLng(action), waitAfter
End Sub

' ── 마우스 ───────────────────────────────────────────────────

' 목적   : 화면 절대 좌표 위치 클릭
' 인수   : lngX      - 클릭 X 좌표 (픽셀)
'          lngY      - 클릭 Y 좌표 (픽셀)
'          blnLeft   - True: 좌클릭 / False: 우클릭 (기본: True)
'          strDelay  - 클릭 전 대기 시간 문자열 (예: "00:00:01")
'          strWait   - 클릭 후 대기 시간 문자열
' 예시   : ClickAtPosition 500, 300
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    Application.Run REF & "ClickAtPosition", lngX, lngY, blnLeft, strDelay, strWait
End Sub

' 목적   : 특정 시각까지 대기 (Application.Wait 래퍼)
' 인수   : strTime - 대기 종료 시각 문자열 (예: "00:00:02" → 2초 후)
' 예시   : WaitTime "00:00:03"
Public Sub WaitTime(ByVal strTime As String)
    Application.Run REF & "WaitTime", strTime
End Sub

' ── 창 제어 ───────────────────────────────────────────────────

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 SendMessage(BM_CLICK)로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
' 반환   : Boolean - 버튼을 찾아 클릭 메시지를 보냈는지 여부
' 예시   : ClickButtonByText("WMC SCDK", "확인")
Public Function ClickButtonByText(ByVal strWindowTitle As String, ByVal strButtonText As String) As Boolean
    ClickButtonByText = Application.Run(REF & "ClickButtonByText", strWindowTitle, strButtonText)
End Function

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 중심 좌표를 마우스로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
'          strClickDelay  - 좌표 클릭 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 버튼을 찾아 클릭했는지 여부
Public Function ClickButtonByRect(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                  Optional ByVal strClickDelay As String = "00:00:02") As Boolean
    ClickButtonByRect = Application.Run(REF & "ClickButtonByRect", strWindowTitle, strButtonText, strClickDelay)
End Function

' 목적   : ClickButtonByRect 재시도 래퍼 (반복문 등 타이밍이 불안정한 구간용)
' 인수   : strWindowTitle - 대상 창 제목
'          strButtonText  - 찾을 버튼 텍스트
'          lngMaxRetry    - 최대 재시도 횟수 (기본 3)
'          strRetryWait   - 실패 시 재시도 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 클릭 성공 여부
Public Function ClickButtonByRect_Retry(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                        Optional ByVal lngMaxRetry As Long = 3, _
                                        Optional ByVal strRetryWait As String = "00:00:02") As Boolean
    ClickButtonByRect_Retry = Application.Run(REF & "ClickButtonByRect_Retry", strWindowTitle, strButtonText, lngMaxRetry, strRetryWait)
End Function

' 목적   : 다른 프로세스 소유의 창을 강제로 포그라운드로 활성화
' 인수   : strWindowTitle - 대상 창 제목 (완전 일치)
' 반환   : Boolean - 창을 찾아 활성화를 시도했는지 여부
Public Function ForceActivateWindow(ByVal strWindowTitle As String) As Boolean
    ForceActivateWindow = Application.Run(REF & "ForceActivateWindow", strWindowTitle)
End Function

' 목적   : UI Automation 으로 창 안의 버튼을 이름으로 찾아 Invoke
' 인수   : strWindowName - 대상 창 이름 (UIA Name 속성 기준)
'          strButtonName - 찾을 버튼 이름 (UIA Name 속성 기준)
' 반환   : Boolean - Invoke 성공 여부
Public Function ClickButtonByUIA(ByVal strWindowName As String, ByVal strButtonName As String) As Boolean
    ClickButtonByUIA = Application.Run(REF & "ClickButtonByUIA", strWindowName, strButtonName)
End Function

' 목적   : UI Automation 으로 창 안의 항목(트리·목록·탭 항목 등)을 이름으로 찾아 선택
' 인수   : strWindowName - 대상 창 이름 (부분 일치)
'          strItemName   - 항목 이름 (앞뒤 공백 무시 앞부분 일치 / * ? 포함 시 Like 패턴)
'          blnExpand     - 선택 후 하위 항목 펼치기 (기본 False)
'          strType       - 종류 제한 (예: "TreeItem" / 생략 시 전체)
' 반환   : Boolean - 찾아서 선택했는지 여부
' 예시   : SelectItemByUIA("WMC Messenger", "Target", True, "TreeItem")
Public Function SelectItemByUIA(ByVal strWindowName As String, _
                                ByVal strItemName As String, _
                                Optional ByVal blnExpand As Boolean = False, _
                                Optional ByVal strType As String = "") As Boolean
    SelectItemByUIA = Application.Run(REF & "SelectItemByUIA", strWindowName, strItemName, blnExpand, strType)
End Function

' 목적   : UI Automation 으로 창 안의 항목을 이름으로 찾아 펼치거나 접음
' 인수   : strWindowName - 대상 창 이름 (부분 일치)
'          strItemName   - 항목 이름 (SelectItemByUIA 와 같은 일치 규칙)
'          blnExpand     - True=펼치기, False=접기 (기본 True)
'          strType       - 종류 제한 (생략 시 전체)
' 반환   : Boolean - 찾아서 펼치기/접기를 실행했는지 여부
' 예시   : ExpandItemByUIA("WMC Messenger", "Target")
Public Function ExpandItemByUIA(ByVal strWindowName As String, _
                                ByVal strItemName As String, _
                                Optional ByVal blnExpand As Boolean = True, _
                                Optional ByVal strType As String = "") As Boolean
    ExpandItemByUIA = Application.Run(REF & "ExpandItemByUIA", strWindowName, strItemName, blnExpand, strType)
End Function

' 목적   : 항목을 찾아 선택한 뒤 우클릭 메뉴를 열고, 메뉴 항목을 이름으로 찾아 실행
' 인수   : strWindowName - 대상 창 이름 (부분 일치)
'          strItemName   - 우클릭할 항목 이름 (SelectItemByUIA 와 같은 일치 규칙)
'          strMenuName   - 실행할 메뉴 이름 (예: "친구추가")
'          strType       - 항목 종류 제한 (예: "TreeItem" / 생략 시 전체)
'          blnAllowMouse - 메시지로 메뉴가 안 열리면 실제 마우스 우클릭 허용 (기본 True)
' 반환   : Boolean - 메뉴 항목을 실행했는지 여부 (실패 단계는 직접 실행 창에 출력)
' 예시   : ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem")
Public Function ClickContextMenuByUIA(ByVal strWindowName As String, _
                                      ByVal strItemName As String, _
                                      ByVal strMenuName As String, _
                                      Optional ByVal strType As String = "", _
                                      Optional ByVal blnAllowMouse As Boolean = True) As Boolean
    ClickContextMenuByUIA = Application.Run(REF & "ClickContextMenuByUIA", strWindowName, strItemName, strMenuName, strType, blnAllowMouse)
End Function

' 목적   : 프로그램 창이 있으면 앞으로 가져오고, 없으면 실행한 뒤 창(과 준비 항목)이 나타날 때까지 대기
' 인수   : strWindowName - 창 이름 (부분 일치)
'          strExePath    - 실행 파일 전체 경로 (창이 없을 때만 실행 / DumpWindows 의 '실행 파일 경로' 열 참고)
'          strReadyItem  - 준비 완료로 볼 항목 이름 (예: "Target" / 생략 시 창만 확인)
'          lngTimeoutSec - 최대 대기 시간(초, 기본 60)
'          strArgs       - 실행 인수 (생략 가능)
' 반환   : Boolean - 창이 준비되었는지 여부 (실패 단계는 직접 실행 창에 출력)
' 예시   : OpenAppWindow("WMC Messenger", "C:\Program Files\WMC\WMC_Msg.exe", "Target")
Public Function OpenAppWindow(ByVal strWindowName As String, _
                              ByVal strExePath As String, _
                              Optional ByVal strReadyItem As String = "", _
                              Optional ByVal lngTimeoutSec As Long = 60, _
                              Optional ByVal strArgs As String = "") As Boolean
    OpenAppWindow = Application.Run(REF & "OpenAppWindow", strWindowName, strExePath, strReadyItem, lngTimeoutSec, strArgs)
End Function

' 목적   : 창이 화면에 열려 있는지 확인 (숨겨진 창은 False)
' 인수   : strWindowName - 창 이름 (부분 일치 / "=이름" 완전 일치 / "A|B" 여러 이름 중 하나)
' 반환   : Boolean
' 예시   : IsWindowOpen("=Login")
Public Function IsWindowOpen(ByVal strWindowName As String) As Boolean
    IsWindowOpen = Application.Run(REF & "IsWindowOpen", strWindowName)
End Function

' 목적   : 창 안의 요소(항목·버튼·입력칸 등)가 나타날 때까지 대기
' 인수   : strWindowName - 창 이름 (IsWindowOpen 과 같은 규칙)
'          strElement    - 요소 이름 (앞부분 일치 / * ? Like 패턴 / "#아이디" AutomationId)
'          lngTimeoutSec - 최대 대기 시간(초, 기본 30)
' 반환   : Boolean - 시간 안에 나타났는지 여부
' 예시   : WaitElementByUIA("WMC Messenger", "Target", 60)
Public Function WaitElementByUIA(ByVal strWindowName As String, _
                                 ByVal strElement As String, _
                                 Optional ByVal lngTimeoutSec As Long = 30) As Boolean
    WaitElementByUIA = Application.Run(REF & "WaitElementByUIA", strWindowName, strElement, lngTimeoutSec)
End Function

' 목적   : 입력칸에 값을 넣음 (UI Automation 값 입력 우선, 안 되면 키보드 입력)
' 인수   : strWindowName    - 창 이름 (IsWindowOpen 과 같은 규칙)
'          strElement       - 입력칸 (이름 앞부분 일치 / "#아이디" AutomationId - 예: "#teUserID")
'          strValue         - 넣을 값
'          blnAllowKeyboard - 값 입력이 안 되면 키보드 입력으로 대체 (기본 True)
' 반환   : Boolean - 값을 넣었는지 여부
' 예시   : SetValueByUIA("=Login", "#teUserID", "hong")
Public Function SetValueByUIA(ByVal strWindowName As String, _
                              ByVal strElement As String, _
                              ByVal strValue As String, _
                              Optional ByVal blnAllowKeyboard As Boolean = True) As Boolean
    SetValueByUIA = Application.Run(REF & "SetValueByUIA", strWindowName, strElement, strValue, blnAllowKeyboard)
End Function

' 목적   : 버튼 등 요소를 이름 또는 AutomationId 로 찾아 실행(Invoke)
' 인수   : strWindowName - 창 이름 (IsWindowOpen 과 같은 규칙)
'          strElement    - 요소 (이름 앞부분 일치 / "#아이디" AutomationId - 예: "#btnLogIn")
'          strType       - 종류 제한 (예: "Button" / 생략 시 전체)
' 반환   : Boolean - 실행했는지 여부
' 예시   : InvokeByUIA("=Login", "#btnLogIn")
Public Function InvokeByUIA(ByVal strWindowName As String, _
                            ByVal strElement As String, _
                            Optional ByVal strType As String = "") As Boolean
    InvokeByUIA = Application.Run(REF & "InvokeByUIA", strWindowName, strElement, strType)
End Function
