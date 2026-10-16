Attribute VB_Name = "am_Automation"
Option Explicit
' 버전 : 2026-10-16 16:32 · UIA 항목 선택·우클릭 메뉴·프로그램 실행·입력칸/버튼 함수 추가

' +---------------------------------------------------------+
' |  am_Automation                                          |
' |  역할 : 키보드/마우스 입력 시뮬레이션, 창 탐색/활성화/버튼 제어 |
' +---------------------------------------------------------+
' ※ 참조 필요: UIA 계열(…ByUIA, OpenAppWindow - Early Binding) 사용을 위해 VBE 도구 > 참조에서
'   "UIAutomationClient" 라이브러리를 체크해야 함 (프로젝트 Late Binding 원칙의 명시적 예외 - TROUBLESHOOTING.md TS-V16)

' -- Windows API --
Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" _
        (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function EnumChildWindows Lib "user32" _
        (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowTextA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function SendMessageA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" _
        (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" _
        (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function PostMessageA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    lngLeft   As Long
    lngTop    As Long
    lngRight  As Long
    lngBottom As Long
End Type

Private Const KEYEVENTF_KEYDOWN     As Long = 0
Private Const KEYEVENTF_KEYUP       As Long = 2
Private Const MOUSEEVENTF_LEFTDOWN  As Long = &H2
Private Const MOUSEEVENTF_LEFTUP    As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

Private Const BM_CLICK   As Long = &HF5
Private Const SW_RESTORE As Long = 9
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const SW_SHOW        As Long = 5
Private Const GWL_STYLE      As Long = -16
Private Const WS_CAPTION     As Long = &HC00000

' UI Automation 속성/패턴 ID
' ※ Early Binding 사용 - VBE 참조에 "UIAutomationClient" 라이브러리 추가 필수 (프로젝트 Late Binding 원칙의 명시적 예외)
'   CreateObject("UIAutomationClient.CUIAutomation") Late Binding은 이 환경(Excel 2010 32bit)에서
'   런타임 429 오류(ActiveX 구성 요소는 개체를 만들 수 없습니다)로 실패 확인됨 → TROUBLESHOOTING.md TS-V16 참고
Private Const UIA_NamePropertyId        As Long = 30005
Private Const UIA_InvokePatternId       As Long = 10000
Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_ButtonControlTypeId   As Long = 50000
Private Const TreeScope_Descendants     As Long = 4
Private Const UIA_ExpandCollapsePatternId As Long = 10005
Private Const UIA_SelectionItemPatternId  As Long = 10010
Private Const UIA_ScrollItemPatternId     As Long = 10017
Private Const UIA_MAX_SEARCH_NODES        As Long = 5000   ' 항목 탐색 상한 (대형 창 멈춤 방지)
Private Const UIA_MenuItemControlTypeId   As Long = 50011
Private Const UIA_ValuePatternId          As Long = 10002
Private Const UIA_BoundingRectanglePropertyId    As Long = 30001
Private Const UIA_NativeWindowHandlePropertyId   As Long = 30020

' 버튼 탐색 결과 임시 저장 (prv_EnumChildProc 콜백 전용)
Private m_hWndFound     As LongPtr
Private m_strTargetText As String

' 최상위 창 탐색 결과 임시 저장 (prv_EnumTopProc 콜백 전용)
Private m_strTopTitle     As String
Private m_hWndTopVisible  As LongPtr
Private m_hWndTopHidden   As LongPtr

' 가상 키코드
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

' 키 동작 정의
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

' ==========================================================
'  키보드
' ==========================================================

' 목적   : KeyActions Enum에 정의된 키 동작 실행
' 인수   : action    - 실행할 키 동작 (KeyActions Enum 또는 Long)
'          waitAfter - 실행 후 대기 ms (기본 100)
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Select Case action
        Case ACTION_COPY        : prv_SendKeyCombo VK_CONTROL, 67   ' Ctrl+C
        Case ACTION_PASTE       : prv_SendKeyCombo VK_CONTROL, 86   ' Ctrl+V
        Case ACTION_TAB         : prv_SendKey VK_TAB
        Case ACTION_ENTER       : prv_SendKey VK_RETURN
        Case ACTION_ALT_TAB     : prv_SendKeyCombo VK_MENU, VK_TAB
        Case ACTION_ESCAPE      : prv_SendKey VK_ESCAPE
        Case ACTION_ARROW_DOWN  : prv_SendKey VK_DOWN
        Case ACTION_ARROW_UP    : prv_SendKey VK_UP
        Case ACTION_ARROW_LEFT  : prv_SendKey VK_LEFT
        Case ACTION_ARROW_RIGHT : prv_SendKey VK_RIGHT
    End Select

    If waitAfter > 0 Then prv_WaitMs waitAfter
End Sub

' 목적   : 여러 키 동작을 순서대로 실행
' 인수   : actions - KeyActions 또는 Array(KeyActions, waitMs) 의 ParamArray
Public Sub ExecuteKeySequence(ParamArray actions() As Variant)
    Dim i As Long

    For i = LBound(actions) To UBound(actions)
        If IsArray(actions(i)) Then
            ExecuteKeyAction actions(i)(0), actions(i)(1)
        Else
            ExecuteKeyAction actions(i)
        End If
    Next i
End Sub

' 목적   : 범위의 각 셀에 키 시퀀스 적용
' 인수   : rng     - 대상 범위
'          actions - ExecuteKeySequence에 전달할 동작 ParamArray
Public Sub ProcessRangeWithKeySequence(ByVal rng As Range, _
                                       ParamArray actions() As Variant)
    Dim cel As Range

    For Each cel In rng
        If Not IsEmpty(cel) Then
            cel.Select
            ExecuteKeySequence actions
        End If
    Next cel
End Sub

Private Sub prv_SendKey(ByVal bKeyCode As Byte, _
                        Optional ByVal blnRelease As Boolean = True)
    keybd_event bKeyCode, 0, KEYEVENTF_KEYDOWN, 0
    If blnRelease Then keybd_event bKeyCode, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub prv_SendKeyCombo(ParamArray keyCodes() As Variant)
    Dim i As Long

    For i = LBound(keyCodes) To UBound(keyCodes)
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYDOWN, 0
    Next i

    prv_WaitMs 50

    For i = UBound(keyCodes) To LBound(keyCodes) Step -1
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYUP, 0
    Next i
End Sub

' 목적   : ms 단위 대기 (키보드/마우스 내부 공용)
Private Sub prv_WaitMs(ByVal lngMs As Long)
    Application.Wait Now + TimeSerial(0, 0, lngMs / 1000)
End Sub

' ==========================================================
'  마우스
' ==========================================================

' 목적   : 현재 마우스 커서 위치 반환
' 인수   : lngX - (반환) X 좌표
'          lngY - (반환) Y 좌표
Public Sub GetMousePosition(ByRef lngX As Long, ByRef lngY As Long)
    Dim pt As POINTAPI
    GetCursorPos pt
    lngX = pt.x
    lngY = pt.y
End Sub

' 목적   : 마우스를 지정 좌표로 이동 후 클릭
' 인수   : lngX     - X 좌표
'          lngY     - Y 좌표
'          blnLeft  - True=왼쪽 클릭, False=오른쪽 클릭 (기본 True)
'          strDelay - 이동 후 클릭 전 대기 시간 "HH:MM:SS" (생략 시 생략)
'          strWait  - 클릭 후 대기 시간 "HH:MM:SS" (생략 시 생략)
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    SetCursorPos lngX, lngY

    If strDelay <> "" Then WaitTime strDelay

    If blnLeft Then
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Else
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    End If

    If strWait <> "" Then WaitTime strWait
End Sub

' 목적   : 입력된 시간만큼 대기 (Application.Wait 기반)
' 인수   : strTime - 대기 시간 "HH:MM:SS" 형식 (예: "00:00:03" = 3초)
Public Sub WaitTime(ByVal strTime As String)
    Application.Wait Now + TimeValue(strTime)
End Sub

' ==========================================================
'  창 제어
' ==========================================================

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 SendMessage(BM_CLICK)로 클릭
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
' 반환   : Boolean - 버튼을 찾아 클릭 메시지를 보냈는지 여부
' 예시   : ClickButtonByText("WMC SCDK", "확인")
Public Function ClickButtonByText(ByVal strWindowTitle As String, ByVal strButtonText As String) As Boolean
    Dim hWndParent As LongPtr

    hWndParent = FindWindow(vbNullString, strWindowTitle)
    If hWndParent = 0 Then
        ClickButtonByText = False
        Exit Function
    End If

    m_hWndFound = 0
    m_strTargetText = strButtonText
    EnumChildWindows hWndParent, AddressOf prv_EnumChildProc, 0

    If m_hWndFound = 0 Then
        ClickButtonByText = False
        Exit Function
    End If

    SendMessageA m_hWndFound, BM_CLICK, 0, 0
    ClickButtonByText = True
End Function

' 목적   : 창 제목으로 자식 컨트롤(버튼 등)을 텍스트로 찾아 중심 좌표를 마우스로 클릭
'          (SendMessage(BM_CLICK)가 통하지 않는 커스텀/논표준 버튼용)
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
'          strButtonText  - 찾을 버튼/컨트롤 텍스트 (부분 일치)
'          strClickDelay  - 좌표 클릭 전 대기 시간 "HH:MM:SS" (기본 "00:00:02")
' 반환   : Boolean - 버튼을 찾아 클릭했는지 여부
Public Function ClickButtonByRect(ByVal strWindowTitle As String, ByVal strButtonText As String, _
                                  Optional ByVal strClickDelay As String = "00:00:02") As Boolean
    Dim hWndParent As LongPtr
    Dim udtRect    As RECT
    Dim lngCenterX As Long, lngCenterY As Long

    hWndParent = FindWindow(vbNullString, strWindowTitle)
    If hWndParent = 0 Then
        ClickButtonByRect = False
        Exit Function
    End If

    m_hWndFound = 0
    m_strTargetText = strButtonText
    EnumChildWindows hWndParent, AddressOf prv_EnumChildProc, 0

    If m_hWndFound = 0 Then
        ClickButtonByRect = False
        Exit Function
    End If

    GetWindowRect m_hWndFound, udtRect
    lngCenterX = (udtRect.lngLeft + udtRect.lngRight) \ 2
    lngCenterY = (udtRect.lngTop + udtRect.lngBottom) \ 2

    ClickAtPosition lngCenterX, lngCenterY, strDelay:=strClickDelay
    ClickButtonByRect = True
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
    Dim lngRetryCount As Long

    For lngRetryCount = 1 To lngMaxRetry
        If ClickButtonByRect(strWindowTitle, strButtonText) Then
            ClickButtonByRect_Retry = True
            Exit Function
        End If
        WaitTime strRetryWait
    Next lngRetryCount

    ClickButtonByRect_Retry = False
End Function

' 목적   : 다른 프로세스 소유의 창을 강제로 포그라운드로 활성화 (AttachThreadInput 사용)
' 인수   : strWindowTitle - 대상 창 제목 (FindWindow 기준, 완전 일치)
' 반환   : Boolean - 창을 찾아 활성화를 시도했는지 여부
Public Function ForceActivateWindow(ByVal strWindowTitle As String) As Boolean
    Dim hWndTarget As LongPtr

    hWndTarget = FindWindow(vbNullString, strWindowTitle)
    If hWndTarget = 0 Then
        ForceActivateWindow = False
        Exit Function
    End If

    prv_ActivateHwnd hWndTarget
    ForceActivateWindow = True
End Function

' 목적   : 창 핸들로 강제 포그라운드 활성화 (AttachThreadInput - ForceActivateWindow·ClickContextMenuByUIA 공용)
Private Sub prv_ActivateHwnd(ByVal hWndTarget As LongPtr)
    Dim hWndForeground      As LongPtr
    Dim lngForegroundThread As Long, lngCurrentThread As Long
    Dim lngDummy            As Long

    hWndForeground = GetForegroundWindow()
    lngForegroundThread = GetWindowThreadProcessId(hWndForeground, lngDummy)
    lngCurrentThread = GetCurrentThreadId()

    AttachThreadInput lngCurrentThread, lngForegroundThread, True
    ShowWindow hWndTarget, SW_RESTORE
    BringWindowToTop hWndTarget
    SetForegroundWindow hWndTarget
    AttachThreadInput lngCurrentThread, lngForegroundThread, False
End Sub

' 목적   : 창이 화면에 열려 있는지 확인 (숨겨진 창은 False)
' 인수   : strWindowName - 창 이름 (부분 일치 / "=이름" 완전 일치 / "A|B" 여러 이름 중 하나)
' 반환   : Boolean
' 예시   : IsWindowOpen("=Login")
Public Function IsWindowOpen(ByVal strWindowName As String) As Boolean
    Dim blnVisible As Boolean

    If prv_FindTopHwnd(strWindowName, blnVisible) <> 0 Then IsWindowOpen = blnVisible
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
    Dim objWin  As IUIAutomationElement
    Dim dtLimit As Date

    On Error GoTo ErrHandler

    dtLimit = Now + TimeSerial(0, 0, lngTimeoutSec)
    Do
        If Not prv_UIA_FindItem(strWindowName, strElement, "", objWin) Is Nothing Then
            WaitElementByUIA = True
            Exit Function
        End If
        If Now > dtLimit Then Exit Do
        DoEvents
        Sleep 500
    Loop
    Debug.Print "am_Automation.WaitElementByUIA: " & lngTimeoutSec & "초 안에 나타나지 않음 - " & strWindowName & " / " & strElement
    Exit Function

ErrHandler:
    WaitElementByUIA = False
End Function

' 목적   : 입력칸에 값을 넣음 (마우스·키보드 없이 UI Automation 값 입력 우선)
' 인수   : strWindowName    - 창 이름 (IsWindowOpen 과 같은 규칙)
'          strElement       - 입력칸 (이름 앞부분 일치 / "#아이디" AutomationId - 예: "#teUserID")
'          strValue         - 넣을 값
'          blnAllowKeyboard - 값 입력이 안 되면 입력칸에 포커스를 주고 키보드 입력으로 대체 (기본 True)
' 반환   : Boolean - 값을 넣었는지 여부 (비밀번호 칸은 값을 다시 읽을 수 없어 오류 없이 끝나면 성공)
' 예시   : SetValueByUIA("=Login", "#teUserID", "hong")
' 순서   : ① 요소 자신에 값 입력 → ② 안쪽 입력 요소에 값 입력 → ③ 키보드 입력 (blnAllowKeyboard)
'          ①② 는 값을 다시 읽어 같은지 확인 (비밀번호 칸 제외)
Public Function SetValueByUIA(ByVal strWindowName As String, _
                              ByVal strElement As String, _
                              ByVal strValue As String, _
                              Optional ByVal blnAllowKeyboard As Boolean = True) As Boolean
    Const PROC As String = "am_Automation.SetValueByUIA: "

    On Error GoTo ErrHandler

    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim objWin    As IUIAutomationElement
    Dim objEl     As IUIAutomationElement

    Set objEl = prv_UIA_FindItem(strWindowName, strElement, "", objWin)
    If objEl Is Nothing Then
        Debug.Print PROC & "입력칸을 찾지 못함 - " & strWindowName & " / " & strElement
        Exit Function
    End If

    ' ① 요소 자신
    If prv_UIA_TrySetValue(objEl, strValue) Then
        SetValueByUIA = True
        Exit Function
    End If

    ' ② 안쪽 입력 요소 (외부 컨트롤 라이브러리 입력칸은 실제 입력 영역이 자식으로 있음)
    Set objUIA = New CUIAutomation
    Set objWalker = objUIA.RawViewWalker
    If prv_UIA_SetValueInChildren(objWalker, objEl, strValue, 0) Then
        SetValueByUIA = True
        Exit Function
    End If

    ' ③ 키보드 입력
    If Not blnAllowKeyboard Then
        Debug.Print PROC & "값 입력을 지원하지 않음 (blnAllowKeyboard=False 라 키보드 입력은 시도 안 함) - " & strElement
        Exit Function
    End If

    Debug.Print PROC & "값 입력이 안 되어 키보드 입력으로 대체 - " & strElement
    prv_ActivateHwnd prv_UIA_Hwnd(objWin)
    objEl.SetFocus
    Sleep 150
    SendKeys "{HOME}+{END}{DEL}", True
    SendKeys prv_EscapeSendKeys(strValue), True
    SetValueByUIA = True
    Exit Function

ErrHandler:
    Debug.Print PROC & "오류 " & Err.Number & ": " & Err.Description
    SetValueByUIA = False
End Function

' 목적   : 버튼 등 요소를 이름 또는 AutomationId 로 찾아 실행(Invoke)
' 인수   : strWindowName - 창 이름 (IsWindowOpen 과 같은 규칙)
'          strElement    - 요소 (이름 앞부분 일치 / "#아이디" AutomationId - 예: "#btnLogIn")
'          strType       - 종류 제한 (예: "Button" / 생략 시 전체)
' 반환   : Boolean - 실행했는지 여부
' 예시   : InvokeByUIA("=Login", "#btnLogIn")
' 참고   : ClickButtonByUIA 는 버튼 이름 완전 일치만 지원 - 이름이 비어 있는 버튼은 이 함수 사용
Public Function InvokeByUIA(ByVal strWindowName As String, _
                            ByVal strElement As String, _
                            Optional ByVal strType As String = "") As Boolean
    Const PROC As String = "am_Automation.InvokeByUIA: "

    On Error GoTo ErrHandler

    Dim objWin    As IUIAutomationElement
    Dim objEl     As IUIAutomationElement
    Dim objInvoke As IUIAutomationInvokePattern

    Set objEl = prv_UIA_FindItem(strWindowName, strElement, strType, objWin)
    If objEl Is Nothing Then
        Debug.Print PROC & "요소를 찾지 못함 - " & strWindowName & " / " & strElement
        Exit Function
    End If

    Set objInvoke = objEl.GetCurrentPattern(UIA_InvokePatternId)
    If objInvoke Is Nothing Then
        Debug.Print PROC & "실행(Invoke)을 지원하지 않는 요소 - " & strElement
        Exit Function
    End If
    objInvoke.Invoke

    InvokeByUIA = True
    Exit Function

ErrHandler:
    Debug.Print PROC & "오류 " & Err.Number & ": " & Err.Description
    InvokeByUIA = False
End Function

' 목적   : Value 패턴으로 값 입력 후 확인 (읽기 전용·미지원이면 False)
Private Function prv_UIA_TrySetValue(ByVal objEl As IUIAutomationElement, ByVal strValue As String) As Boolean
    Dim objValue As IUIAutomationValuePattern

    On Error GoTo ErrHandler

    Set objValue = objEl.GetCurrentPattern(UIA_ValuePatternId)
    If objValue Is Nothing Then Exit Function
    If objValue.CurrentIsReadOnly <> 0 Then Exit Function

    objValue.SetValue strValue

    If objEl.CurrentIsPassword <> 0 Then
        prv_UIA_TrySetValue = True           ' 비밀번호 칸은 값을 읽을 수 없음
    Else
        prv_UIA_TrySetValue = (objValue.CurrentValue = strValue)
    End If
    Exit Function

ErrHandler:
    prv_UIA_TrySetValue = False
End Function

' 목적   : 하위 요소(최대 3단계) 중 값 입력이 되는 첫 요소에 입력
Private Function prv_UIA_SetValueInChildren(ByVal objWalker As IUIAutomationTreeWalker, _
                                            ByVal objEl As IUIAutomationElement, _
                                            ByVal strValue As String, _
                                            ByVal lngDepth As Long) As Boolean
    Dim objChild As IUIAutomationElement
    Dim objNext  As IUIAutomationElement

    If lngDepth >= 3 Then Exit Function

    On Error Resume Next
    Set objChild = objWalker.GetFirstChildElement(objEl)
    Do Until objChild Is Nothing
        If prv_UIA_TrySetValue(objChild, strValue) Then
            prv_UIA_SetValueInChildren = True
            Exit Function
        End If
        If prv_UIA_SetValueInChildren(objWalker, objChild, strValue, lngDepth + 1) Then
            prv_UIA_SetValueInChildren = True
            Exit Function
        End If
        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objChild)
        Set objChild = objNext
    Loop
End Function

' 목적   : SendKeys 특수문자(+ ^ % ~ ( ) { } [ ]) 를 글자 그대로 입력되도록 감쌈
Private Function prv_EscapeSendKeys(ByVal strText As String) As String
    Dim i       As Long
    Dim strChar As String
    Dim strOut  As String

    For i = 1 To Len(strText)
        strChar = Mid(strText, i, 1)
        If InStr("+^%~(){}[]", strChar) > 0 Then
            strOut = strOut & "{" & strChar & "}"
        Else
            strOut = strOut & strChar
        End If
    Next i
    prv_EscapeSendKeys = strOut
End Function

' 목적   : 창 제목 일치 규칙 - "A|B" 는 여러 이름 중 하나, "=이름" 은 완전 일치, 그 외 부분 일치 (대소문자 무시)
Private Function prv_TitleMatch(ByVal strTitle As String, ByVal strPattern As String) As Boolean
    Dim arrParts As Variant
    Dim i        As Long
    Dim strPart  As String

    If Len(strTitle) = 0 Then Exit Function
    arrParts = Split(strPattern, "|")
    For i = LBound(arrParts) To UBound(arrParts)
        strPart = Trim(arrParts(i))
        If Len(strPart) > 0 Then
            If Left(strPart, 1) = "=" Then
                If StrComp(strTitle, Mid(strPart, 2), vbTextCompare) = 0 Then prv_TitleMatch = True
            ElseIf InStr(1, strTitle, strPart, vbTextCompare) > 0 Then
                prv_TitleMatch = True
            End If
            If prv_TitleMatch Then Exit Function
        End If
    Next i
End Function

' 목적   : 프로그램 창이 있으면 앞으로 가져오고, 없으면 실행한 뒤 창(과 준비 항목)이 나타날 때까지 대기
' 인수   : strWindowName - 창 이름 (부분 일치 / "=이름" 완전 일치 / "A|B" 여러 이름 중 하나, Excel 자기 창 제외)
'                          예: "WMC Messenger|=Login" → 로그인 창이나 메인 창 중 먼저 뜬 쪽
'          strExePath    - 실행 파일 전체 경로 (창이 없을 때만 실행 / 경로는 DumpWindows 의 '실행 파일 경로' 열 참고)
'          strReadyItem  - 준비 완료로 볼 항목 이름 (예: "Target" - 목록을 다 불러올 때까지 대기 / 생략 시 창만 확인)
'          lngTimeoutSec - 최대 대기 시간(초, 기본 60)
'          strArgs       - 실행 인수 (생략 가능)
' 반환   : Boolean - 창이 준비되었는지 여부 (실패 단계는 직접 실행 창에 출력)
' 예시   : OpenAppWindow("WMC Messenger", "C:\Program Files\WMC\WMC_Msg.exe", "Target")
' 참고   : 최소화된 창은 복원, 트레이로 숨겨진 창(제목 표시줄 있는 창)은 다시 표시.
'          실행 파일은 그 파일이 있는 폴더를 작업 폴더로 실행
Public Function OpenAppWindow(ByVal strWindowName As String, _
                              ByVal strExePath As String, _
                              Optional ByVal strReadyItem As String = "", _
                              Optional ByVal lngTimeoutSec As Long = 60, _
                              Optional ByVal strArgs As String = "") As Boolean
    Const PROC As String = "am_Automation.OpenAppWindow: "

    On Error GoTo ErrHandler

    Dim hWnd       As LongPtr
    Dim blnVisible As Boolean
    Dim objShell   As Object
    Dim objWin     As IUIAutomationElement
    Dim objItem    As IUIAutomationElement
    Dim dtLimit    As Date

    If Len(strWindowName) = 0 Then
        Debug.Print PROC & "창 이름이 비어 있음"
        Exit Function
    End If
    dtLimit = Now + TimeSerial(0, 0, lngTimeoutSec)

    ' ── 1. 이미 있는 창 찾기 (숨겨진 창 포함) ──
    hWnd = prv_FindTopHwnd(strWindowName, blnVisible)

    If hWnd <> 0 Then
        If Not blnVisible Then ShowWindow hWnd, SW_SHOW
    Else
        ' ── 2. 없으면 실행 ──
        If Len(strExePath) = 0 Then
            Debug.Print PROC & "창이 없고 실행 파일 경로가 비어 있음 - " & strWindowName
            Exit Function
        End If
        If Len(Dir(strExePath)) = 0 Then
            Debug.Print PROC & "실행 파일을 찾지 못함 - " & strExePath
            Exit Function
        End If

        Set objShell = CreateObject("WScript.Shell")
        objShell.CurrentDirectory = Left(strExePath, InStrRev(strExePath, "\") - 1)
        objShell.Run Chr(34) & strExePath & Chr(34) & IIf(Len(strArgs) > 0, " " & strArgs, ""), 1, False

        Do
            DoEvents
            Sleep 500
            hWnd = prv_FindTopHwnd(strWindowName, blnVisible)
            If hWnd <> 0 And blnVisible Then Exit Do
            If Now > dtLimit Then
                Debug.Print PROC & "실행했지만 " & lngTimeoutSec & "초 안에 창이 나타나지 않음 - " & strWindowName
                Exit Function
            End If
        Loop
    End If

    prv_ActivateHwnd hWnd

    ' ── 3. 준비 항목 대기 (목록 불러오기 완료 확인) ──
    If Len(strReadyItem) > 0 Then
        Do
            Set objItem = prv_UIA_FindItem(strWindowName, strReadyItem, "", objWin)
            If Not objItem Is Nothing Then Exit Do
            If Now > dtLimit Then
                Debug.Print PROC & "창은 열렸지만 " & lngTimeoutSec & "초 안에 준비 항목이 나타나지 않음 - " & strReadyItem
                Exit Function
            End If
            DoEvents
            Sleep 500
        Loop
    End If

    OpenAppWindow = True
    Exit Function

ErrHandler:
    Debug.Print PROC & "오류 " & Err.Number & ": " & Err.Description
    OpenAppWindow = False
End Function

' 목적   : 제목에 strTitle 이 포함된 최상위 창 핸들 (보이는 창 우선, 없으면 제목 표시줄 있는 숨겨진 창)
'          blnVisible 에 보이는 창인지 돌려줌. Excel 자기 창 제외
Private Function prv_FindTopHwnd(ByVal strTitle As String, ByRef blnVisible As Boolean) As LongPtr
    m_strTopTitle = strTitle
    m_hWndTopVisible = 0
    m_hWndTopHidden = 0
    EnumWindows AddressOf prv_EnumTopProc, 0

    If m_hWndTopVisible <> 0 Then
        prv_FindTopHwnd = m_hWndTopVisible
        blnVisible = True
    Else
        prv_FindTopHwnd = m_hWndTopHidden
        blnVisible = False
    End If
End Function

' 목적   : EnumWindows 콜백 - 제목 부분 일치 창을 보이는 창/숨겨진 창으로 나눠 기록 (보이는 창 찾으면 중단)
Private Function prv_EnumTopProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    Dim strBuffer As String
    Dim lngLen    As Long
    Dim lngPid    As Long

    prv_EnumTopProc = 1

    strBuffer = String(512, vbNullChar)
    lngLen = GetWindowTextA(hWnd, strBuffer, 512)
    If lngLen = 0 Then Exit Function
    If Not prv_TitleMatch(Left(strBuffer, lngLen), m_strTopTitle) Then Exit Function

    GetWindowThreadProcessId hWnd, lngPid
    If lngPid = GetCurrentProcessId() Then Exit Function

    If IsWindowVisible(hWnd) <> 0 Then
        m_hWndTopVisible = hWnd
        prv_EnumTopProc = 0
    ElseIf m_hWndTopHidden = 0 Then
        ' 숨겨진 보조 창(트레이 아이콘용 등)을 잘못 띄우지 않도록 제목 표시줄 있는 창만
        If (GetWindowLongA(hWnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then m_hWndTopHidden = hWnd
    End If
End Function

' 목적   : UI Automation(Early Binding)으로 창 안의 버튼을 이름으로 찾아 Invoke
' 인수   : strWindowName - 대상 창 이름 (UIA Name 속성 기준)
'          strButtonName - 찾을 버튼 이름 (UIA Name 속성 기준)
' 반환   : Boolean - Invoke 성공 여부
' 참고   : Early Binding 사용 - VBE 참조에 "UIAutomationClient" 라이브러리 추가 필수
'          (Late Binding은 이 환경에서 429 오류로 실패 확인 → 프로젝트 Late Binding 원칙의 명시적 예외, TS-V16 참고)
Public Function ClickButtonByUIA(ByVal strWindowName As String, ByVal strButtonName As String) As Boolean
    On Error GoTo ErrHandler

    Dim oAutomation     As New CUIAutomation
    Dim oRoot           As IUIAutomationElement
    Dim oCondition       As IUIAutomationCondition
    Dim oWindowElement   As IUIAutomationElement
    Dim oButtonElement   As IUIAutomationElement
    Dim oInvokePattern   As IUIAutomationInvokePattern

    Set oRoot = oAutomation.GetRootElement

    Set oCondition = oAutomation.CreatePropertyCondition(UIA_NamePropertyId, strWindowName)
    Set oWindowElement = oRoot.FindFirst(TreeScope_Descendants, oCondition)

    If oWindowElement Is Nothing Then
        ClickButtonByUIA = False
        Exit Function
    End If

    Set oCondition = oAutomation.CreatePropertyCondition(UIA_NamePropertyId, strButtonName)
    Set oButtonElement = oWindowElement.FindFirst(TreeScope_Descendants, oCondition)

    If oButtonElement Is Nothing Then
        ClickButtonByUIA = False
        Exit Function
    End If

    Set oInvokePattern = oButtonElement.GetCurrentPattern(UIA_InvokePatternId)
    oInvokePattern.Invoke

    ClickButtonByUIA = True
    Exit Function

ErrHandler:
    ClickButtonByUIA = False
End Function

' 목적   : UI Automation 으로 창 안의 항목(트리·목록·탭 항목 등)을 이름으로 찾아 선택
'          (마우스·키보드를 쓰지 않음 - 화면에 보이지 않는 항목은 스크롤 후 선택)
' 인수   : strWindowName - 대상 창 이름 (부분 일치, Excel 자기 창 제외)
'          strItemName   - 항목 이름 (앞뒤 공백 무시 앞부분 일치 / * ? 포함 시 Like 패턴)
'                          예: "Target" → "Target (52/212)" 일치 (뒤 숫자가 바뀌어도 찾음)
'          blnExpand     - 선택 후 하위 항목 펼치기 (기본 False)
'          strType       - 종류 제한 (예: "TreeItem", "ListItem", "TabItem" / 생략 시 전체)
' 반환   : Boolean - 찾아서 선택했는지 여부
' 예시   : SelectItemByUIA("WMC Messenger", "Target", True, "TreeItem")
' 참고   : 항목 이름·종류는 am_Inspect.DumpElementTree 결과 시트로 확인
Public Function SelectItemByUIA(ByVal strWindowName As String, _
                                ByVal strItemName As String, _
                                Optional ByVal blnExpand As Boolean = False, _
                                Optional ByVal strType As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim objWin    As IUIAutomationElement
    Dim objItem   As IUIAutomationElement
    Dim objSelect As IUIAutomationSelectionItemPattern
    Dim objInvoke As IUIAutomationInvokePattern

    Set objItem = prv_UIA_FindItem(strWindowName, strItemName, strType, objWin)
    If objItem Is Nothing Then Exit Function

    prv_UIA_ScrollIntoView objItem

    Set objSelect = objItem.GetCurrentPattern(UIA_SelectionItemPatternId)
    If Not objSelect Is Nothing Then
        objSelect.Select
    Else
        ' 선택 패턴이 없는 항목(버튼형 탭 등)은 Invoke 로 대체
        Set objInvoke = objItem.GetCurrentPattern(UIA_InvokePatternId)
        If objInvoke Is Nothing Then Exit Function
        objInvoke.Invoke
    End If

    If blnExpand Then prv_UIA_SetExpand objItem, True

    SelectItemByUIA = True
    Exit Function

ErrHandler:
    SelectItemByUIA = False
End Function

' 목적   : UI Automation 으로 창 안의 항목을 이름으로 찾아 펼치거나 접음 (선택 상태는 바꾸지 않음)
' 인수   : strWindowName - 대상 창 이름 (부분 일치, Excel 자기 창 제외)
'          strItemName   - 항목 이름 (SelectItemByUIA 와 같은 일치 규칙)
'          blnExpand     - True=펼치기, False=접기 (기본 True)
'          strType       - 종류 제한 (생략 시 전체)
' 반환   : Boolean - 찾아서 펼치기/접기를 실행했는지 여부
' 예시   : ExpandItemByUIA("WMC Messenger", "Target")
Public Function ExpandItemByUIA(ByVal strWindowName As String, _
                                ByVal strItemName As String, _
                                Optional ByVal blnExpand As Boolean = True, _
                                Optional ByVal strType As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim objWin  As IUIAutomationElement
    Dim objItem As IUIAutomationElement

    Set objItem = prv_UIA_FindItem(strWindowName, strItemName, strType, objWin)
    If objItem Is Nothing Then Exit Function

    ExpandItemByUIA = prv_UIA_SetExpand(objItem, blnExpand)
    Exit Function

ErrHandler:
    ExpandItemByUIA = False
End Function

' 목적   : 항목을 찾아 선택한 뒤 우클릭 메뉴를 열고, 메뉴 항목을 이름으로 찾아 실행
' 인수   : strWindowName - 대상 창 이름 (부분 일치, Excel 자기 창 제외)
'          strItemName   - 우클릭할 항목 이름 (SelectItemByUIA 와 같은 일치 규칙)
'          strMenuName   - 실행할 메뉴 이름 (같은 일치 규칙, 예: "친구추가")
'          strType       - 항목 종류 제한 (예: "TreeItem" / 생략 시 전체)
'          blnAllowMouse - 메시지로 메뉴가 안 열리면 실제 마우스 우클릭 허용 (기본 True, 커서는 원위치)
' 반환   : Boolean - 메뉴 항목을 실행했는지 여부 (실패 단계는 직접 실행 창에 출력)
' 예시   : ClickContextMenuByUIA("WMC Messenger", "Target", "친구추가", "TreeItem")
' 참고   : 메뉴 열기 순서 ① 대상 창 활성화 ② 컨트롤에 WM_CONTEXTMENU 메시지 (마우스 불필요)
'          ③ 안 열리면 항목 화면 위치를 매번 계산해 마우스 우클릭 (blnAllowMouse=True 일 때)
'          메뉴 항목은 UI Automation Invoke 로 실행, 사용 불가(흐린) 항목은 실행하지 않음
Public Function ClickContextMenuByUIA(ByVal strWindowName As String, _
                                      ByVal strItemName As String, _
                                      ByVal strMenuName As String, _
                                      Optional ByVal strType As String = "", _
                                      Optional ByVal blnAllowMouse As Boolean = True) As Boolean
    Const PROC As String = "am_Automation.ClickContextMenuByUIA: "

    On Error GoTo ErrHandler

    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim objWin    As IUIAutomationElement
    Dim objItem   As IUIAutomationElement
    Dim objMenu   As IUIAutomationElement
    Dim objSelect As IUIAutomationSelectionItemPattern
    Dim objInvoke As IUIAutomationInvokePattern
    Dim dicBefore As Object
    Dim hWndMain  As LongPtr
    Dim hWndCtl   As LongPtr
    Dim lngPid    As Long
    Dim lngX      As Long
    Dim lngY      As Long
    Dim udtOld    As POINTAPI
    Dim blnMoved  As Boolean

    ' ── 1. 항목 찾기·선택 ──
    Set objItem = prv_UIA_FindItem(strWindowName, strItemName, strType, objWin)
    If objItem Is Nothing Then
        Debug.Print PROC & "항목을 찾지 못함 - " & strWindowName & " / " & strItemName
        Exit Function
    End If

    prv_UIA_ScrollIntoView objItem
    Set objSelect = objItem.GetCurrentPattern(UIA_SelectionItemPatternId)
    If Not objSelect Is Nothing Then objSelect.Select

    If Not prv_UIA_Center(objItem, lngX, lngY) Then
        Debug.Print PROC & "항목의 화면 위치를 얻지 못함 (화면 밖)"
        Exit Function
    End If

    Set objUIA = New CUIAutomation
    Set objWalker = objUIA.ControlViewWalker
    hWndMain = prv_UIA_Hwnd(objWin)
    hWndCtl = prv_UIA_OwnerHwnd(objWalker, objItem)
    lngPid = objWin.CurrentProcessId

    ' ── 2. 메뉴 열기 - 메시지 방식 먼저 ──
    ' WinForms 우클릭 메뉴는 UIA 에서 메인 창 아래에 붙어 보이므로, 열기 전 메뉴 항목을 기억해 두고
    ' 연 뒤 새로 나타난 항목에서만 찾음 (상단 메뉴 '친구찾기' 등과 이름이 겹쳐도 구분)
    Set dicBefore = prv_UIA_MenuSnapshot(objUIA, objWalker, lngPid)

    prv_ActivateHwnd hWndMain
    If hWndCtl <> 0 Then PostMessageA hWndCtl, WM_CONTEXTMENU, hWndCtl, prv_MakeLParam(lngX, lngY)
    Set objMenu = prv_UIA_WaitMenuItem(objUIA, objWalker, lngPid, strMenuName, dicBefore, 1500)

    If objMenu Is Nothing And blnAllowMouse Then
        Debug.Print PROC & "메시지로 메뉴가 열리지 않아 마우스 우클릭으로 대체"
        GetCursorPos udtOld
        blnMoved = True
        SetCursorPos lngX, lngY
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
        Set objMenu = prv_UIA_WaitMenuItem(objUIA, objWalker, lngPid, strMenuName, dicBefore, 2000)
    End If

    If blnMoved Then SetCursorPos udtOld.x, udtOld.y
    blnMoved = False

    If objMenu Is Nothing Then
        Debug.Print PROC & "메뉴 항목을 찾지 못함 - " & strMenuName & _
                    IIf(blnAllowMouse, "", " (blnAllowMouse=False 라 마우스 우클릭은 시도 안 함)")
        ' 메뉴는 열렸는데 항목만 없는 경우 메뉴 닫기
        If Not prv_UIA_WaitMenuItem(objUIA, objWalker, lngPid, "*", dicBefore, 0) Is Nothing Then prv_SendKey VK_ESCAPE
        Exit Function
    End If

    ' ── 3. 메뉴 항목 실행 ──
    If objMenu.CurrentIsEnabled = 0 Then
        Debug.Print PROC & "메뉴 항목이 사용 불가 상태 - " & strMenuName
        prv_SendKey VK_ESCAPE
        Exit Function
    End If

    Set objInvoke = objMenu.GetCurrentPattern(UIA_InvokePatternId)
    If objInvoke Is Nothing Then
        Debug.Print PROC & "메뉴 항목이 실행(Invoke)을 지원하지 않음 - " & strMenuName
        prv_SendKey VK_ESCAPE
        Exit Function
    End If
    objInvoke.Invoke

    ClickContextMenuByUIA = True
    Exit Function

ErrHandler:
    Debug.Print PROC & "오류 " & Err.Number & ": " & Err.Description
    If blnMoved Then SetCursorPos udtOld.x, udtOld.y
    ClickContextMenuByUIA = False
End Function

' 목적   : 대상 프로세스의 모든 최상위 창에 있는 메뉴 항목 목록 (키=RuntimeId) - 메뉴 열기 전 기준
Private Function prv_UIA_MenuSnapshot(ByVal objUIA As CUIAutomation, _
                                      ByVal objWalker As IUIAutomationTreeWalker, _
                                      ByVal lngPid As Long) As Object
    Dim dicKeys  As Object
    Dim objWin   As IUIAutomationElement
    Dim objNext  As IUIAutomationElement
    Dim lngCount As Long
    Dim lngWinPid As Long

    Set dicKeys = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set objWin = objWalker.GetFirstChildElement(objUIA.GetRootElement)
    Do Until objWin Is Nothing
        lngWinPid = 0
        lngWinPid = objWin.CurrentProcessId
        If lngWinPid = lngPid Then prv_UIA_CollectMenuKeys objWalker, objWin, dicKeys, lngCount

        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objWin)
        Set objWin = objNext
    Loop

    Set prv_UIA_MenuSnapshot = dicKeys
End Function

Private Sub prv_UIA_CollectMenuKeys(ByVal objWalker As IUIAutomationTreeWalker, _
                                    ByVal objEl As IUIAutomationElement, _
                                    ByVal dicKeys As Object, _
                                    ByRef lngCount As Long)
    Dim objChild As IUIAutomationElement
    Dim objNext  As IUIAutomationElement
    Dim lngType  As Long

    On Error Resume Next
    Set objChild = objWalker.GetFirstChildElement(objEl)
    Do Until objChild Is Nothing
        lngCount = lngCount + 1
        If lngCount > UIA_MAX_SEARCH_NODES Then Exit Sub

        lngType = 0
        lngType = objChild.CurrentControlType
        If lngType = UIA_MenuItemControlTypeId Then dicKeys(prv_UIA_Key(objChild)) = True
        prv_UIA_CollectMenuKeys objWalker, objChild, dicKeys, lngCount

        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objChild)
        Set objChild = objNext
    Loop
End Sub

' 목적   : 메뉴를 연 뒤 새로 나타난 메뉴 항목 중 이름이 맞는 첫 항목을 찾을 때까지 대기
'          (lngTimeoutMs=0 이면 한 번만 확인)
Private Function prv_UIA_WaitMenuItem(ByVal objUIA As CUIAutomation, _
                                      ByVal objWalker As IUIAutomationTreeWalker, _
                                      ByVal lngPid As Long, _
                                      ByVal strMenuName As String, _
                                      ByVal dicBefore As Object, _
                                      ByVal lngTimeoutMs As Long) As IUIAutomationElement
    Dim objWin    As IUIAutomationElement
    Dim objNext   As IUIAutomationElement
    Dim objFound  As IUIAutomationElement
    Dim lngCount  As Long
    Dim lngWait   As Long
    Dim lngWinPid As Long

    On Error Resume Next
    Do
        Set objWin = objWalker.GetFirstChildElement(objUIA.GetRootElement)
        Do Until objWin Is Nothing
            lngWinPid = 0
            lngWinPid = objWin.CurrentProcessId
            If lngWinPid = lngPid Then
                lngCount = 0
                Set objFound = Nothing
                Set objFound = prv_UIA_FindNewMenu(objWalker, objWin, strMenuName, dicBefore, lngCount)
                If Not objFound Is Nothing Then
                    Set prv_UIA_WaitMenuItem = objFound
                    Exit Function
                End If
            End If
            Set objNext = Nothing
            Set objNext = objWalker.GetNextSiblingElement(objWin)
            Set objWin = objNext
        Loop

        If lngWait >= lngTimeoutMs Then Exit Do
        DoEvents
        Sleep 100
        lngWait = lngWait + 100
    Loop
End Function

Private Function prv_UIA_FindNewMenu(ByVal objWalker As IUIAutomationTreeWalker, _
                                     ByVal objEl As IUIAutomationElement, _
                                     ByVal strMenuName As String, _
                                     ByVal dicBefore As Object, _
                                     ByRef lngCount As Long) As IUIAutomationElement
    Dim objChild As IUIAutomationElement
    Dim objNext  As IUIAutomationElement
    Dim objFound As IUIAutomationElement
    Dim lngType  As Long
    Dim strName  As String

    On Error Resume Next
    Set objChild = objWalker.GetFirstChildElement(objEl)
    Do Until objChild Is Nothing
        lngCount = lngCount + 1
        If lngCount > UIA_MAX_SEARCH_NODES Then Exit Function

        lngType = 0
        strName = ""
        lngType = objChild.CurrentControlType
        strName = objChild.CurrentName

        If lngType = UIA_MenuItemControlTypeId And prv_UIA_NameMatch(strName, strMenuName) Then
            If Not dicBefore.Exists(prv_UIA_Key(objChild)) Then
                Set prv_UIA_FindNewMenu = objChild
                Exit Function
            End If
        End If

        Set objFound = Nothing
        Set objFound = prv_UIA_FindNewMenu(objWalker, objChild, strMenuName, dicBefore, lngCount)
        If Not objFound Is Nothing Then
            Set prv_UIA_FindNewMenu = objFound
            Exit Function
        End If

        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objChild)
        Set objChild = objNext
    Loop
End Function

' 목적   : 요소 식별 키 (UIA RuntimeId 를 점으로 연결)
Private Function prv_UIA_Key(ByVal objEl As IUIAutomationElement) As String
    Dim vntId As Variant
    Dim i     As Long
    Dim strKey As String

    On Error Resume Next
    vntId = objEl.GetRuntimeId
    If Not IsArray(vntId) Then Exit Function
    For i = LBound(vntId) To UBound(vntId)
        strKey = strKey & "." & vntId(i)
    Next i
    prv_UIA_Key = strKey
End Function

' 목적   : 요소의 Win32 창 핸들 (없으면 0)
Private Function prv_UIA_Hwnd(ByVal objEl As IUIAutomationElement) As LongPtr
    On Error Resume Next
    prv_UIA_Hwnd = CLng(objEl.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId))
End Function

' 목적   : 요소 자신 또는 가장 가까운 상위 요소 중 Win32 창 핸들이 있는 것의 핸들 (트리 항목 → 트리 컨트롤)
Private Function prv_UIA_OwnerHwnd(ByVal objWalker As IUIAutomationTreeWalker, _
                                   ByVal objEl As IUIAutomationElement) As LongPtr
    Dim objNext As IUIAutomationElement
    Dim hWnd    As LongPtr
    Dim i       As Long

    On Error Resume Next
    For i = 1 To 50
        If objEl Is Nothing Then Exit Function
        hWnd = prv_UIA_Hwnd(objEl)
        If hWnd <> 0 Then
            prv_UIA_OwnerHwnd = hWnd
            Exit Function
        End If
        Set objNext = Nothing
        Set objNext = objWalker.GetParentElement(objEl)
        Set objEl = objNext
    Next i
End Function

' 목적   : 요소 화면 위치의 중심 좌표 (크기 0·화면 밖이면 False)
Private Function prv_UIA_Center(ByVal objEl As IUIAutomationElement, _
                                ByRef lngX As Long, _
                                ByRef lngY As Long) As Boolean
    Dim vntRect As Variant
    Dim lngBase As Long

    On Error GoTo ErrHandler

    vntRect = objEl.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
    If Not IsArray(vntRect) Then Exit Function
    If UBound(vntRect) - LBound(vntRect) < 3 Then Exit Function

    lngBase = LBound(vntRect)
    If vntRect(lngBase + 2) <= 0 Or vntRect(lngBase + 3) <= 0 Then Exit Function

    lngX = CLng(vntRect(lngBase) + vntRect(lngBase + 2) / 2)
    lngY = CLng(vntRect(lngBase + 1) + vntRect(lngBase + 3) / 2)
    prv_UIA_Center = True
    Exit Function

ErrHandler:
    prv_UIA_Center = False
End Function

' 목적   : 화면 좌표 → WM_CONTEXTMENU 의 lParam (하위 16bit x, 상위 16bit y, 음수 좌표 포함)
Private Function prv_MakeLParam(ByVal lngX As Long, ByVal lngY As Long) As LongPtr
    Dim dblValue As Double

    dblValue = (lngY And &HFFFF&) * 65536# + (lngX And &HFFFF&)
    If dblValue >= 2147483648# Then dblValue = dblValue - 4294967296#
    prv_MakeLParam = CLng(dblValue)
End Function

' 목적   : 창을 이름(부분 일치)으로 찾은 뒤 그 안에서 항목을 이름·종류로 찾음 (깊이 우선, 첫 일치)
'          objWin 에 찾은 최상위 창을 돌려줌
Private Function prv_UIA_FindItem(ByVal strWindowName As String, _
                                  ByVal strItemName As String, _
                                  ByVal strType As String, _
                                  ByRef objWin As IUIAutomationElement) As IUIAutomationElement
    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim lngTypeId As Long
    Dim lngCount  As Long

    If Len(strWindowName) = 0 Or Len(strItemName) = 0 Then Exit Function

    lngTypeId = prv_UIA_TypeId(strType)
    If lngTypeId < 0 Then Exit Function    ' 알 수 없는 종류 이름

    Set objUIA = New CUIAutomation
    Set objWalker = objUIA.ControlViewWalker

    Set objWin = prv_UIA_FindWindow(objUIA, objWalker, strWindowName)
    If objWin Is Nothing Then Exit Function

    Set prv_UIA_FindItem = prv_UIA_Search(objWalker, objWin, strItemName, lngTypeId, lngCount)
End Function

' 목적   : 최상위 창 중 이름에 strWindowName 이 포함된 첫 창 (Excel 자기 창 제외 - 자기 UI 조회 시 멈춤 방지)
Private Function prv_UIA_FindWindow(ByVal objUIA As CUIAutomation, _
                                    ByVal objWalker As IUIAutomationTreeWalker, _
                                    ByVal strWindowName As String) As IUIAutomationElement
    Dim objWin  As IUIAutomationElement
    Dim objNext As IUIAutomationElement
    Dim lngSelf As Long
    Dim lngPid  As Long
    Dim strName As String

    lngSelf = GetCurrentProcessId()

    ' 속성은 변수로 먼저 읽음 - Resume Next 상태에서 If 조건식 오류 시 블록 안으로 진입하는 것 방지
    On Error Resume Next
    Set objWin = objWalker.GetFirstChildElement(objUIA.GetRootElement)
    Do Until objWin Is Nothing
        lngPid = lngSelf
        strName = ""
        lngPid = objWin.CurrentProcessId
        strName = objWin.CurrentName

        If lngPid <> lngSelf And prv_TitleMatch(strName, strWindowName) Then
            Set prv_UIA_FindWindow = objWin
            Exit Function
        End If
        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objWin)
        Set objWin = objNext
    Loop
End Function

' 목적   : objEl 하위에서 이름·종류가 맞는 첫 요소 (재귀, lngCount 로 탐색 상한 관리)
Private Function prv_UIA_Search(ByVal objWalker As IUIAutomationTreeWalker, _
                                ByVal objEl As IUIAutomationElement, _
                                ByVal strItemName As String, _
                                ByVal lngTypeId As Long, _
                                ByRef lngCount As Long) As IUIAutomationElement
    Dim objChild As IUIAutomationElement
    Dim objNext  As IUIAutomationElement
    Dim objFound As IUIAutomationElement
    Dim lngType  As Long
    Dim strName  As String
    Dim blnById  As Boolean

    ' "#아이디" 는 AutomationId 완전 일치(대소문자 무시)로 찾음 - 이름이 비어 있는 입력칸·버튼용
    blnById = (Left(strItemName, 1) = "#")

    On Error Resume Next
    Set objChild = objWalker.GetFirstChildElement(objEl)
    Do Until objChild Is Nothing
        lngCount = lngCount + 1
        If lngCount > UIA_MAX_SEARCH_NODES Then Exit Function

        lngType = -1
        strName = ""
        lngType = objChild.CurrentControlType
        If blnById Then
            strName = objChild.CurrentAutomationId
        Else
            strName = objChild.CurrentName
        End If

        If (lngTypeId = 0 Or lngType = lngTypeId) And _
           IIf(blnById, StrComp(strName, Mid(strItemName, 2), vbTextCompare) = 0, prv_UIA_NameMatch(strName, strItemName)) Then
            Set prv_UIA_Search = objChild
            Exit Function
        End If

        Set objFound = Nothing
        Set objFound = prv_UIA_Search(objWalker, objChild, strItemName, lngTypeId, lngCount)
        If Not objFound Is Nothing Then
            Set prv_UIA_Search = objFound
            Exit Function
        End If

        Set objNext = Nothing
        Set objNext = objWalker.GetNextSiblingElement(objChild)
        Set objChild = objNext
    Loop
End Function

' 목적   : 이름 일치 규칙 - 앞뒤 공백 무시, * ? 포함 시 Like 패턴, 그 외 앞부분 일치 (대소문자 무시)
Private Function prv_UIA_NameMatch(ByVal strName As String, ByVal strPattern As String) As Boolean
    On Error Resume Next    ' 잘못된 Like 패턴 등 오류 시 불일치(False)
    strName = Trim(strName)
    strPattern = Trim(strPattern)
    If Len(strName) = 0 Then Exit Function

    If InStr(strPattern, "*") > 0 Or InStr(strPattern, "?") > 0 Then
        prv_UIA_NameMatch = LCase(strName) Like LCase(strPattern)
    Else
        prv_UIA_NameMatch = (StrComp(Left(strName, Len(strPattern)), strPattern, vbTextCompare) = 0)
    End If
End Function

' 목적   : 종류 이름 → ControlType ID (""=제한 없음 0, 알 수 없는 이름 -1)
Private Function prv_UIA_TypeId(ByVal strType As String) As Long
    Dim arrNames As Variant
    Dim i        As Long

    If Len(strType) = 0 Then Exit Function

    arrNames = Split("Button,Calendar,CheckBox,ComboBox,Edit,Hyperlink,Image,ListItem,List,Menu," & _
                     "MenuBar,MenuItem,ProgressBar,RadioButton,ScrollBar,Slider,Spinner,StatusBar,Tab,TabItem," & _
                     "Text,ToolBar,ToolTip,Tree,TreeItem,Custom,Group,Thumb,DataGrid,DataItem," & _
                     "Document,SplitButton,Window,Pane,Header,HeaderItem,Table,TitleBar,Separator,SemanticZoom," & _
                     "AppBar", ",")

    For i = LBound(arrNames) To UBound(arrNames)
        If StrComp(arrNames(i), strType, vbTextCompare) = 0 Then
            prv_UIA_TypeId = 50000 + i
            Exit Function
        End If
    Next i

    prv_UIA_TypeId = -1
End Function

' 목적   : 펼치기/접기 실행 (패턴 미지원 시 False)
Private Function prv_UIA_SetExpand(ByVal objItem As IUIAutomationElement, ByVal blnExpand As Boolean) As Boolean
    Dim objExpand As IUIAutomationExpandCollapsePattern

    On Error GoTo ErrHandler

    Set objExpand = objItem.GetCurrentPattern(UIA_ExpandCollapsePatternId)
    If objExpand Is Nothing Then Exit Function

    If blnExpand Then
        objExpand.Expand
    Else
        objExpand.Collapse
    End If
    prv_UIA_SetExpand = True
    Exit Function

ErrHandler:
    prv_UIA_SetExpand = False
End Function

' 목적   : 항목이 화면 밖이면 보이도록 스크롤 (지원하지 않으면 무시)
Private Sub prv_UIA_ScrollIntoView(ByVal objItem As IUIAutomationElement)
    Dim objScroll As IUIAutomationScrollItemPattern

    On Error Resume Next
    Set objScroll = objItem.GetCurrentPattern(UIA_ScrollItemPatternId)
    If Not objScroll Is Nothing Then objScroll.ScrollIntoView
End Sub

' 목적   : EnumChildWindows 콜백 - m_strTargetText 를 포함하는 첫 자식 창을 m_hWndFound 에 기록
Private Function prv_EnumChildProc(ByVal hWndChild As LongPtr, ByVal lParam As LongPtr) As Long
    Dim strBuffer As String
    Dim lngLen    As Long

    strBuffer = String(256, vbNullChar)
    lngLen = GetWindowTextA(hWndChild, strBuffer, 256)
    strBuffer = Left(strBuffer, lngLen)

    If InStr(strBuffer, m_strTargetText) > 0 Then
        m_hWndFound = hWndChild
        prv_EnumChildProc = 0
    Else
        prv_EnumChildProc = 1
    End If
End Function
