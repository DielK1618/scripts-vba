Attribute VB_Name = "am_Inspect"
Option Explicit
' 버전 : 2026-10-16 16:32 · 신규 - 외부 프로그램 창·요소 인식

' ┌─────────────────────────────────────────────────────────┐
' │  am_Inspect                                             │
' │  역할 : 외부 프로그램 UI 요소 인식 (UI Automation 조회 전용) │
' └─────────────────────────────────────────────────────────┘
' ※ 이 모듈은 요소를 "찾고 읽기"만 한다. 클릭·입력 등 제어는 am_Automation 담당.
' ※ 참조 필요: VBE 도구 > 참조에서 "UIAutomationClient" 체크 (Early Binding)
'   CreateObject Late Binding 은 이 환경에서 실패 확인 → 프로젝트 Late Binding 원칙의 명시적 예외 (TS-V16)
'
' 사용 순서 (CWB 에서 ref_Inspect 래퍼로 호출)
'   1. DumpWindows                 → 열린 창 목록 시트 생성 (대상 창 이름·프로세스 확인)
'   2. DumpElementTree "창 이름"   → 그 창 안의 요소 트리 시트 생성
'   3. DumpElementAtCursor 3       → 3초 안에 마우스를 대상 요소 위에 올려 두면 그 요소와 상위 경로 기록
'
' 제한
'   - Excel 자기 창은 조회하지 않음 (실행 중인 자기 UI 를 UI Automation 으로 조회하면 멈출 수 있음)
'     → 창 찾기에서 제외, 커서가 Excel 위에 있으면 조회 중단
'
' 결과 해석
'   - 요소가 이름·종류와 함께 자세히 나옴        → am_Automation 의 UIA 계열로 직접 제어 가능
'   - Chrome_WidgetWin_1 아래 Document 가 비어 있음 → Electron/웹 기반, 접근성 비활성 상태
'   - Pane/Custom 몇 개만 나오고 하위가 없음      → 직접 인식 불가, 단축키·좌표 방식으로 대체

' ── Windows API ───────────────────────────────────────────────
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
#If Win64 Then
' 64bit: POINT 구조체를 값으로 넘기므로 x(하위 32bit)·y(상위 32bit)를 LongLong 하나로 묶어 전달
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal llPoint As LongLong) As LongPtr
#Else
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type

' ── UI Automation 속성 ID ────────────────────────────────────
Private Const PID_RECT            As Long = 30001   ' BoundingRectangle
Private Const PID_PROCESSID       As Long = 30002
Private Const PID_CONTROLTYPE     As Long = 30003
Private Const PID_NAME            As Long = 30005
Private Const PID_ISENABLED       As Long = 30010
Private Const PID_AUTOMATIONID    As Long = 30011
Private Const PID_CLASSNAME       As Long = 30012
Private Const PID_ISPASSWORD      As Long = 30019
Private Const PID_HWND            As Long = 30020   ' NativeWindowHandle
Private Const PID_FRAMEWORKID     As Long = 30024
Private Const PID_VALUE           As Long = 30045   ' ValueValue

' 패턴 지원 여부 속성 ID (Is...PatternAvailable)
Private Const PID_HAS_EXPAND      As Long = 30028
Private Const PID_HAS_INVOKE      As Long = 30031
Private Const PID_HAS_SCROLL      As Long = 30034
Private Const PID_HAS_SELECTITEM  As Long = 30036
Private Const PID_HAS_TEXT        As Long = 30040
Private Const PID_HAS_TOGGLE      As Long = 30041
Private Const PID_HAS_VALUE       As Long = 30043
Private Const PID_HAS_WINDOW      As Long = 30044

Private Const MAX_TEXT_LEN        As Long = 200
Private Const MAX_INDENT          As Long = 15     ' Excel IndentLevel 최대값
Private Const MAX_HIT_NODES       As Long = 5000   ' 커서 위치 요소 탐색 상한

' ══════════════════════════════════════════════════════════
'  창 목록
' ══════════════════════════════════════════════════════════

' 목적   : 바탕화면의 최상위 창 목록 반환
' 인수   : strKeyword - 창 이름 또는 프로세스 이름에 포함된 문자열로 필터 (생략 시 전체)
' 반환   : Variant - 2차원 배열 (1행 머리글) / 실패 시 Empty
'          열: 이름, 프로세스, 프로세스ID, 클래스, 프레임워크, 창 핸들, 화면 위치, 실행 파일 경로
' 예시   : arr = ListWindows("메신저")
Public Function ListWindows(Optional ByVal strKeyword As String = "") As Variant
    Dim strErr As String
    ListWindows = prv_ListWindows(strKeyword, strErr)
End Function

' 목적   : 최상위 창 목록을 새 시트에 기록
' 인수   : strKeyword - 창 이름 또는 프로세스 이름 필터 (생략 시 전체)
'          wb         - 결과 시트를 추가할 통합문서 (ref_ 래퍼가 ThisWorkbook 자동 입력)
' 반환   : Long - 기록한 창 수 (실패 시 0)
Public Function DumpWindows(Optional ByVal strKeyword As String = "", _
                            Optional ByVal wb As Workbook = Nothing) As Long

    On Error GoTo ErrHandler

    Dim arrData As Variant
    Dim strErr  As String

    If wb Is Nothing Then
        MsgBox "결과를 기록할 통합문서가 지정되지 않았습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    arrData = prv_ListWindows(strKeyword, strErr)
    If Not IsArray(arrData) Then
        MsgBox strErr, vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    prv_WriteSheet wb, "UI_창목록", _
                   "열린 창 목록" & IIf(strKeyword = "", "", " · 필터: " & strKeyword), _
                   arrData, 0, 0
    DumpWindows = UBound(arrData, 1) - 1

    Exit Function

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME

End Function

' ══════════════════════════════════════════════════════════
'  요소 트리
' ══════════════════════════════════════════════════════════

' 목적   : 지정한 창 안의 UI 요소를 트리 순서로 반환
' 인수   : strWindowKeyword - 창 이름에 포함된 문자열 (없으면 프로세스 이름으로 재탐색)
'          lngMaxDepth      - 내려갈 최대 깊이 (기본 10)
'          lngMaxRows       - 최대 요소 수 (기본 3000, 대형 앱 멈춤 방지)
'          blnRawView       - True=모든 요소 / False=컨트롤 요소만 (기본 True)
' 반환   : Variant - 2차원 배열 (1행 머리글) / 실패 시 Empty
'          열: 깊이, 위치번호, 이름, 종류, AutomationId, 클래스, 프레임워크,
'              지원 패턴, 값, 사용 가능, 화면 위치, 창 핸들
' 예시   : arr = GetElementTree("메신저", 8)
Public Function GetElementTree(ByVal strWindowKeyword As String, _
                               Optional ByVal lngMaxDepth As Long = 10, _
                               Optional ByVal lngMaxRows As Long = 3000, _
                               Optional ByVal blnRawView As Boolean = True) As Variant
    Dim strWinName As String
    Dim strErr     As String
    GetElementTree = prv_GetElementTree(strWindowKeyword, lngMaxDepth, lngMaxRows, blnRawView, _
                                        strWinName, strErr)
End Function

' 목적   : 지정한 창 안의 UI 요소 트리를 새 시트에 기록 (이름 열은 깊이만큼 들여쓰기)
' 인수   : strWindowKeyword - 창 이름에 포함된 문자열
'          lngMaxDepth      - 최대 깊이 (기본 10)
'          lngMaxRows       - 최대 요소 수 (기본 3000)
'          blnRawView       - True=모든 요소 / False=컨트롤 요소만 (기본 True)
'          wb               - 결과 시트를 추가할 통합문서 (ref_ 래퍼가 ThisWorkbook 자동 입력)
' 반환   : Long - 기록한 요소 수 (실패 시 0)
Public Function DumpElementTree(ByVal strWindowKeyword As String, _
                                Optional ByVal lngMaxDepth As Long = 10, _
                                Optional ByVal lngMaxRows As Long = 3000, _
                                Optional ByVal blnRawView As Boolean = True, _
                                Optional ByVal wb As Workbook = Nothing) As Long

    On Error GoTo ErrHandler

    Dim arrData    As Variant
    Dim strWinName As String
    Dim strErr     As String

    If wb Is Nothing Then
        MsgBox "결과를 기록할 통합문서가 지정되지 않았습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    arrData = prv_GetElementTree(strWindowKeyword, lngMaxDepth, lngMaxRows, blnRawView, _
                                 strWinName, strErr)
    If Not IsArray(arrData) Then
        MsgBox strErr, vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    prv_WriteSheet wb, "UI_트리", _
                   "UI 요소 트리 · 창: " & strWinName & " · 최대 깊이 " & lngMaxDepth & _
                   IIf(blnRawView, " · 전체 요소", " · 컨트롤 요소만"), _
                   arrData, 3, 1
    DumpElementTree = UBound(arrData, 1) - 1

    If DumpElementTree >= lngMaxRows Then
        MsgBox "요소가 많아 " & lngMaxRows & "개까지만 기록했습니다." & vbCrLf & _
               "lngMaxDepth 를 줄이거나 lngMaxRows 를 늘려 다시 실행하세요.", _
               vbInformation, am_Core.AM_NAME
    End If

    Exit Function

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME

End Function

' ══════════════════════════════════════════════════════════
'  마우스 위치 요소
' ══════════════════════════════════════════════════════════

' 목적   : 대기 후 마우스 커서 아래 요소와 그 상위 경로(최상위 창까지)를 반환
' 인수   : lngDelaySec - 커서를 대상 위에 올려 둘 대기 시간(초, 기본 3)
'          blnRawView  - True=모든 요소 기준 경로 / False=컨트롤 요소 기준 (기본 True)
' 반환   : Variant - 2차원 배열 (1행 머리글, 2행=최상위 창 … 마지막 행=커서 아래 요소) / 실패 시 Empty
'          열: GetElementTree 와 동일 (위치번호는 비움)
' 예시   : arr = GetElementAtCursor(5)
Public Function GetElementAtCursor(Optional ByVal lngDelaySec As Long = 3, _
                                   Optional ByVal blnRawView As Boolean = True) As Variant
    Dim strErr As String
    GetElementAtCursor = prv_GetElementAtCursor(lngDelaySec, blnRawView, strErr)
End Function

' 목적   : 대기 후 마우스 커서 아래 요소와 상위 경로를 새 시트에 기록
' 인수   : lngDelaySec - 대기 시간(초, 기본 3)
'          blnRawView  - True=모든 요소 기준 / False=컨트롤 요소 기준 (기본 True)
'          wb          - 결과 시트를 추가할 통합문서 (ref_ 래퍼가 ThisWorkbook 자동 입력)
' 반환   : Long - 기록한 행 수 (실패 시 0)
Public Function DumpElementAtCursor(Optional ByVal lngDelaySec As Long = 3, _
                                    Optional ByVal blnRawView As Boolean = True, _
                                    Optional ByVal wb As Workbook = Nothing) As Long

    On Error GoTo ErrHandler

    Dim arrData As Variant
    Dim strErr  As String

    If wb Is Nothing Then
        MsgBox "결과를 기록할 통합문서가 지정되지 않았습니다.", vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    arrData = prv_GetElementAtCursor(lngDelaySec, blnRawView, strErr)
    If Not IsArray(arrData) Then
        MsgBox strErr, vbExclamation, am_Core.AM_NAME
        Exit Function
    End If

    prv_WriteSheet wb, "UI_커서", _
                   "마우스 위치 요소 · 마지막 행이 커서 아래 요소, 위쪽은 상위 경로", _
                   arrData, 3, 1
    DumpElementAtCursor = UBound(arrData, 1) - 1

    Exit Function

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME

End Function

' ══════════════════════════════════════════════════════════
'  내부 - 조회
' ══════════════════════════════════════════════════════════

Private Function prv_ListWindows(ByVal strKeyword As String, ByRef strErr As String) As Variant

    On Error GoTo ErrHandler

    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim objWin    As IUIAutomationElement
    Dim dicProc   As Object
    Dim colRows   As Collection
    Dim strName   As String
    Dim strProc   As String
    Dim lngPid    As Long

    Set objUIA = New CUIAutomation
    Set objWalker = objUIA.ControlViewWalker
    Set dicProc = prv_ProcessNames()
    Set colRows = New Collection

    Set objWin = prv_FirstChild(objWalker, objUIA.GetRootElement)
    Do Until objWin Is Nothing
        strName = prv_Text(objWin, PID_NAME)
        lngPid = prv_Long(objWin, PID_PROCESSID)
        strProc = prv_ProcName(dicProc, lngPid)

        If strKeyword = "" _
           Or InStr(1, strName, strKeyword, vbTextCompare) > 0 _
           Or InStr(1, strProc, strKeyword, vbTextCompare) > 0 Then
            colRows.Add Array(prv_Cell(strName), _
                              prv_Cell(strProc), _
                              lngPid, _
                              prv_Cell(prv_Text(objWin, PID_CLASSNAME)), _
                              prv_Cell(prv_Text(objWin, PID_FRAMEWORKID)), _
                              prv_Handle(objWin), _
                              prv_Cell(prv_Rect(objWin)), _
                              prv_Cell(prv_ProcPath(dicProc, lngPid)))
        End If

        Set objWin = prv_NextSibling(objWalker, objWin)
    Loop

    If colRows.Count = 0 Then
        strErr = "조건에 맞는 창이 없습니다" & IIf(strKeyword = "", ".", ": " & strKeyword)
        Exit Function
    End If

    prv_ListWindows = prv_ToArray(Array("이름", "프로세스", "프로세스ID", "클래스", _
                                        "프레임워크", "창 핸들", "화면 위치(좌,상,폭,높이)", _
                                        "실행 파일 경로"), colRows)
    Exit Function

ErrHandler:
    strErr = "창 목록 조회 중 오류 " & Err.Number & ": " & Err.Description

End Function

Private Function prv_GetElementTree(ByVal strWindowKeyword As String, _
                                    ByVal lngMaxDepth As Long, _
                                    ByVal lngMaxRows As Long, _
                                    ByVal blnRawView As Boolean, _
                                    ByRef strWinName As String, _
                                    ByRef strErr As String) As Variant

    On Error GoTo ErrHandler

    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim objWin    As IUIAutomationElement
    Dim colRows   As Collection

    If Len(strWindowKeyword) = 0 Then
        strErr = "찾을 창 이름을 입력하세요."
        Exit Function
    End If

    Set objUIA = New CUIAutomation
    If blnRawView Then
        Set objWalker = objUIA.RawViewWalker
    Else
        Set objWalker = objUIA.ControlViewWalker
    End If

    Set objWin = prv_FindWindow(objUIA, strWindowKeyword)
    If objWin Is Nothing Then
        strErr = "창을 찾지 못했습니다: " & strWindowKeyword & vbCrLf & _
                 "DumpWindows 로 정확한 창 이름을 먼저 확인하세요."
        Exit Function
    End If
    strWinName = prv_Text(objWin, PID_NAME)

    Set colRows = New Collection
    prv_Walk objWalker, objWin, 0, "1", lngMaxDepth, lngMaxRows, colRows

    prv_GetElementTree = prv_ToArray(prv_ElementHeaders(), colRows)
    Exit Function

ErrHandler:
    strErr = "요소 트리 조회 중 오류 " & Err.Number & ": " & Err.Description

End Function

Private Function prv_GetElementAtCursor(ByVal lngDelaySec As Long, _
                                        ByVal blnRawView As Boolean, _
                                        ByRef strErr As String) As Variant

    On Error GoTo ErrHandler

    Dim objUIA    As CUIAutomation
    Dim objWalker As IUIAutomationTreeWalker
    Dim objRoot   As IUIAutomationElement
    Dim objEl     As IUIAutomationElement
    Dim colChain  As Collection
    Dim colRows   As Collection
    Dim udtPt     As POINTAPI
    Dim hWnd      As LongPtr
    Dim lngPid    As Long
    Dim i         As Long

    If lngDelaySec > 0 Then Application.Wait Now + TimeSerial(0, 0, lngDelaySec)

    ' ElementFromPoint 는 구조체를 값으로 받아 VBA 에서 호출 불가(컴파일 오류)
    ' → Win32 로 커서 아래 창을 구한 뒤 그 안에서 화면 위치로 가장 안쪽 요소를 찾음
    GetCursorPos udtPt
    hWnd = prv_WindowAtPoint(udtPt.x, udtPt.y)
    If hWnd = 0 Then
        strErr = "커서 위치(" & udtPt.x & ", " & udtPt.y & ")에서 창을 찾지 못했습니다."
        Exit Function
    End If

    GetWindowThreadProcessId hWnd, lngPid
    If lngPid = GetCurrentProcessId() Then
        strErr = "커서가 Excel 창 위에 있습니다." & vbCrLf & _
                 "대기 시간 안에 마우스를 조회할 프로그램의 요소 위로 옮기세요."
        Exit Function
    End If

    Set objUIA = New CUIAutomation
    If blnRawView Then
        Set objWalker = objUIA.RawViewWalker
    Else
        Set objWalker = objUIA.ControlViewWalker
    End If
    Set objRoot = objUIA.GetRootElement

    ' 형식 라이브러리 인수가 void* 라서 ByVal 을 명시해야 핸들 값이 그대로 전달됨
    Set objEl = objUIA.ElementFromHandle(ByVal hWnd)
    If objEl Is Nothing Then
        strErr = "커서 위치(" & udtPt.x & ", " & udtPt.y & ")에서 요소를 찾지 못했습니다."
        Exit Function
    End If
    Set objEl = prv_HitTest(objWalker, objEl, udtPt.x, udtPt.y)

    ' 커서 아래 요소 → 바탕화면 직전(최상위 창)까지 거슬러 올라가며 앞쪽에 쌓음
    Set colChain = New Collection
    Do Until objEl Is Nothing
        If objUIA.CompareElements(objEl, objRoot) <> 0 Then Exit Do
        If colChain.Count = 0 Then
            colChain.Add objEl
        Else
            colChain.Add objEl, Before:=1
        End If
        Set objEl = prv_Parent(objWalker, objEl)
    Loop

    Set colRows = New Collection
    For i = 1 To colChain.Count
        colRows.Add prv_ElementRow(colChain(i), i - 1, "")
    Next i

    prv_GetElementAtCursor = prv_ToArray(prv_ElementHeaders(), colRows)
    Exit Function

ErrHandler:
    strErr = "커서 위치 요소 조회 중 오류 " & Err.Number & ": " & Err.Description

End Function

' 목적   : 요소와 하위 요소를 깊이 우선으로 순회하며 행 추가
'          strPath 는 형제 순번을 점으로 이은 위치번호 (예: 1.3.2)
Private Sub prv_Walk(ByVal objWalker As IUIAutomationTreeWalker, _
                     ByVal objEl As IUIAutomationElement, _
                     ByVal lngDepth As Long, _
                     ByVal strPath As String, _
                     ByVal lngMaxDepth As Long, _
                     ByVal lngMaxRows As Long, _
                     ByVal colRows As Collection)

    Dim objChild As IUIAutomationElement
    Dim lngIndex As Long

    If colRows.Count >= lngMaxRows Then Exit Sub
    colRows.Add prv_ElementRow(objEl, lngDepth, strPath)
    If lngDepth >= lngMaxDepth Then Exit Sub

    Set objChild = prv_FirstChild(objWalker, objEl)
    Do Until objChild Is Nothing
        lngIndex = lngIndex + 1
        prv_Walk objWalker, objChild, lngDepth + 1, strPath & "." & lngIndex, _
                 lngMaxDepth, lngMaxRows, colRows
        If colRows.Count >= lngMaxRows Then Exit Do
        Set objChild = prv_NextSibling(objWalker, objChild)
    Loop

End Sub

' 목적   : 최상위 창 중 이름에 키워드가 포함된 첫 창 반환 (Excel 자기 창 제외)
'          이름으로 못 찾으면 프로세스 이름으로 재탐색 (예: "nateon")
Private Function prv_FindWindow(ByVal objUIA As CUIAutomation, _
                                ByVal strKeyword As String) As IUIAutomationElement

    Dim objWalker As IUIAutomationTreeWalker
    Dim objWin    As IUIAutomationElement
    Dim dicProc   As Object
    Dim intPass   As Integer
    Dim strTarget As String
    Dim lngSelf   As Long
    Dim lngPid    As Long

    Set objWalker = objUIA.ControlViewWalker
    lngSelf = GetCurrentProcessId()

    For intPass = 1 To 2
        If intPass = 2 Then Set dicProc = prv_ProcessNames()

        Set objWin = prv_FirstChild(objWalker, objUIA.GetRootElement)
        Do Until objWin Is Nothing
            lngPid = prv_Long(objWin, PID_PROCESSID)
            If intPass = 1 Then
                strTarget = prv_Text(objWin, PID_NAME)
            Else
                strTarget = prv_ProcName(dicProc, lngPid)
            End If

            If lngPid <> lngSelf And InStr(1, strTarget, strKeyword, vbTextCompare) > 0 Then
                Set prv_FindWindow = objWin
                Exit Function
            End If

            Set objWin = prv_NextSibling(objWalker, objWin)
        Loop
    Next intPass

End Function

' ══════════════════════════════════════════════════════════
'  내부 - 요소 정보
' ══════════════════════════════════════════════════════════

Private Function prv_ElementHeaders() As Variant
    prv_ElementHeaders = Array("깊이", "위치번호", "이름", "종류", "AutomationId", "클래스", _
                               "프레임워크", "지원 패턴", "값", "사용 가능", _
                               "화면 위치(좌,상,폭,높이)", "창 핸들")
End Function

' 목적   : 요소 1개를 결과 행(1차원 배열)으로 변환 - 열 순서는 prv_ElementHeaders 와 동일
Private Function prv_ElementRow(ByVal objEl As IUIAutomationElement, _
                                ByVal lngDepth As Long, _
                                ByVal strPath As String) As Variant
    Dim strValue As String

    ' 값은 Value 패턴 지원 + 비밀번호 칸이 아닐 때만 읽음
    If prv_Bool(objEl, PID_HAS_VALUE) And Not prv_Bool(objEl, PID_ISPASSWORD) Then
        strValue = prv_Text(objEl, PID_VALUE)
    End If

    prv_ElementRow = Array(lngDepth, _
                           prv_Cell(strPath), _
                           prv_Cell(prv_Text(objEl, PID_NAME)), _
                           prv_Cell(prv_TypeName(prv_Long(objEl, PID_CONTROLTYPE))), _
                           prv_Cell(prv_Text(objEl, PID_AUTOMATIONID)), _
                           prv_Cell(prv_Text(objEl, PID_CLASSNAME)), _
                           prv_Cell(prv_Text(objEl, PID_FRAMEWORKID)), _
                           prv_Cell(prv_Patterns(objEl)), _
                           prv_Cell(strValue), _
                           IIf(prv_Bool(objEl, PID_ISENABLED), "O", "X"), _
                           prv_Cell(prv_Rect(objEl)), _
                           prv_Handle(objEl))
End Function

' 목적   : 요소가 지원하는 제어 방식(패턴) 목록 - 이후 am_Automation 제어 함수 선택 기준
'          Invoke=클릭, Value=값 입력, Toggle=체크, SelectionItem=선택, ExpandCollapse=펼침,
'          Text=본문 읽기, Scroll=스크롤, Window=창 닫기/최소화
Private Function prv_Patterns(ByVal objEl As IUIAutomationElement) As String
    Dim strList As String

    If prv_Bool(objEl, PID_HAS_INVOKE) Then strList = strList & ", Invoke"
    If prv_Bool(objEl, PID_HAS_VALUE) Then strList = strList & ", Value"
    If prv_Bool(objEl, PID_HAS_TOGGLE) Then strList = strList & ", Toggle"
    If prv_Bool(objEl, PID_HAS_SELECTITEM) Then strList = strList & ", SelectionItem"
    If prv_Bool(objEl, PID_HAS_EXPAND) Then strList = strList & ", ExpandCollapse"
    If prv_Bool(objEl, PID_HAS_TEXT) Then strList = strList & ", Text"
    If prv_Bool(objEl, PID_HAS_SCROLL) Then strList = strList & ", Scroll"
    If prv_Bool(objEl, PID_HAS_WINDOW) Then strList = strList & ", Window"

    If Len(strList) > 0 Then prv_Patterns = Mid(strList, 3)
End Function

' 목적   : ControlType ID(50000~50040) → 이름
Private Function prv_TypeName(ByVal lngType As Long) As String
    Dim arrNames As Variant

    arrNames = Split("Button,Calendar,CheckBox,ComboBox,Edit,Hyperlink,Image,ListItem,List,Menu," & _
                     "MenuBar,MenuItem,ProgressBar,RadioButton,ScrollBar,Slider,Spinner,StatusBar,Tab,TabItem," & _
                     "Text,ToolBar,ToolTip,Tree,TreeItem,Custom,Group,Thumb,DataGrid,DataItem," & _
                     "Document,SplitButton,Window,Pane,Header,HeaderItem,Table,TitleBar,Separator,SemanticZoom," & _
                     "AppBar", ",")

    If lngType >= 50000 And lngType - 50000 <= UBound(arrNames) Then
        prv_TypeName = arrNames(lngType - 50000)
    ElseIf lngType <> 0 Then
        prv_TypeName = CStr(lngType)
    End If
End Function

' 목적   : 화면 위치 "좌,상,폭,높이" 문자열 (화면 밖·크기 0 이면 빈 문자열)
Private Function prv_Rect(ByVal objEl As IUIAutomationElement) As String
    Dim vntRect As Variant

    vntRect = prv_Prop(objEl, PID_RECT)
    If Not IsArray(vntRect) Then Exit Function
    If UBound(vntRect) - LBound(vntRect) < 3 Then Exit Function
    If vntRect(LBound(vntRect) + 2) = 0 And vntRect(LBound(vntRect) + 3) = 0 Then Exit Function

    prv_Rect = CLng(vntRect(LBound(vntRect))) & "," & CLng(vntRect(LBound(vntRect) + 1)) & "," & _
               CLng(vntRect(LBound(vntRect) + 2)) & "," & CLng(vntRect(LBound(vntRect) + 3))
End Function

' 목적   : Win32 창 핸들 (없으면 빈 값) - am_Automation 의 FindWindow 계열과 대조용
Private Function prv_Handle(ByVal objEl As IUIAutomationElement) As Variant
    Dim lngHwnd As Long

    lngHwnd = prv_Long(objEl, PID_HWND)
    If lngHwnd <> 0 Then
        prv_Handle = lngHwnd
    Else
        prv_Handle = ""
    End If
End Function

' ── 속성 읽기 (요소가 사라지는 등 오류 시 기본값) ─────────────

Private Function prv_Prop(ByVal objEl As IUIAutomationElement, ByVal lngPropId As Long) As Variant
    On Error Resume Next
    prv_Prop = objEl.GetCurrentPropertyValue(lngPropId)
End Function

Private Function prv_Text(ByVal objEl As IUIAutomationElement, ByVal lngPropId As Long) As String
    Dim vntValue As Variant

    vntValue = prv_Prop(objEl, lngPropId)
    If IsArray(vntValue) Or IsNull(vntValue) Or IsEmpty(vntValue) Or IsObject(vntValue) Then Exit Function

    On Error Resume Next
    prv_Text = CStr(vntValue)
End Function

Private Function prv_Long(ByVal objEl As IUIAutomationElement, ByVal lngPropId As Long) As Long
    On Error Resume Next
    prv_Long = CLng(prv_Prop(objEl, lngPropId))
End Function

Private Function prv_Bool(ByVal objEl As IUIAutomationElement, ByVal lngPropId As Long) As Boolean
    On Error Resume Next
    prv_Bool = CBool(prv_Prop(objEl, lngPropId))
End Function

' ── 트리 이동 (오류 시 Nothing → 순회 종료) ───────────────────

Private Function prv_FirstChild(ByVal objWalker As IUIAutomationTreeWalker, _
                                ByVal objEl As IUIAutomationElement) As IUIAutomationElement
    On Error Resume Next
    Set prv_FirstChild = objWalker.GetFirstChildElement(objEl)
End Function

Private Function prv_NextSibling(ByVal objWalker As IUIAutomationTreeWalker, _
                                 ByVal objEl As IUIAutomationElement) As IUIAutomationElement
    On Error Resume Next
    Set prv_NextSibling = objWalker.GetNextSiblingElement(objEl)
End Function

Private Function prv_Parent(ByVal objWalker As IUIAutomationTreeWalker, _
                            ByVal objEl As IUIAutomationElement) As IUIAutomationElement
    On Error Resume Next
    Set prv_Parent = objWalker.GetParentElement(objEl)
End Function

' 목적   : 시작 요소의 하위 전체에서 좌표를 포함하는 가장 작은(안쪽) 요소 반환
' 참고   : 부모 영역 밖에 있는 자식도 찾기 위해 하위 전체를 탐색
'          (예: 트리 항목은 부모 항목의 영역이 자기 글자 줄뿐이고 자식은 그 아래에 있음)
'          면적이 같으면 더 깊은 쪽(나중에 찾은 쪽)을 고름. 탐색 상한 MAX_HIT_NODES
Private Function prv_HitTest(ByVal objWalker As IUIAutomationTreeWalker, _
                             ByVal objEl As IUIAutomationElement, _
                             ByVal lngX As Long, _
                             ByVal lngY As Long) As IUIAutomationElement

    Dim objBest  As IUIAutomationElement
    Dim dblBest  As Double
    Dim lngCount As Long

    Set objBest = objEl
    If Not prv_Contains(objEl, lngX, lngY, dblBest) Then dblBest = 1E+300

    prv_HitSearch objWalker, objEl, lngX, lngY, objBest, dblBest, lngCount
    Set prv_HitTest = objBest

End Function

Private Sub prv_HitSearch(ByVal objWalker As IUIAutomationTreeWalker, _
                          ByVal objEl As IUIAutomationElement, _
                          ByVal lngX As Long, _
                          ByVal lngY As Long, _
                          ByRef objBest As IUIAutomationElement, _
                          ByRef dblBest As Double, _
                          ByRef lngCount As Long)

    Dim objChild As IUIAutomationElement
    Dim dblArea  As Double

    Set objChild = prv_FirstChild(objWalker, objEl)
    Do Until objChild Is Nothing
        lngCount = lngCount + 1
        If lngCount > MAX_HIT_NODES Then Exit Sub

        If prv_Contains(objChild, lngX, lngY, dblArea) Then
            If dblArea <= dblBest Then
                Set objBest = objChild
                dblBest = dblArea
            End If
        End If
        prv_HitSearch objWalker, objChild, lngX, lngY, objBest, dblBest, lngCount

        Set objChild = prv_NextSibling(objWalker, objChild)
    Loop

End Sub

' 목적   : 요소의 화면 위치가 좌표를 포함하는지 (dblArea 에 면적 반환)
Private Function prv_Contains(ByVal objEl As IUIAutomationElement, _
                              ByVal lngX As Long, _
                              ByVal lngY As Long, _
                              ByRef dblArea As Double) As Boolean
    Dim vntRect As Variant
    Dim lngBase As Long

    vntRect = prv_Prop(objEl, PID_RECT)
    If Not IsArray(vntRect) Then Exit Function
    If UBound(vntRect) - LBound(vntRect) < 3 Then Exit Function

    lngBase = LBound(vntRect)
    If vntRect(lngBase + 2) <= 0 Or vntRect(lngBase + 3) <= 0 Then Exit Function

    If lngX >= vntRect(lngBase) And lngX < vntRect(lngBase) + vntRect(lngBase + 2) _
       And lngY >= vntRect(lngBase + 1) And lngY < vntRect(lngBase + 1) + vntRect(lngBase + 3) Then
        dblArea = vntRect(lngBase + 2) * vntRect(lngBase + 3)
        prv_Contains = True
    End If
End Function

' ── 프로세스 ──────────────────────────────────────────────────

' 목적   : 화면 좌표 아래 창 핸들 (UI Automation 을 거치지 않는 Win32 조회)
' 참고   : 64bit 는 y(상위 32bit)·x(하위 32bit)를 한 값으로 묶음. LongLong 리터럴(&H…^)은 32bit 편집기에서
'          구문 오류(빨간 줄)로 표시되므로 Double 계산 후 변환 (좌표 범위에서 Double 은 정확)
Private Function prv_WindowAtPoint(ByVal lngX As Long, ByVal lngY As Long) As LongPtr
#If Win64 Then
    Dim dblLow As Double

    dblLow = lngX
    If dblLow < 0 Then dblLow = dblLow + 4294967296#    ' 음수 x(왼쪽 보조 모니터)를 부호 없는 32bit 로
    prv_WindowAtPoint = WindowFromPoint(CLngLng(lngY * 4294967296# + dblLow))
#Else
    prv_WindowAtPoint = WindowFromPoint(lngX, lngY)
#End If
End Function

' 목적   : 실행 중 프로세스 ID → Array(실행 파일 이름, 전체 경로) 사전 (WMI, Late Binding)
'          경로는 권한이 없는 프로세스(관리자·시스템)면 빈 값
Private Function prv_ProcessNames() As Object
    Dim dicProc As Object
    Dim objWMI  As Object
    Dim objProc As Object

    Set dicProc = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    If Not objWMI Is Nothing Then
        For Each objProc In objWMI.ExecQuery("SELECT ProcessId, Name, ExecutablePath FROM Win32_Process")
            dicProc(CLng(objProc.ProcessId)) = Array(CStr(objProc.Name), CStr(objProc.ExecutablePath & ""))
        Next objProc
    End If

    Set prv_ProcessNames = dicProc
End Function

Private Function prv_ProcName(ByVal dicProc As Object, ByVal lngPid As Long) As String
    If dicProc Is Nothing Then Exit Function
    Dim vntInfo As Variant
    If dicProc.Exists(lngPid) Then vntInfo = dicProc(lngPid): prv_ProcName = vntInfo(0)
End Function

Private Function prv_ProcPath(ByVal dicProc As Object, ByVal lngPid As Long) As String
    If dicProc Is Nothing Then Exit Function
    Dim vntInfo As Variant
    If dicProc.Exists(lngPid) Then vntInfo = dicProc(lngPid): prv_ProcPath = vntInfo(1)
End Function

' ══════════════════════════════════════════════════════════
'  내부 - 결과 출력
' ══════════════════════════════════════════════════════════

' 목적   : 결과 문자열 정리 - 줄바꿈 제거, 길이 제한
Private Function prv_Cell(ByVal strText As String) As String
    strText = Replace(Replace(Replace(strText, vbCr, " "), vbLf, " "), vbTab, " ")
    If Len(strText) > MAX_TEXT_LEN Then strText = Left(strText, MAX_TEXT_LEN) & "…"
    prv_Cell = strText
End Function

' 목적   : 머리글 + 행 Collection → 2차원 배열 (1 To 행, 1 To 열)
Private Function prv_ToArray(ByVal arrHeaders As Variant, ByVal colRows As Collection) As Variant
    Dim arrData() As Variant
    Dim vntRow    As Variant
    Dim lngCols   As Long
    Dim lngRow    As Long
    Dim lngCol    As Long

    lngCols = UBound(arrHeaders) - LBound(arrHeaders) + 1
    ReDim arrData(1 To colRows.Count + 1, 1 To lngCols)

    For lngCol = 1 To lngCols
        arrData(1, lngCol) = arrHeaders(LBound(arrHeaders) + lngCol - 1)
    Next lngCol

    lngRow = 1
    For Each vntRow In colRows
        lngRow = lngRow + 1
        For lngCol = 1 To lngCols
            arrData(lngRow, lngCol) = vntRow(LBound(vntRow) + lngCol - 1)
        Next lngCol
    Next vntRow

    prv_ToArray = arrData
End Function

' 목적   : 결과 배열을 새 시트에 기록 (1행 제목, 3행부터 표)
' 인수   : lngIndentCol - 깊이만큼 들여쓸 열 번호 (0=안 함)
'          lngDepthCol  - 깊이 값이 든 열 번호
Private Sub prv_WriteSheet(ByVal wb As Workbook, _
                           ByVal strPrefix As String, _
                           ByVal strTitle As String, _
                           ByVal arrData As Variant, _
                           ByVal lngIndentCol As Long, _
                           ByVal lngDepthCol As Long)

    Dim ws       As Worksheet
    Dim lngRows  As Long
    Dim lngCols  As Long
    Dim lngRow   As Long
    Dim lngCol   As Long
    Dim lngDepth As Long

    lngRows = UBound(arrData, 1)
    lngCols = UBound(arrData, 2)

    ' 문자열 앞에 ' 를 붙여 숫자·날짜·수식으로 바뀌지 않게 함 (셀에는 표시되지 않음)
    For lngRow = 1 To lngRows
        For lngCol = 1 To lngCols
            If VarType(arrData(lngRow, lngCol)) = vbString Then
                If Len(arrData(lngRow, lngCol)) > 0 Then arrData(lngRow, lngCol) = "'" & arrData(lngRow, lngCol)
            End If
        Next lngCol
    Next lngRow

    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    On Error Resume Next
    ws.Name = Left(strPrefix & "_" & Format(Now, "mmdd_hhnnss"), 31)
    On Error GoTo 0

    ws.Range("A1").Value = strTitle
    ws.Range("A1").Font.Bold = True

    With ws.Range("A3").Resize(lngRows, lngCols)
        .Value = arrData
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(221, 235, 247)
        .AutoFilter
    End With

    If lngIndentCol > 0 Then
        For lngRow = 2 To lngRows
            lngDepth = arrData(lngRow, lngDepthCol)
            If lngDepth > MAX_INDENT Then lngDepth = MAX_INDENT
            If lngDepth > 0 Then ws.Cells(lngRow + 2, lngIndentCol).IndentLevel = lngDepth
        Next lngRow
    End If

    ' 제목(A1)은 제외하고 표 범위 기준으로만 열 너비 맞춤
    ws.Range("A3").Resize(lngRows, lngCols).Columns.AutoFit
    For lngCol = 1 To lngCols
        If ws.Columns(lngCol).ColumnWidth > 60 Then ws.Columns(lngCol).ColumnWidth = 60
    Next lngCol

    wb.Activate
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True

End Sub
