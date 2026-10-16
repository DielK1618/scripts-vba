Attribute VB_Name = "tpl_Messenger"
Option Explicit
' 버전 : 2026-10-16 16:32 · 신규 - WMC Messenger 자동화

' ┌─────────────────────────────────────────────────────────┐
' │  tpl_Messenger                                           │
' │  역할 : WMC Messenger 자동화 (실행·로그인·그룹 우클릭 메뉴) │
' └─────────────────────────────────────────────────────────┘
' ■ 필요 모듈
'   CWB  : ref_Automation, ref_Credential, tpl_Messenger(이 모듈)
'   xlam : am_Automation, am_Credential + VBE 참조 "UIAutomationClient"
'
' ■ 빠른 사용법 (직접 실행 창 - 인수가 있으면 괄호 필수)
'   ?WmcOpen                                    ' 실행 → 로그인 → 사용자 목록 준비까지
'   ?WmcAddFriend("Target")                     ' Target 그룹 우클릭 → 친구추가 (WmcOpen 포함)
'   ?WmcRunGroupMenu("Target", "쪽지 보내기")   ' 다른 우클릭 메뉴 실행
'   결과 True = 성공 / False = 실패 → 직접 실행 창의 "am_Automation.…:" / "tpl_Messenger.…:" 줄이 실패 단계
'
' ■ 로그인 방식 - 아래 WMC_LOGIN_MODE 상수로 선택
'   "prompt" (기본) : 로그인할 때마다 Windows 자격 증명 입력 창(비밀번호 ● 표시)으로 받음. 어디에도 저장 안 함
'   "stored"        : Windows 자격 증명 관리자에 저장된 값으로 자동 로그인
'                     처음 한 번 WmcSaveCredential 실행 → 아이디·비밀번호 입력 → 저장
'                     ※ 회사 보안 정책 확인 후 사용 - 매크로가 저장된 비밀번호를 읽는 동작은 보안 솔루션이
'                       경고·차단할 수 있음 (2026-10-16 협의: 지금은 "prompt" 로 사용)
'   "manual"        : 로그인 창만 띄우고, 사람이 직접 로그인할 때까지 기다림 (WMC_MANUAL_WAIT_SEC 초)
'
' ■ 로그인 창 세 번째 칸(txLifeDispNo)
'   WMC_LIFEDISPNO 상수에 값을 넣으면 그 값을 자동 입력, 비워 두면 로그인할 때마다 입력 창으로 받음
'
' ■ 자격 증명 관리자 ("stored" 방식용)
'   WmcSaveCredential      ' 입력 창 → 저장 (비밀번호를 바꾸면 다시 실행)
'   WmcDeleteCredential    ' 삭제
'   ?WmcHasCredential      ' 저장 여부 확인
'   제어판 > 자격 증명 관리자 > Windows 자격 증명 > 일반 자격 증명 에서 WMC_CRED_TARGET 이름으로도 확인·삭제 가능
'
' ■ WMC 화면 구조 (2026-10-16 am_Inspect 로 확인 - 바뀌면 DumpWindows / DumpElementTree 로 다시 확인)
'   로그인 창 : 제목 "Login" / 입력칸 #teUserID, #tePassword, #txLifeDispNo / 로그인 버튼 #btnLogIn
'   메인 창   : 제목 "WMC Messenger" / 사용자 탭 트리 tree_user 의 그룹 = TreeItem (이름 "Target (52/212)" 형식)
'   우클릭 메뉴 : 실제 마우스 우클릭에서만 열림 → ClickContextMenuByUIA 가 항목 위치를 계산해 우클릭 (커서는 원위치)
'                 KM Link 로 입력을 다른 PC 에 넘긴 상태에서는 실패할 수 있음 (TROUBLESHOOTING.md TS-V16)

' ── 설정 ──────────────────────────────────────────────────────
Private Const WMC_EXE             As String = "C:\Program Files (x86)\WMC\WMC MSG V4\WMC_Msg.exe"
Private Const WMC_WINDOW          As String = "WMC Messenger"        ' 메인 창 제목 (부분 일치)
Private Const WMC_LOGIN_WINDOW    As String = "=Login"               ' 로그인 창 제목 (완전 일치)
Private Const WMC_READY_ITEM      As String = "Target"               ' 목록 불러오기 완료 확인용 항목
Private Const WMC_LOGIN_MODE      As String = "prompt"               ' "prompt" / "stored" / "manual"
Private Const WMC_LIFEDISPNO      As String = ""                     ' 세 번째 칸 고정값 (비우면 매번 입력)
Private Const WMC_CRED_TARGET     As String = "corelib/WMC_Messenger" ' 자격 증명 관리자 저장 이름
Private Const WMC_START_WAIT_SEC  As Long = 60                       ' 실행 후 창 대기
Private Const WMC_LOGIN_WAIT_SEC  As Long = 30                       ' 로그인 버튼 후 로그인 창이 닫힐 때까지
Private Const WMC_MANUAL_WAIT_SEC As Long = 180                      ' "manual" 방식에서 사람이 로그인할 시간
Private Const WMC_LIST_WAIT_SEC   As Long = 90                       ' 로그인 후 사용자 목록 대기

Private Const PROC As String = "tpl_Messenger."

' ══════════════════════════════════════════════════════════
'  실행·로그인
' ══════════════════════════════════════════════════════════

' 목적   : WMC 를 실행(또는 앞으로)하고, 로그인 창이 있으면 로그인한 뒤, 사용자 목록이 준비될 때까지 대기
' 반환   : Boolean - 목록까지 준비되었는지 여부
' 예시   : If Not WmcOpen Then Exit Sub
Public Function WmcOpen() As Boolean
    If Not OpenAppWindow(WMC_WINDOW & "|" & WMC_LOGIN_WINDOW, WMC_EXE, "", WMC_START_WAIT_SEC) Then
        Debug.Print PROC & "WmcOpen: WMC 창을 열지 못함"
        Exit Function
    End If

    If IsWindowOpen(WMC_LOGIN_WINDOW) Then
        If Not prv_Login() Then Exit Function
    End If

    WmcOpen = WaitElementByUIA(WMC_WINDOW, WMC_READY_ITEM, WMC_LIST_WAIT_SEC)
    If Not WmcOpen Then Debug.Print PROC & "WmcOpen: 사용자 목록이 준비되지 않음 (" & WMC_READY_ITEM & ")"
End Function

' ══════════════════════════════════════════════════════════
'  그룹 우클릭 메뉴
' ══════════════════════════════════════════════════════════

' 목적   : 그룹 우클릭 → 친구추가
' 인수   : strGroup - 그룹 이름 (앞부분 일치 - "Target" → "Target (52/212)")
' 반환   : Boolean
' 예시   : ?WmcAddFriend("Target")
Public Function WmcAddFriend(ByVal strGroup As String) As Boolean
    WmcAddFriend = WmcRunGroupMenu(strGroup, "친구추가")
End Function

' 목적   : 그룹 우클릭 → 지정한 메뉴 실행 (WMC 실행·로그인 포함)
' 인수   : strGroup - 그룹 이름 (앞부분 일치)
'          strMenu  - 메뉴 이름 (앞부분 일치 - 예: "친구추가", "쪽지 보내기", "대화하기")
' 반환   : Boolean
' 예시   : ?WmcRunGroupMenu("Target", "쪽지 보내기")
Public Function WmcRunGroupMenu(ByVal strGroup As String, ByVal strMenu As String) As Boolean
    If Not WmcOpen() Then Exit Function
    WmcRunGroupMenu = ClickContextMenuByUIA(WMC_WINDOW, strGroup, strMenu, "TreeItem")
End Function

' ══════════════════════════════════════════════════════════
'  자격 증명 관리자 ("stored" 방식용)
' ══════════════════════════════════════════════════════════

' 목적   : 입력 창으로 아이디·비밀번호를 받아 자격 증명 관리자에 저장
Public Sub WmcSaveCredential()
    Dim vntCred As Variant

    vntCred = PromptCredential("WMC Messenger 로그인 정보 저장", _
                               "저장할 WMC Messenger 아이디와 비밀번호를 입력하세요." & vbCrLf & _
                               "(Windows 자격 증명 관리자에 '" & WMC_CRED_TARGET & "' 이름으로 저장)")
    If Not IsArray(vntCred) Then Exit Sub

    If SaveCredential(WMC_CRED_TARGET, vntCred(0), vntCred(1)) Then
        MsgBox "저장했습니다. WMC_LOGIN_MODE 를 ""stored"" 로 바꾸면 자동 로그인합니다.", vbInformation, "WMC Messenger"
    Else
        MsgBox "저장하지 못했습니다. 회사 정책으로 막혀 있을 수 있습니다.", vbExclamation, "WMC Messenger"
    End If
    vntCred = Empty
End Sub

' 목적   : 자격 증명 관리자에서 WMC 로그인 정보 삭제
Public Sub WmcDeleteCredential()
    If DeleteCredential(WMC_CRED_TARGET) Then
        MsgBox "삭제했습니다.", vbInformation, "WMC Messenger"
    Else
        MsgBox "저장된 로그인 정보가 없습니다.", vbInformation, "WMC Messenger"
    End If
End Sub

' 목적   : 자격 증명 관리자에 WMC 로그인 정보가 있는지
Public Function WmcHasCredential() As Boolean
    WmcHasCredential = HasCredential(WMC_CRED_TARGET)
End Function

' ══════════════════════════════════════════════════════════
'  내부
' ══════════════════════════════════════════════════════════

' 목적   : 로그인 창에 입력하고 로그인 버튼 실행 → 로그인 창이 닫히면 성공
Private Function prv_Login() As Boolean
    Dim vntCred As Variant
    Dim strNo   As String

    ' ── 1. 로그인 정보 가져오기 ──
    Select Case WMC_LOGIN_MODE
        Case "manual"
            Debug.Print PROC & "prv_Login: 로그인 창에서 직접 로그인하세요 (" & WMC_MANUAL_WAIT_SEC & "초 대기)"
            prv_Login = prv_WaitLoginClosed(WMC_MANUAL_WAIT_SEC)
            Exit Function

        Case "stored"
            vntCred = ReadCredential(WMC_CRED_TARGET)
            If Not IsArray(vntCred) Then
                Debug.Print PROC & "prv_Login: 저장된 로그인 정보 없음 - WmcSaveCredential 을 먼저 실행하세요"
                Exit Function
            End If

        Case Else    ' "prompt"
            vntCred = PromptCredential("WMC Messenger 로그인", "WMC Messenger 아이디와 비밀번호를 입력하세요.")
            If Not IsArray(vntCred) Then
                Debug.Print PROC & "prv_Login: 로그인 입력을 취소함"
                Exit Function
            End If
    End Select

    strNo = WMC_LIFEDISPNO
    If Len(strNo) = 0 Then
        strNo = InputBox("로그인 창 세 번째 칸에 넣을 값을 입력하세요.", "WMC Messenger 로그인")
        If Len(strNo) = 0 Then
            Debug.Print PROC & "prv_Login: 세 번째 칸 입력을 취소함"
            vntCred = Empty
            Exit Function
        End If
    End If

    ' ── 2. 입력·로그인 ──
    If Not SetValueByUIA(WMC_LOGIN_WINDOW, "#teUserID", vntCred(0)) Then GoTo Fail
    If Not SetValueByUIA(WMC_LOGIN_WINDOW, "#tePassword", vntCred(1)) Then GoTo Fail
    If Not SetValueByUIA(WMC_LOGIN_WINDOW, "#txLifeDispNo", strNo) Then GoTo Fail
    vntCred = Empty

    If Not InvokeByUIA(WMC_LOGIN_WINDOW, "#btnLogIn") Then GoTo Fail

    ' ── 3. 로그인 창이 닫히면 성공 (남아 있으면 비밀번호 오류 등) ──
    prv_Login = prv_WaitLoginClosed(WMC_LOGIN_WAIT_SEC)
    If Not prv_Login Then Debug.Print PROC & "prv_Login: 로그인 창이 닫히지 않음 - 아이디·비밀번호·세 번째 칸 값을 확인하세요"
    Exit Function

Fail:
    vntCred = Empty
    Debug.Print PROC & "prv_Login: 로그인 창 입력 실패"
End Function

' 목적   : 로그인 창이 닫힐 때까지 대기
Private Function prv_WaitLoginClosed(ByVal lngTimeoutSec As Long) As Boolean
    Dim dtLimit As Date

    dtLimit = Now + TimeSerial(0, 0, lngTimeoutSec)
    Do
        If Not IsWindowOpen(WMC_LOGIN_WINDOW) Then
            prv_WaitLoginClosed = True
            Exit Function
        End If
        If Now > dtLimit Then Exit Function
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
End Function
