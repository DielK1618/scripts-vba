Attribute VB_Name = "ref_Core"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Core                                               │
' │  역할 : am_Core Application.Run 래퍼                   │
' └─────────────────────────────────────────────────────────┘

Private Const REF     As String = "corelib.xlam!am_Core."
Public  Const LOCK_PW As String = "1234"    ' SheetLock / WB_Lock 기본 비밀번호 — 프로젝트에 맞게 수정

' ── Application 상태 관리 ────────────────────────────────────

' 목적   : 화면 갱신·경고창 비활성화 (작업 시작 전 호출)
' 예시   : DPUpdate_Off
Public Sub DPUpdate_Off()
    Application.Run REF & "DPUpdate_Off"
End Sub

' 목적   : 화면 갱신·경고창 활성화 복원 (작업 완료 후 호출)
' 예시   : DPUpdate_On
Public Sub DPUpdate_On()
    Application.Run REF & "DPUpdate_On"
End Sub

' 목적   : 이벤트 비활성화
' 예시   : Event_Off
Public Sub Event_Off()
    Application.Run REF & "Event_Off"
End Sub

' 목적   : 이벤트 활성화 복원
' 예시   : Event_On
Public Sub Event_On()
    Application.Run REF & "Event_On"
End Sub

' 목적   : 자동 계산 중단 (수동 모드 전환)
' 예시   : Calculate_Off
Public Sub Calculate_Off()
    Application.Run REF & "Calculate_Off"
End Sub

' 목적   : 자동 계산 복원
' 예시   : Calculate_On
Public Sub Calculate_On()
    Application.Run REF & "Calculate_On"
End Sub

' ── 워크북 보호 ───────────────────────────────────────────────

' 목적   : 워크북 구조(시트 추가·삭제·이동) 보호
' 인수   : wb    - 보호할 워크북 (기본: ActiveWorkbook)
'          strPW - 보호 비밀번호 (기본: LOCK_PW)
' 예시   : WB_Lock
'          WB_Lock ThisWorkbook, "pw"
Public Sub WB_Lock(Optional ByVal wb As Workbook = Nothing, _
                   Optional ByVal strPW As String = LOCK_PW)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Application.Run REF & "WB_Lock", wb, strPW
End Sub

' 목적   : 워크북 구조 보호 해제
' 인수   : wb    - 해제할 워크북 (기본: ActiveWorkbook)
'          strPW - 보호 비밀번호 (기본: LOCK_PW)
' 예시   : WB_UnLock
'          WB_UnLock ThisWorkbook, "pw"
Public Sub WB_UnLock(Optional ByVal wb As Workbook = Nothing, _
                     Optional ByVal strPW As String = LOCK_PW)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Application.Run REF & "WB_UnLock", wb, strPW
End Sub

' ── 상태 확인 ────────────────────────────────────────────────

' 목적   : corelib.xlam 이 현재 열려있는지 확인
' 반환   : Boolean - True: 로드됨 / False: 미로드
' 예시   : If IsXlamLoaded Then ...
Public Function IsXlamLoaded() As Boolean
    IsXlamLoaded = Application.Run(REF & "IsXlamLoaded")
End Function

' 목적   : xlam 초기화 완료 상태 반환
' 반환   : Boolean - True: 정상 초기화 / False: 미초기화
' 예시   : If IsReady Then ...
Public Function IsReady() As Boolean
    IsReady = Application.Run(REF & "IsReady")
End Function

' 목적   : xlam 폴더 경로 반환
' 반환   : String - 예) "D:\Tools"
' 예시   : XlamPath → "D:\Tools"
Public Function XlamPath() As String
    XlamPath = Application.Run(REF & "XlamPath")
End Function

' 목적   : xlam 전체 파일 경로 반환
' 반환   : String - 예) "D:\Tools\corelib.xlam"
' 예시   : XlamFullName → "D:\Tools\corelib.xlam"
Public Function XlamFullName() As String
    XlamFullName = Application.Run(REF & "XlamFullName")
End Function

' 목적   : xlam 버전 반환
' 반환   : String - 예) "1.0.0"
' 예시   : Version → "1.0.0"
Public Function Version() As String
    Version = Application.Run(REF & "Version")
End Function
