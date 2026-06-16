Attribute VB_Name = "ref_Error"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Error                                              │
' │  역할 : am_Error Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Error."

' ── 로그 활성화 제어 ─────────────────────────────────────────

' 목적   : 에러 로그 파일 기록 활성화/비활성화
' 인수   : blnEnabled - True: 로그 파일 기록 / False: 기록 안 함
' 예시   : SetLogEnabled True   ' 로그 파일 기록 시작
'          SetLogEnabled False  ' 로그 파일 기록 중단
Public Sub SetLogEnabled(ByVal blnEnabled As Boolean)
    Application.Run REF & "SetLogEnabled", blnEnabled
End Sub

' 목적   : 현재 에러 로그 기록 활성화 상태 반환
' 반환   : Boolean - True: 로그 기록 중 / False: 비활성화 상태
' 예시   : If GetLogEnabled Then ...
Public Function GetLogEnabled() As Boolean
    GetLogEnabled = Application.Run(REF & "GetLogEnabled")
End Function

' ── 에러 핸들링 ───────────────────────────────────────────────

' 목적   : 오류 발생 시 메시지 표시 및 로그 기록
'          ErrHandler 레이블에서 호출하는 표준 에러 핸들러
' 인수   : strProcName    - 오류 발생 프로시저명 (기본: "")
'          strAdditional  - 추가 설명 (기본: "")
'          blnShowMessage - True: MsgBox 표시 (기본: True)
' 예시   : ErrHandler:
'              HandleError "Sub_이름", "추가 설명"
Public Sub HandleError(Optional ByVal strProcName As String = "", _
                       Optional ByVal strAdditional As String = "", _
                       Optional ByVal blnShowMessage As Boolean = True)
    Application.Run REF & "HandleError", strProcName, strAdditional, blnShowMessage
End Sub

' 목적   : 로그 파일에 메시지 기록 (에러 외 일반 로그)
' 인수   : strMessage - 기록할 메시지 문자열
'          strType    - 로그 유형 레이블 (기본: "INFO", 예: "WARN", "ERROR")
' 예시   : WriteLog "작업 시작", "INFO"
'          WriteLog "파일 없음: " & strPath, "WARN"
Public Sub WriteLog(ByVal strMessage As String, _
                    Optional ByVal strType As String = "INFO")
    Application.Run REF & "WriteLog", strMessage, strType
End Sub
