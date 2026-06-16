Attribute VB_Name = "ref_Path"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Path                                               │
' │  역할 : am_Path Application.Run 래퍼                   │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Path."

' 목적   : 토큰이 포함된 경로를 실제 로컬 경로로 반환
' 인수   : strPath        - 변환할 경로 (토큰 포함 가능)
'          strCwbPath     - 호출한 CWB 의 폴더 경로
'          strCwbFile     - 호출한 CWB 의 전체 파일 경로
'          blnChangeDrive - True: 드라이브를 CWB 드라이브로 강제 교체
'          blnAddDrive    - True: 드라이브 없는 경로에 CWB 드라이브 추가
' 반환   : String - 변환된 로컬 경로, 실패 시 ""
' 예시   : ReplacePath("{cPath}\DB\main.accdb", ThisWorkbook.Path, ThisWorkbook.FullName)
Public Function ReplacePath(ByVal strPath As String, _
                            ByVal strCwbPath As String, _
                            ByVal strCwbFile As String, _
                            Optional ByVal blnChangeDrive As Boolean = False, _
                            Optional ByVal blnAddDrive As Boolean = True) As String
    ReplacePath = Application.Run(REF & "ReplacePath", _
                                  strPath, strCwbPath, strCwbFile, _
                                  blnChangeDrive, blnAddDrive)
End Function
