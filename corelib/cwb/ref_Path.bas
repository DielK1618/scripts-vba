Attribute VB_Name = "ref_Path"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Path                                               │
' │  역할 : am_Path Application.Run 래퍼                   │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Path."

' 목적   : 토큰이 포함된 경로를 실제 로컬 경로로 반환
' 인수   : strPath        - 변환할 경로 (토큰 포함 가능)
'          blnChangeDrive - True: 드라이브를 CWB 드라이브로 강제 교체 (UNC 경로 제외)
'          blnAddDrive    - True: 드라이브 없는 경로에 CWB 드라이브 추가 (UNC 경로 제외)
'          blnCreate      - True: 폴더가 없으면 생성 (끝이 "\" → 전체 폴더 / 아니면 부모 폴더만)
'          strCwbFile     - CWB 전체 파일 경로 (기본: ThisWorkbook.FullName, 폴더 경로 불가)
' 반환   : String - 변환된 경로, 실패 시 "" (드라이브·공유 루트 접근 불가, 폴더 생성 실패)
' 예시   : ReplacePath("{cPath}\DB\main.accdb")
'          ReplacePath("{cPath}\00_Bak\", blnCreate:=True)                    ' 폴더 생성
'          ReplacePath("{cPath}\Report\" & strName & ".pdf", blnCreate:=True) ' 부모 폴더만 생성
'          ReplacePath(strPath, True)                                         ' 드라이브 강제 교체
Public Function ReplacePath(ByVal strPath As String, _
                            Optional ByVal blnChangeDrive As Boolean = False, _
                            Optional ByVal blnAddDrive As Boolean = True, _
                            Optional ByVal blnCreate As Boolean = False, _
                            Optional ByVal strCwbFile As String = "") As String
    If strCwbFile = "" Then strCwbFile = ThisWorkbook.FullName
    ReplacePath = Application.Run(REF & "ReplacePath", _
                                  strPath, blnChangeDrive, blnAddDrive, blnCreate, strCwbFile)
End Function

' 목적   : 네트워크(UNC)·OneDrive URL 경로를 로컬 경로로 정규화 (폴더/파일 공용)
' 인수   : strPath - 정규화할 경로
' 반환   : String - 로컬 경로 (슬래시 → 백슬래시, 끝 "\" 제거)
' 예시   : NormalizeLocalPath(ThisWorkbook.Path)
Public Function NormalizeLocalPath(ByVal strPath As String) As String
    NormalizeLocalPath = Application.Run(REF & "NormalizeLocalPath", strPath)
End Function
