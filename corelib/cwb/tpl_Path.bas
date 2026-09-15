Attribute VB_Name = "tpl_Path"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  tpl_Path                                               │
' │  역할 : CWB 경로 정보 및 경로 관련 유틸리티            │
' └─────────────────────────────────────────────────────────┘

' 참고 : 토큰 경로 변환은 ref_Path.ReplacePath 사용 (CWB 파일 자동 입력)
'        → ReplacePath("{cPath}\DB\main.accdb")

' ── 1. CWB 폴더 경로 ─────────────────────────────────────────
' 목적   : CWB 폴더 경로 반환
' 반환   : String - CWB 폴더 경로
' 예시   : CwbPath → "H:\Repository\Excel\cwb"
Public Property Get CwbPath() As String
    CwbPath = ThisWorkbook.Path
End Property

' ── 2. CWB 전체 파일 경로 ────────────────────────────────────
' 목적   : CWB 전체 파일 경로 반환
' 반환   : String - CWB 전체 파일 경로
' 예시   : CwbFile → "H:\Repository\Excel\cwb\cwb_01.xlsm"
Public Property Get CwbFile() As String
    CwbFile = ThisWorkbook.FullName
End Property

' ── 3. 경로 존재 여부 확인 ───────────────────────────────────
' 목적   : 파일 또는 폴더 존재 여부 확인
' 인수   : strPath - 확인할 경로 (토큰 포함 가능)
' 반환   : Boolean - True: 존재 / False: 없음
' 예시   : If PathExists("{cPath}\DB\main.accdb") Then ...
Public Function PathExists(ByVal strPath As String) As Boolean
    Dim strResolved As String
    strResolved = ReplacePath(strPath)
    If strResolved = "" Then Exit Function
    PathExists = (Dir(strResolved, vbDirectory) <> "") Or _
                 (Dir(strResolved) <> "")
End Function
