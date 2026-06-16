Attribute VB_Name = "ref_File"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_File                                               │
' │  역할 : am_File Application.Run 래퍼                   │
' └─────────────────────────────────────────────────────────┘

' 참고 : GetFoldersList 는 ByRef 배열 파라미터로 인해 래핑 불가
'        → Application.Run "corelib.xlam!am_File.GetFoldersList", ... 직접 호출

Private Const REF As String = "corelib.xlam!am_File."

' ── 다이얼로그 ────────────────────────────────────────────────

' 목적   : 파일/폴더 선택 다이얼로그 호출
' 인수   : strCwbPath     - 호출한 CWB 폴더 경로
'          DilogType      - 다이얼로그 종류 (기본: msoFileDialogFilePicker)
'          strInitPath    - 초기 열기 경로 (기본: CWB 경로)
'          blnMultiSelect - True: 다중 선택 허용
'          strExtComment  - 확장자 설명
'          strFileName    - 초기 파일명
'          strExt         - 확장자 필터 (기본: *.*)
'          strCancelPath  - 취소 시 사용할 기본 경로
'          blnToken       - True: 반환 경로를 {cPath} 토큰으로 역치환
' 반환   : String - 선택된 경로 (다중 선택 시 ";" 로 연결)
' 예시   : GetPath(ThisWorkbook.Path)
'          Split(GetPath(ThisWorkbook.Path, blnMultiSelect:=True), ";")
Public Function GetPath(ByVal strCwbPath As String, _
                        Optional ByVal DilogType As MsoFileDialogType = msoFileDialogFilePicker, _
                        Optional ByVal strInitPath As String = "", _
                        Optional ByVal blnMultiSelect As Boolean = False, _
                        Optional ByVal strExtComment As String = "Select Item", _
                        Optional ByVal strFileName As String = "", _
                        Optional ByVal strExt As String = "*.*", _
                        Optional ByVal strCancelPath As String = "", _
                        Optional ByVal blnToken As Boolean = True) As String
    GetPath = Application.Run(REF & "GetPath", _
                              strCwbPath, DilogType, strInitPath, blnMultiSelect, _
                              strExtComment, strFileName, strExt, strCancelPath, blnToken)
End Function

' ── 파일/폴더 존재 확인 ──────────────────────────────────────

' 목적   : 파일 존재 여부 확인
' 인수   : strPath - 확인할 파일 경로
' 반환   : Boolean - True: 존재 / False: 없음
' 예시   : CheckFileExistence("C:\test.xlsx") → True
Public Function CheckFileExistence(ByVal strPath As String) As Boolean
    CheckFileExistence = Application.Run(REF & "CheckFileExistence", strPath)
End Function

' 목적   : 폴더 존재 여부 확인
' 인수   : strPath - 확인할 폴더 경로
' 반환   : Boolean - True: 존재 / False: 없음
' 예시   : CheckFolderExistence("C:\TestFolder") → True
Public Function CheckFolderExistence(ByVal strPath As String) As Boolean
    CheckFolderExistence = Application.Run(REF & "CheckFolderExistence", strPath)
End Function

' ── 파일/폴더 조작 ───────────────────────────────────────────

' 목적   : 파일 확장자 반환
' 인수   : strFileName - 파일명 또는 전체 경로
' 반환   : String - 확장자 (예: ".xlsx")
' 예시   : GetExt("test.xlsx") → ".xlsx"
Public Function GetExt(ByVal strFileName As String) As String
    GetExt = Application.Run(REF & "GetExt", strFileName)
End Function

' 목적   : 폴더 생성 (중간 경로 없어도 자동 생성)
' 인수   : strPath - 생성할 폴더 경로
' 예시   : MkFolder("C:\A\B\C")
Public Sub MkFolder(ByVal strPath As String)
    Application.Run REF & "MkFolder", strPath
End Sub

' 목적   : 폴더 삭제 (내부 파일·하위 폴더 포함, 재귀)
' 인수   : strPath - 삭제할 폴더 경로
' 예시   : DelFolder("C:\TestFolder")
Public Sub DelFolder(ByVal strPath As String)
    Application.Run REF & "DelFolder", strPath
End Sub

' 목적   : 파일 삭제
' 인수   : strPath - 삭제할 파일 경로
' 예시   : DelFile("C:\test.xlsx")
Public Sub DelFile(ByVal strPath As String)
    Application.Run REF & "DelFile", strPath
End Sub

' 목적   : 파일 이름 변경 또는 이동
' 인수   : strOldPath - 원본 파일 경로
'          strNewPath - 변경할 파일 경로
' 예시   : RenFile("C:\old.xlsx", "C:\new.xlsx")
Public Sub RenFile(ByVal strOldPath As String, ByVal strNewPath As String)
    Application.Run REF & "RenFile", strOldPath, strNewPath
End Sub

' 목적   : 파일 복사
' 인수   : strSrcPath  - 원본 파일 경로
'          strDestPath - 복사할 파일 경로
' 예시   : CopyFile("C:\src.xlsx", "D:\dest.xlsx")
Public Sub CopyFile(ByVal strSrcPath As String, ByVal strDestPath As String)
    Application.Run REF & "CopyFile", strSrcPath, strDestPath
End Sub

' 목적   : 폴더 복사
' 인수   : strSrcPath  - 원본 폴더 경로
'          strDestPath - 복사할 폴더 경로
' 예시   : CopyFolder("C:\SrcFolder", "D:\DestFolder")
Public Sub CopyFolder(ByVal strSrcPath As String, ByVal strDestPath As String)
    Application.Run REF & "CopyFolder", strSrcPath, strDestPath
End Sub

' ── 파일/폴더 검색 ───────────────────────────────────────────

' 목적   : 파일 목록 수집
' 인수   : strPath       - 검색할 루트 경로
'          strFilter     - 파일명 필터 (기본: *)
'          blnSubFolders - True: 하위 폴더 파일도 검색
' 반환   : Variant - 파일 전체 경로 배열
' 예시   : arr = GetFilesList("C:\Root", "*.xlsx")
Public Function GetFilesList(ByVal strPath As String, _
                             Optional ByVal strFilter As String = "*", _
                             Optional ByVal blnSubFolders As Boolean = True) As Variant
    GetFilesList = Application.Run(REF & "GetFilesList", strPath, strFilter, blnSubFolders)
End Function
