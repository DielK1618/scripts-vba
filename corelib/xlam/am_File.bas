Attribute VB_Name = "am_File"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_File                                                │
' │  역할 : 파일/폴더 생성, 삭제, 복사, 검색, 다이얼로그   │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  다이얼로그
' ══════════════════════════════════════════════════════════

' 목적   : 파일/폴더 선택 다이얼로그 호출
' 인수   : DilogType      - 다이얼로그 종류 (기본: msoFileDialogFilePicker)
'          strInitPath    - 초기 열기 폴더 (기본: CWB 폴더)
'          blnMultiSelect - True: 다중 선택 허용
'          strExtComment  - 확장자 설명
'          strFileName    - 초기 파일명
'          strExt         - 확장자 필터 (기본: *.*)
'          strCancelPath  - 취소 시 사용할 기본 경로
'          blnToken       - True: 반환 경로 중 CWB 폴더 부분을 {cPath} 토큰으로 역치환
'          strCwbPath     - 호출한 CWB 폴더 경로 (로컬 경로로 정규화된 값, 파일 경로 불가)
'                           strInitPath 생략 또는 blnToken = True 일 때 필수
' 반환   : String - 선택된 경로 (다중 선택 시 ";" 로 연결), 취소 시 ""
' 예시   : GetPath(strCwbPath:=ThisWorkbook.Path)
'          ref_File 래퍼 사용 시 strCwbPath 생략 → GetPath(), Split(GetPath(blnMultiSelect:=True), ";")
Public Function GetPath(Optional ByVal DilogType As MsoFileDialogType = msoFileDialogFilePicker, _
                        Optional ByVal strInitPath As String = "", _
                        Optional ByVal blnMultiSelect As Boolean = False, _
                        Optional ByVal strExtComment As String = "Select Item", _
                        Optional ByVal strFileName As String = "", _
                        Optional ByVal strExt As String = "*.*", _
                        Optional ByVal strCancelPath As String = "", _
                        Optional ByVal blnToken As Boolean = True, _
                        Optional ByVal strCwbPath As String = "") As String

    Dim fdg           As FileDialog
    Dim arrSelected() As String
    Dim i             As Integer

    strCwbPath = prv_TrimTrailingSep(strCwbPath)
    If strInitPath = "" Then strInitPath = strCwbPath
    strInitPath = prv_TrimTrailingSep(strInitPath)

    If strCwbPath = "" And (strInitPath = "" Or blnToken) Then
        MsgBox "[am_File.GetPath] CWB 폴더 경로(strCwbPath)가 전달되지 않았습니다.", _
               vbCritical, am_Core.AM_NAME
        Exit Function
    End If

    Set fdg = Application.FileDialog(DilogType)

    With fdg
        On Error Resume Next
        .Filters.Clear
        .Filters.Add strExtComment, strExt
        .InitialView = msoFileDialogViewLargeIcons
        .AllowMultiSelect = blnMultiSelect
        ' 폴더와 파일명을 한 번에 지정 (따로 지정하면 뒤의 값이 폴더를 덮어씀)
        .InitialFileName = strInitPath & "\" & strFileName
        On Error GoTo 0

        ' ── 선택 완료 ────────────────────────────────────────
        If .Show = -1 Then
            ReDim arrSelected(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                arrSelected(i) = .SelectedItems(i)
                If blnToken Then arrSelected(i) = prv_ToCwbToken(arrSelected(i), strCwbPath)
            Next i
            GetPath = Join(arrSelected, ";")

        ' ── 선택 취소 ────────────────────────────────────────
        Else
            GetPath = ""
            If strCancelPath <> "" Then
                If MsgBox("선택된 경로 없습니다." & vbCrLf & _
                          "[" & strCancelPath & "] 을 경로로 사용하시겠습니까?", _
                          vbQuestion + vbYesNo) = vbYes Then
                    GetPath = strCancelPath
                    Exit Function
                End If
            End If
        End If
    End With

    Set fdg = Nothing

End Function

' ══════════════════════════════════════════════════════════
'  파일/폴더 존재 확인
' ══════════════════════════════════════════════════════════

' 목적   : 파일 존재 여부 확인
' 인수   : strPath - 확인할 파일 경로
' 반환   : Boolean - True: 존재 / False: 없음
' 예시   : CheckFileExistence("C:\test.xlsx") → True
Public Function CheckFileExistence(ByVal strPath As String) As Boolean
    On Error Resume Next
    CheckFileExistence = (Dir(strPath) <> "")
    If Err.Number <> 0 Then CheckFileExistence = False
    On Error GoTo 0
End Function

' 목적   : 폴더 존재 여부 확인
' 인수   : strPath - 확인할 폴더 경로
' 반환   : Boolean - True: 존재 / False: 없음
' 예시   : CheckFolderExistence("C:\TestFolder") → True
Public Function CheckFolderExistence(ByVal strPath As String) As Boolean
    On Error Resume Next
    CheckFolderExistence = (Dir(strPath, vbDirectory) <> "")
    If Err.Number <> 0 Then CheckFolderExistence = False
    On Error GoTo 0
End Function

' ══════════════════════════════════════════════════════════
'  파일/폴더 조작
' ══════════════════════════════════════════════════════════

' 목적   : 파일 확장자 반환
' 인수   : strFileName - 파일명 또는 전체 경로
' 반환   : String - 확장자 (예: ".xlsx")
' 예시   : GetExt("test.xlsx") → ".xlsx"
Public Function GetExt(ByVal strFileName As String) As String
    GetExt = Mid(strFileName, InStrRev(strFileName, "."))
End Function

' 목적   : 폴더 생성 (중간 경로 없어도 자동 생성)
' 인수   : strPath - 생성할 폴더 경로
' 예시   : MkFolder("C:\A\B\C")
Public Sub MkFolder(ByVal strPath As String)

    Dim arrPath() As String
    Dim strCur    As String
    Dim i         As Integer

    strPath = IIf(Right(strPath, 1) = "\", Left(strPath, Len(strPath) - 1), strPath)
    arrPath = Split(strPath, "\")

    For i = LBound(arrPath) To UBound(arrPath)
        strCur = strCur & arrPath(i) & "\"
        If Dir(strCur, vbDirectory) = "" Then MkDir strCur
    Next i

End Sub

' 목적   : 폴더 삭제 (내부 파일·하위 폴더 포함, 재귀)
' 인수   : strPath - 삭제할 폴더 경로
' 예시   : DelFolder("C:\TestFolder")
Public Sub DelFolder(ByVal strPath As String)
    If Dir(strPath, vbDirectory) = "" Then Exit Sub
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFolder strPath, True
    Set objFSO = Nothing
    On Error GoTo 0
End Sub

' 목적   : 파일 삭제
' 인수   : strPath - 삭제할 파일 경로
' 예시   : DelFile("C:\test.xlsx")
Public Sub DelFile(ByVal strPath As String)
    If Dir(strPath) <> "" Then
        On Error Resume Next
        Kill strPath
        On Error GoTo 0
    End If
End Sub

' 목적   : 파일 이름 변경 또는 이동
' 인수   : strOldPath - 원본 파일 경로
'          strNewPath - 변경할 파일 경로
' 예시   : RenFile("C:\old.xlsx", "C:\new.xlsx")
Public Sub RenFile(ByVal strOldPath As String, _
                   ByVal strNewPath As String)
    Dim strFolder As String
    strFolder = Left(strNewPath, InStrRev(strNewPath, "\") - 1)
    Call MkFolder(strFolder)
    Name strOldPath As strNewPath
End Sub

' 목적   : 파일 복사
' 인수   : strSrcPath  - 원본 파일 경로
'          strDestPath - 복사할 파일 경로
' 예시   : CopyFile("C:\src.xlsx", "D:\dest.xlsx")
Public Sub CopyFile(ByVal strSrcPath As String, _
                    ByVal strDestPath As String)
    Dim strFolder As String
    strFolder = Left(strDestPath, InStrRev(strDestPath, "\") - 1)
    Call MkFolder(strFolder)
    FileCopy strSrcPath, strDestPath
End Sub

' 목적   : 폴더 복사
' 인수   : strSrcPath  - 원본 폴더 경로
'          strDestPath - 복사할 폴더 경로
' 예시   : CopyFolder("C:\SrcFolder", "D:\DestFolder")
Public Sub CopyFolder(ByVal strSrcPath As String, _
                      ByVal strDestPath As String)
    Dim objFSO As Object
    Call MkFolder(strDestPath)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFolder strSrcPath, strDestPath
    Set objFSO = Nothing
End Sub

' ══════════════════════════════════════════════════════════
'  파일/폴더 검색
' ══════════════════════════════════════════════════════════

' 목적   : 하위 폴더 목록 수집 (재귀)
' 인수   : strPath       - 검색할 루트 경로
'          arrFolders    - 수집된 경로 배열 (ByRef)
'          blnSubFolders - True: 하위 폴더 재귀 검색
'          strFilter     - 폴더명 필터 (기본: *)
'          i             - 배열 인덱스 (내부 전달용)
' 예시   : GetFoldersList "C:\Root", arrFolders
Public Sub GetFoldersList(ByVal strPath As String, _
                          ByRef arrFolders() As Variant, _
                          Optional ByVal blnSubFolders As Boolean = True, _
                          Optional ByVal strFilter As String = "*", _
                          Optional ByRef i As Long = 0)

    Dim objFSO       As Object
    Dim objFolder    As Object
    Dim objSubFolder As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strPath)

    For Each objSubFolder In objFolder.SubFolders
        If objSubFolder.Name Like strFilter Then
            ReDim Preserve arrFolders(i)
            arrFolders(i) = objSubFolder.Path
            i = i + 1
        End If
        If blnSubFolders Then
            Call GetFoldersList(objSubFolder.Path, arrFolders, , strFilter, i)
        End If
    Next objSubFolder

    Set objFolder = Nothing
    Set objFSO = Nothing

End Sub

' 목적   : 파일 목록 수집
' 인수   : strPath       - 검색할 루트 경로
'          strFilter     - 파일명 필터 (기본: *)
'          blnSubFolders - True: 하위 폴더 파일도 검색
' 반환   : Variant - 파일 전체 경로 배열
' 예시   : arr = GetFilesList("C:\Root", "*.xlsx")
Public Function GetFilesList(ByVal strPath As String, _
                             Optional ByVal strFilter As String = "*", _
                             Optional ByVal blnSubFolders As Boolean = True) As Variant

    Dim arrFolders() As Variant
    Dim vntFolder    As Variant
    Dim arrFiles()   As Variant
    Dim objFSO       As Object
    Dim objFolder    As Object
    Dim objFile      As Object
    Dim i            As Long

    ReDim arrFolders(0)
    arrFolders(0) = strPath
    Call GetFoldersList(strPath, arrFolders, blnSubFolders, , 1)

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    For Each vntFolder In arrFolders
        Set objFolder = objFSO.GetFolder(vntFolder)
        For Each objFile In objFolder.Files
            If objFile.Path Like strFilter Then
                ReDim Preserve arrFiles(i)
                arrFiles(i) = objFile.Path
                i = i + 1
            End If
        Next objFile
    Next vntFolder

    Set objFolder = Nothing
    Set objFSO = Nothing

    GetFilesList = arrFiles

End Function

' ══════════════════════════════════════════════════════════
'  PRIVATE - 보조 함수
' ══════════════════════════════════════════════════════════

' 목적   : 경로가 CWB 폴더로 시작하면 해당 부분을 {cPath} 토큰으로 치환
'          (대소문자 무시 비교, 원본 대소문자 보존, 폴더 경계 확인)
' 예시   : prv_ToCwbToken("D:\Work\DB\a.accdb", "D:\Work") → "{cPath}\DB\a.accdb"
'          prv_ToCwbToken("D:\Work2\a.accdb", "D:\Work")  → "D:\Work2\a.accdb"
Private Function prv_ToCwbToken(ByVal strPath As String, _
                                ByVal strCwbPath As String) As String

    Dim lngLen As Long
    lngLen = Len(strCwbPath)
    prv_ToCwbToken = strPath

    If lngLen = 0 Then Exit Function
    If StrComp(Left(strPath, lngLen), strCwbPath, vbTextCompare) <> 0 Then Exit Function
    If Len(strPath) > lngLen Then
        If Mid(strPath, lngLen + 1, 1) <> "\" Then Exit Function
    End If

    prv_ToCwbToken = "{cPath}" & Mid(strPath, lngLen + 1)

End Function

' 목적   : 경로 끝의 구분자("\") 1개 제거
' 예시   : prv_TrimTrailingSep("D:\Work\") → "D:\Work"
'          prv_TrimTrailingSep("D:\")      → "D:"
Private Function prv_TrimTrailingSep(ByVal strPath As String) As String

    If Len(strPath) > 1 And Right(strPath, 1) = "\" Then
        strPath = Left(strPath, Len(strPath) - 1)
    End If

    prv_TrimTrailingSep = strPath

End Function
