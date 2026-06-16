Attribute VB_Name = "am_Sheet"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Sheet                                               │
' │  역할 : 워크북/시트 조작 (백업, 표시/숨김, 정렬, 보호) │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  워크북 / 시트 백업
' ══════════════════════════════════════════════════════════

' 목적   : 시트를 별도 파일로 백업
' 인수   : strPath - 백업 저장 폴더 경로
'          wb      - 호출한 CWB 워크북 개체 (기본: ActiveWorkbook)
'          strFile - 저장 파일명 (기본: 자동 생성)
'          ws      - 백업할 시트 (기본: ActiveSheet)
' 예시   : BackupSheet "C:\Backup"
'          BackupSheet "C:\Backup", ThisWorkbook
Public Sub BackupSheet(ByVal strPath As String, _
                       Optional ByVal wb As Workbook = Nothing, _
                       Optional ByVal strFile As String = "", _
                       Optional ByVal ws As Worksheet = Nothing)

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If ws Is Nothing Then Set ws = wb.ActiveSheet

    ws.Copy

    strPath = IIf(Right(strPath, 1) <> "\", strPath & "\", strPath)
    Call prv_MkFolder(strPath)

    If strFile = "" Then
        strFile = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & _
                  Left(wb.Name, InStrRev(wb.Name, ".") - 1) & "_" & ws.Name & ".xlsx"
    End If

    Application.DisplayAlerts = False
    Select Case prv_GetExt(strFile)
        Case ".xlsx": ActiveWorkbook.SaveAs strPath & strFile, 51
        Case ".xlsm": ActiveWorkbook.SaveAs strPath & strFile, 52
        Case ".csv":  ActiveWorkbook.SaveAs strPath & strFile, 6
        Case ".txt":  ActiveWorkbook.SaveAs strPath & strFile, 4158
    End Select
    Application.DisplayAlerts = True

    ActiveWorkbook.Close False

End Sub

' 목적   : 전체 워크북을 복사본으로 백업
' 인수   : strPath - 백업 저장 폴더 경로
'          wb      - 호출한 CWB 워크북 개체 (기본: ActiveWorkbook)
'          strFile - 저장 파일명 (기본: 자동 생성)
' 예시   : BackupWorkbook "C:\Backup"
'          BackupWorkbook "C:\Backup", ThisWorkbook
Public Sub BackupWorkbook(ByVal strPath As String, _
                          Optional ByVal wb As Workbook = Nothing, _
                          Optional ByVal strFile As String = "")

    If wb Is Nothing Then Set wb = ActiveWorkbook
    strPath = IIf(Right(strPath, 1) <> "\", strPath & "\", strPath)
    Call prv_MkFolder(strPath)

    If strFile = "" Then
        strFile = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & wb.Name
    End If

    wb.SaveCopyAs strPath & strFile

End Sub

' ══════════════════════════════════════════════════════════
'  시트 표시 / 숨김
' ══════════════════════════════════════════════════════════

' 목적   : 특정 시트만 표시하고 나머지 숨김
' 인수   : ws - 표시할 시트
'          wb - 대상 워크북 (기본: ws.Parent)
' 예시   : HideAllSheetsExceptOne Sheet1
'          HideAllSheetsExceptOne Sheet1, wbOther
Public Sub HideAllSheetsExceptOne(ByVal ws As Worksheet, _
                                  Optional ByVal wb As Workbook = Nothing)

    If wb Is Nothing Then Set wb = ws.Parent
    On Error Resume Next

    ws.Visible = True
    ws.Activate

    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
    End With

    Dim objSht As Object
    For Each objSht In wb.Sheets
        If objSht.Name <> ws.Name Then
            objSht.Visible = xlSheetHidden
        End If
    Next objSht

    On Error GoTo 0

End Sub

' 목적   : 워크북 내 모든 시트 표시
' 인수   : wb - 대상 워크북 (기본: ActiveWorkbook)
' 예시   : VisibleAllSheets
'          VisibleAllSheets wbOther
Public Sub VisibleAllSheets(Optional ByVal wb As Workbook = Nothing)

    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next

    Dim objSht As Object
    For Each objSht In wb.Sheets
        objSht.Visible = xlSheetVisible
    Next objSht

    On Error GoTo 0

End Sub

' ══════════════════════════════════════════════════════════
'  시트 정보 조회
' ══════════════════════════════════════════════════════════

' 목적   : 워크북 내 시트명 배열 반환
' 인수   : wb - 대상 워크북 (기본: ActiveWorkbook)
' 반환   : Variant - 시트명 1차원 배열 (0-based)
' 예시   : arr = GetSheetNames()
'          arr = GetSheetNames(wbOther)
Public Function GetSheetNames(Optional ByVal wb As Workbook = Nothing) As Variant

    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim arrResult() As Variant
    Dim objSht      As Object
    Dim i           As Long

    For Each objSht In wb.Sheets
        ReDim Preserve arrResult(i)
        arrResult(i) = objSht.Name
        i = i + 1
    Next objSht

    GetSheetNames = arrResult

End Function

' ══════════════════════════════════════════════════════════
'  시트 순서 정렬
' ══════════════════════════════════════════════════════════

' 목적   : 배열 순서대로 시트 탭 순서 재정렬
' 인수   : arrSheetNames - 정렬 순서 시트명 배열
'          wb            - 대상 워크북 (기본: ActiveWorkbook)
' 예시   : SortSheets Array("Home", "입력", "DB")
'          SortSheets Array("Home", "입력", "DB"), wbOther
Public Sub SortSheets(ByVal arrSheetNames As Variant, _
                      Optional ByVal wb As Workbook = Nothing)

    If wb Is Nothing Then Set wb = ActiveWorkbook
    Call VisibleAllSheets(wb)

    Dim vntName As Variant
    Dim objSht  As Object

    For Each vntName In arrSheetNames
        On Error Resume Next
        Set objSht = wb.Sheets(CStr(vntName))
        If Not objSht Is Nothing Then
            objSht.Move After:=wb.Sheets(wb.Sheets.Count)
        End If
        Set objSht = Nothing
        On Error GoTo 0
    Next vntName

End Sub

' ══════════════════════════════════════════════════════════
'  시트 보호
' ══════════════════════════════════════════════════════════

' 목적   : 시트 보호 설정
'          서식 포함 사용 범위(ws.UsedRange) 내 배경색 없는 셀만 잠금 해제
'          → 배경색 없는 셀 = 입력 가능 셀 규약
' 인수   : strPW - 보호 비밀번호
'          ws    - 보호할 시트 (기본: ActiveSheet)
' 예시   : SheetLock "1234"
'          SheetLock "1234", Sheet2
Public Sub SheetLock(ByVal strPW As String, _
                     Optional ByVal ws As Worksheet = Nothing)

    If ws Is Nothing Then Set ws = ActiveSheet

    Dim blnScreen  As Boolean
    Dim rngUsed    As Range
    Dim rngUnlock  As Range
    Dim rngBlanks  As Range
    Dim cel        As Range
    Dim f          As Range
    Dim strFirst   As String

    blnScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    On Error GoTo ErrHandler

    ws.Unprotect Password:=strPW

    ' 1. 시트 전체 잠금 (기본 상태: 모두 잠김)
    ws.Cells.Locked        = True
    ws.Cells.FormulaHidden = True

    Set rngUsed = ws.UsedRange

    ' 2a. 내용 있는 셀 — SearchFormat 으로 배경 없는 셀 일괄 수집 (네이티브, 빠름)
    Application.FindFormat.Clear
    Application.FindFormat.Interior.Pattern = xlNone

    Set f = rngUsed.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                         SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                         SearchFormat:=True)
    If Not f Is Nothing Then
        strFirst      = f.Address
        Set rngUnlock = f
        Do
            Set f = rngUsed.FindNext(f)
            If f Is Nothing         Then Exit Do
            If f.Address = strFirst Then Exit Do
            Set rngUnlock = Union(rngUnlock, f)
        Loop
    End If

    Application.FindFormat.Clear

    ' 2b. 빈 셀 — SpecialCells 후 Pattern 체크 (빈 셀만 순회)
    On Error Resume Next
    Set rngBlanks = rngUsed.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If Not rngBlanks Is Nothing Then
        For Each cel In rngBlanks.Cells
            If cel.Interior.Pattern = xlNone Then
                If rngUnlock Is Nothing Then
                    Set rngUnlock = cel
                Else
                    Set rngUnlock = Union(rngUnlock, cel)
                End If
            End If
        Next cel
    End If

    ' 3. 배경 없는 셀 일괄 잠금 해제 (벌크 연산)
    If Not rngUnlock Is Nothing Then
        rngUnlock.Locked        = False
        rngUnlock.FormulaHidden = False
    End If

    With ws
        .Protect Password:=strPW, Contents:=True, Scenarios:=True, _
                 AllowSorting:=True, AllowFiltering:=True
        .EnableSelection = xlNoRestrictions
    End With

    GoTo CleanUp

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, "am_Sheet.SheetLock"

CleanUp:
    Application.FindFormat.Clear
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = blnScreen

End Sub

' 목적   : 시트 보호 해제
' 인수   : strPW - 보호 비밀번호
'          ws    - 해제할 시트 (기본: ActiveSheet)
' 예시   : SheetUnLock "1234"
'          SheetUnLock "1234", Sheet2
Public Sub SheetUnLock(ByVal strPW As String, _
                       Optional ByVal ws As Worksheet = Nothing)

    If ws Is Nothing Then Set ws = ActiveSheet
    ws.Unprotect Password:=strPW

End Sub

' ══════════════════════════════════════════════════════════
'  내부 전용 함수 (Private)
' ══════════════════════════════════════════════════════════

' 목적   : 폴더 생성 (내부 전용)
Private Sub prv_MkFolder(ByVal strPath As String)

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

' 목적   : 파일 확장자 반환 (내부 전용)
Private Function prv_GetExt(ByVal strFileName As String) As String
    prv_GetExt = Mid(strFileName, InStrRev(strFileName, "."))
End Function


