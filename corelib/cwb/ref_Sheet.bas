Attribute VB_Name = "ref_Sheet"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Sheet                                              │
' │  역할 : am_Sheet Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Sheet."

' ── 백업 ─────────────────────────────────────────────────────

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
    Application.Run REF & "BackupSheet", strPath, wb, strFile, ws
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
    Application.Run REF & "BackupWorkbook", strPath, wb, strFile
End Sub

' ── 시트 표시 / 숨김 ─────────────────────────────────────────

' 목적   : 특정 시트만 표시하고 나머지 숨김
' 인수   : ws - 표시할 시트
'          wb - 대상 워크북 (기본: ws.Parent)
' 예시   : HideAllSheetsExceptOne Sheet1
'          HideAllSheetsExceptOne Sheet1, wbOther
Public Sub HideAllSheetsExceptOne(ByVal ws As Worksheet, _
                                  Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ws.Parent
    Application.Run REF & "HideAllSheetsExceptOne", ws, wb
End Sub

' 목적   : 워크북 내 모든 시트 표시
' 인수   : wb - 대상 워크북 (기본: ActiveWorkbook)
' 예시   : VisibleAllSheets
'          VisibleAllSheets wbOther
Public Sub VisibleAllSheets(Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Application.Run REF & "VisibleAllSheets", wb
End Sub

' ── 시트 정보 조회 ───────────────────────────────────────────

' 목적   : 워크북 내 시트명 배열 반환
' 인수   : wb - 대상 워크북 (기본: ActiveWorkbook)
' 반환   : Variant - 시트명 1차원 배열 (0-based)
' 예시   : arr = GetSheetNames()
'          arr = GetSheetNames(wbOther)
Public Function GetSheetNames(Optional ByVal wb As Workbook = Nothing) As Variant
    If wb Is Nothing Then Set wb = ActiveWorkbook
    GetSheetNames = Application.Run(REF & "GetSheetNames", wb)
End Function

' ── 시트 순서 정렬 ───────────────────────────────────────────

' 목적   : 배열 순서대로 시트 탭 순서 재정렬
' 인수   : arrSheetNames - 정렬 순서 시트명 배열
'          wb            - 대상 워크북 (기본: ActiveWorkbook)
' 예시   : SortSheets Array("Home", "입력", "DB")
'          SortSheets Array("Home", "입력", "DB"), wbOther
Public Sub SortSheets(ByVal arrSheetNames As Variant, _
                      Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Application.Run REF & "SortSheets", arrSheetNames, wb
End Sub

' ── 시트 보호 ────────────────────────────────────────────────

' 목적   : 시트 보호 설정
'          서식 포함 사용 범위(ws.UsedRange) 내 배경색 없는 셀만 잠금 해제
'          → 배경색 없는 셀 = 입력 가능 셀 규약
' 인수   : strPW - 보호 비밀번호 (기본: LOCK_PW)
'          ws    - 보호할 시트 (기본: ActiveSheet)
' 예시   : SheetLock
'          SheetLock "pw", Sheet2
Public Sub SheetLock(Optional ByVal strPW As String = LOCK_PW, _
                     Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ActiveSheet
    Application.Run REF & "SheetLock", strPW, ws
End Sub

' 목적   : 시트 보호 해제
' 인수   : strPW - 보호 비밀번호 (기본: LOCK_PW)
'          ws    - 해제할 시트 (기본: ActiveSheet)
' 예시   : SheetUnLock
'          SheetUnLock "pw", Sheet2
Public Sub SheetUnLock(Optional ByVal strPW As String = LOCK_PW, _
                       Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ActiveSheet
    Application.Run REF & "SheetUnLock", strPW, ws
End Sub
