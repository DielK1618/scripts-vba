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
'          wb      - 호출한 CWB 워크북 개체
'          strFile - 저장 파일명 (기본: 자동 생성)
'          ws      - 백업할 시트 (기본: ActiveSheet)
' 예시   : BackupSheet("C:\Backup", ThisWorkbook)
Public Sub BackupSheet(ByVal strPath As String, _
                       ByVal wb As Workbook, _
                       Optional ByVal strFile As String = "", _
                       Optional ByVal ws As Worksheet)
    Application.Run REF & "BackupSheet", strPath, wb, strFile, ws
End Sub

' 목적   : 전체 워크북을 복사본으로 백업
' 인수   : strPath - 백업 저장 폴더 경로
'          wb      - 호출한 CWB 워크북 개체
'          strFile - 저장 파일명 (기본: 자동 생성)
' 예시   : BackupWorkbook("C:\Backup", ThisWorkbook)
Public Sub BackupWorkbook(ByVal strPath As String, _
                          ByVal wb As Workbook, _
                          Optional ByVal strFile As String = "")
    Application.Run REF & "BackupWorkbook", strPath, wb, strFile
End Sub

' ── 시트 표시 / 숨김 ─────────────────────────────────────────

' 목적   : 특정 시트만 표시하고 나머지 숨김
' 인수   : ws - 표시할 시트
'          wb - 대상 워크북
' 예시   : HideAllSheetsExceptOne(Sheet1, ThisWorkbook)
Public Sub HideAllSheetsExceptOne(ByVal ws As Worksheet, ByVal wb As Workbook)
    Application.Run REF & "HideAllSheetsExceptOne", ws, wb
End Sub

' 목적   : 워크북 내 모든 시트 표시
' 인수   : wb - 대상 워크북
' 예시   : VisibleAllSheets(ThisWorkbook)
Public Sub VisibleAllSheets(ByVal wb As Workbook)
    Application.Run REF & "VisibleAllSheets", wb
End Sub

' ── 시트 정보 조회 ───────────────────────────────────────────

' 목적   : 워크북 내 시트명 배열 반환
' 인수   : wb - 대상 워크북
' 반환   : Variant - 시트명 1차원 배열 (0-based)
' 예시   : arr = GetSheetNames(ThisWorkbook)
Public Function GetSheetNames(ByVal wb As Workbook) As Variant
    GetSheetNames = Application.Run(REF & "GetSheetNames", wb)
End Function

' ── 시트 순서 정렬 ───────────────────────────────────────────

' 목적   : 배열 순서대로 시트 탭 순서 재정렬
' 인수   : wb            - 대상 워크북
'          arrSheetNames - 정렬 순서 시트명 배열
' 예시   : SortSheets(ThisWorkbook, Array("Home", "입력", "DB"))
Public Sub SortSheets(ByVal wb As Workbook, ByVal arrSheetNames As Variant)
    Application.Run REF & "SortSheets", wb, arrSheetNames
End Sub

' ── 시트 보호 ────────────────────────────────────────────────

' 목적   : 시트 보호 설정
'          실제 사용 범위 전체 잠금 후, 배경색 없는 셀만 잠금 해제
'          → 배경색 없는 셀 = 입력 가능 셀 규약
' 인수   : ws    - 보호할 시트
'          strPW - 보호 비밀번호
' 예시   : SheetLock(ActiveSheet, "1234")
Public Sub SheetLock(ByVal ws As Worksheet, ByVal strPW As String)
    Application.Run REF & "SheetLock", ws, strPW
End Sub

' 목적   : 시트 보호 해제
' 인수   : ws    - 해제할 시트
'          strPW - 보호 비밀번호
' 예시   : SheetUnLock(ActiveSheet, "1234")
Public Sub SheetUnLock(ByVal ws As Worksheet, ByVal strPW As String)
    Application.Run REF & "SheetUnLock", ws, strPW
End Sub
