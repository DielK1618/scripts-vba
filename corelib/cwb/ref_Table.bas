Attribute VB_Name = "ref_Table"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Table                                              │
' │  역할 : am_Table Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

' 참고 : 아래 프로시저는 래핑 불가 → 직접 Application.Run 사용
'        - TblFindVals_MC   : ParamArray 파라미터
'        - TblFindVal_One   : ParamArray 파라미터
'        - TblFindRng_MC    : ParamArray 파라미터 + Range 반환

Private Const REF As String = "corelib.xlam!am_Table."

' ── 테이블 조회 ───────────────────────────────────────────────

' 목적   : 워크북 내 모든 시트의 ListObject 이름 배열 반환
' 인수   : wb - 대상 워크북 (기본: ActiveWorkbook)
' 반환   : Variant - 테이블명 1차원 배열
' 예시   : arr = GetAllSheetTableNames()
'          arr = GetAllSheetTableNames(wbOther)
Public Function GetAllSheetTableNames(Optional ByVal wb As Workbook = Nothing) As Variant
    If wb Is Nothing Then Set wb = ActiveWorkbook
    GetAllSheetTableNames = Application.Run(REF & "GetAllSheetTableNames", wb)
End Function

' 목적   : 시트 내 ListObject 이름 배열 반환 (특정 테이블 제외 가능)
' 인수   : strPassName - 제외할 테이블명 (기본: 없음)
'          ws          - 대상 시트 (기본: ActiveSheet)
' 반환   : Variant - 테이블명 1차원 배열
' 예시   : GetTableNames()
'          GetTableNames "tbl_임시", Sheet1
Public Function GetTableNames(Optional ByVal strPassName As String = "", _
                              Optional ByVal ws As Worksheet = Nothing) As Variant
    If ws Is Nothing Then Set ws = ActiveSheet
    GetTableNames = Application.Run(REF & "GetTableNames", strPassName, ws)
End Function

' 목적   : 테이블 각 열 너비 배열 반환
' 인수   : tbl          - 대상 ListObject (기본: ActiveCell 포함 테이블)
'          intRound     - 반올림 자릿수 (기본: 0 = 정수)
'          sngMultiplier - 결과에 곱할 배수 (기본: 1)
' 반환   : Variant - 열 너비 배열
' 예시   : GetTableColumnsWidth(Sheet1.ListObjects("tbl_Data"), 1, 2.5)
Public Function GetTableColumnsWidth(Optional ByVal tbl As ListObject, _
                                     Optional ByVal intRound As Integer = 0, _
                                     Optional ByVal sngMultiplier As Single = 1) As Variant
    GetTableColumnsWidth = Application.Run(REF & "GetTableColumnsWidth", tbl, intRound, sngMultiplier)
End Function

' 목적   : 테이블 크기를 데이터에 맞게 자동 조정
' 인수   : tbl - 조정할 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : ResizeTable(Sheet1.ListObjects("tbl_Data"))
Public Sub ResizeTable(Optional ByVal tbl As ListObject)
    Application.Run REF & "ResizeTable", tbl
End Sub

' 목적   : Range 가 ListObject 에 속하는지 확인
' 인수   : rng - 확인할 범위
' 반환   : Boolean - True: 테이블 내부 / False: 테이블 외부
' 예시   : IsTable(ActiveCell) → True
Public Function IsTable(ByVal rng As Range) As Boolean
    IsTable = Application.Run(REF & "IsTable", rng)
End Function

' ── 테이블 행 / 열 조작 ──────────────────────────────────────

' 목적   : 테이블의 모든 자동 필터 해제
' 인수   : tbl - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : ClearFiltersInTable(Sheet1.ListObjects("tbl_Data"))
Public Sub ClearFiltersInTable(Optional ByVal tbl As ListObject)
    Application.Run REF & "ClearFiltersInTable", tbl
End Sub

' 목적   : 테이블에 행 추가
' 인수   : intAddRows    - 추가할 행 수
'          intSelectRow  - 삽입 위치 행 번호 (0: 마지막 행 아래)
'          tbl           - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : AddTableRows 3, 0, Sheet1.ListObjects("tbl_Data")
Public Sub AddTableRows(ByVal intAddRows As Long, _
                        Optional ByVal intSelectRow As Long = 0, _
                        Optional ByVal tbl As ListObject)
    Application.Run REF & "AddTableRows", intAddRows, intSelectRow, tbl
End Sub

' 목적   : 테이블에 열 추가
' 인수   : intAddColumns    - 추가할 열 수
'          intSelectColumn  - 삽입 위치 열 번호 (0: 마지막 열 오른쪽)
'          tbl              - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : AddColumns 2, 3, Sheet1.ListObjects("tbl_Data")
Public Sub AddColumns(ByVal intAddColumns As Long, _
                      Optional ByVal intSelectColumn As Long = 0, _
                      Optional ByVal tbl As ListObject)
    Application.Run REF & "AddColumns", intAddColumns, intSelectColumn, tbl
End Sub

' 목적   : 열 이름 배열로 테이블에 여러 열 한 번에 추가
' 인수   : arrColumnNames   - 추가할 열 이름 배열
'          intSelectColumn  - 삽입 위치 열 번호 (0: 마지막)
'          tbl              - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : AddArrayColumns Array("비고1","비고2"), 0, Sheet1.ListObjects("tbl_Data")
Public Sub AddArrayColumns(ByVal arrColumnNames As Variant, _
                           Optional ByVal intSelectColumn As Long = 0, _
                           Optional ByVal tbl As ListObject)
    Application.Run REF & "AddArrayColumns", arrColumnNames, intSelectColumn, tbl
End Sub

' 목적   : 테이블 열 삭제
' 인수   : intSelectColumn - 삭제할 열 번호 (1-based)
'          intCount        - 삭제할 열 수 (기본: 1)
'          tbl             - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : DelTableColumns 3, 2, Sheet1.ListObjects("tbl_Data")
Public Sub DelTableColumns(ByVal intSelectColumn As Long, _
                           Optional ByVal intCount As Long = 1, _
                           Optional ByVal tbl As ListObject)
    Application.Run REF & "DelTableColumns", intSelectColumn, intCount, tbl
End Sub

' 목적   : 테이블 행 삭제
' 인수   : intSelectRow - 삭제할 행 번호 (1-based)
'          intCount     - 삭제할 행 수 (기본: 1)
'          tbl          - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : DelTableRows 5, 1, Sheet1.ListObjects("tbl_Data")
Public Sub DelTableRows(ByVal intSelectRow As Long, _
                        Optional ByVal intCount As Long = 1, _
                        Optional ByVal tbl As ListObject)
    Application.Run REF & "DelTableRows", intSelectRow, intCount, tbl
End Sub

' 목적   : 테이블 데이터 행 전체 삭제 (헤더 유지)
' 인수   : tbl - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 예시   : DelTableAllRows Sheet1.ListObjects("tbl_Data")
Public Sub DelTableAllRows(Optional ByVal tbl As ListObject)
    Application.Run REF & "DelTableAllRows", tbl
End Sub

' 목적   : 필터 조건에 맞는 행 삭제 후 삭제 건수 반환
' 인수   : strFieldName     - 필터 기준 열 이름
'          strFilterPattern - 필터 패턴 (Like 연산자 패턴)
'          tbl              - 대상 ListObject (기본: ActiveCell 포함 테이블)
' 반환   : Long - 삭제된 행 수
' 예시   : DelTableFilteredRows("상태", "삭제*", Sheet1.ListObjects("tbl_Data")) → 5
Public Function DelTableFilteredRows(ByVal strFieldName As String, _
                                     ByVal strFilterPattern As String, _
                                     Optional ByVal tbl As ListObject) As Long
    DelTableFilteredRows = Application.Run(REF & "DelTableFilteredRows", _
                                           strFieldName, strFilterPattern, tbl)
End Function

' ── 테이블 필터 ───────────────────────────────────────────────

' 목적   : 테이블 특정 열에 단일 값 자동 필터 적용
' 인수   : fieldName - 필터 기준 열 이름 또는 번호
'          strValue  - 필터 값 ("" 또는 "*": 전체 해제)
'          tbl       - 대상 ListObject (기본: ActiveSheet 첫번째 테이블)
' 예시   : AutoTableFilter "상태", "완료"
'          AutoTableFilter "상태", "완료", tbl
Public Sub AutoTableFilter(ByVal fieldName As Variant, _
                           ByVal strValue As String, _
                           Optional ByVal tbl As ListObject = Nothing)
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    Application.Run REF & "AutoTableFilter", fieldName, strValue, tbl
End Sub

' 목적   : 테이블 특정 열에 다중 값 배열로 자동 필터 적용
' 인수   : fieldName - 필터 기준 열 이름
'          arrValues - 필터 값 배열
'          blnPart   - True: 부분 일치 / False: 완전 일치 (기본: False)
'          tbl       - 대상 ListObject (기본: ActiveSheet 첫번째 테이블)
' 예시   : AutoTableFilter_Arr "상태", Array("완료","대기")
'          AutoTableFilter_Arr "상태", Array("완료","대기"), , tbl
Public Sub AutoTableFilter_Arr(ByVal fieldName As String, _
                               ByVal arrValues As Variant, _
                               Optional ByVal blnPart As Boolean = False, _
                               Optional ByVal tbl As ListObject = Nothing)
    If tbl Is Nothing Then Set tbl = ActiveSheet.ListObjects(1)
    Application.Run REF & "AutoTableFilter_Arr", fieldName, arrValues, blnPart, tbl
End Sub

' ── 테이블 정렬 ───────────────────────────────────────────────

' 목적   : 테이블을 특정 열 기준으로 정렬
' 인수   : strTableName   - ListObject 이름
'          strFieldName   - 정렬 기준 열 이름
'          xlOrder        - 정렬 방향 (기본: xlAscending 오름차순)
'          xlOrientation  - 정렬 방향 축 (기본: xlSortColumns 열 기준)
'          xlHeader       - 헤더 포함 여부 (기본: xlYes)
' 예시   : SortTable "tbl_Data", "이름"
Public Sub SortTable(ByVal strTableName As String, _
                     ByVal strFieldName As String, _
                     Optional ByVal xlOrder As XlSortOrder = xlAscending, _
                     Optional ByVal xlOrientation As XlSortOrientation = xlSortColumns, _
                     Optional ByVal xlHeader As XlYesNoGuess = xlYes)
    Application.Run REF & "SortTable", strTableName, strFieldName, xlOrder, xlOrientation, xlHeader
End Sub

' 목적   : 테이블을 사용자 정의 목록 순서로 정렬
' 인수   : strTableName - ListObject 이름
'          strFieldName - 정렬 기준 열 이름
'          customList   - 정렬 순서 배열 (예: Array("높음","중간","낮음"))
'          xlOrder      - 정렬 방향 (기본: xlAscending)
' 예시   : SortTableCustomList "tbl_Data", "우선순위", Array("높음","중간","낮음")
Public Sub SortTableCustomList(ByVal strTableName As String, _
                               ByVal strFieldName As String, _
                               ByVal customList As Variant, _
                               Optional ByVal xlOrder As XlSortOrder = xlAscending)
    Application.Run REF & "SortTableCustomList", strTableName, strFieldName, customList, xlOrder
End Sub

' ── 테이블 값 조작 ───────────────────────────────────────────

' 목적   : 특정 ID 행의 지정 열 값 변경
' 인수   : strID        - 검색할 ID 값
'          strFieldName - 변경할 열 이름
'          strValue     - 변경할 값
'          tbl          - 대상 ListObject (기본: ActiveCell 포함 테이블)
'          strIDField   - ID 열 이름 (기본: "ID")
' 예시   : ChangeTableValue "U001", "상태", "완료", Sheet1.ListObjects("tbl_Data")
Public Sub ChangeTableValue(ByVal strID As String, _
                            ByVal strFieldName As String, _
                            ByVal strValue As String, _
                            Optional ByVal tbl As ListObject, _
                            Optional ByVal strIDField As String = "ID")
    Application.Run REF & "ChangeTableValue", strID, strFieldName, strValue, tbl, strIDField
End Sub

' 목적   : 기준 셀에서 목표 열까지의 열 오프셋 반환
' 인수   : CelFrom    - 기준 셀
'          strTgField - 목표 열 이름
'          blnTable   - True: 테이블 헤더 기준 / False: 시트 열 기준
' 반환   : Integer - 열 오프셋 값
' 예시   : intOffset(ActiveCell, "상태") → 3 (3칸 오른쪽)
Public Function intOffset(ByVal CelFrom As Range, _
                          ByVal strTgField As String, _
                          Optional ByVal blnTable As Boolean = True) As Integer
    intOffset = CInt(Application.Run(REF & "intOffset", CelFrom, strTgField, blnTable))
End Function
