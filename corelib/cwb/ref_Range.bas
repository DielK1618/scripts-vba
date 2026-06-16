Attribute VB_Name = "ref_Range"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Range                                              │
' │  역할 : am_Range Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

' 참고 : Range 반환 함수는 Application.Run 경유 수신 불가
'        - FindRange         → FindRange_IsValid / FindRange_CellValue 사용
'        - FindCellsByColor  → FindCellsByColor_Count 사용
'        - GetUsedRange      → GetUsedRange_IsValid / GetUsedRange_RowCount / GetUsedRange_ColCount 사용

Private Const REF As String = "corelib.xlam!am_Range."

' ── GetUsedRange 스칼라 래퍼 ─────────────────────────────────

' 목적   : 시트에 실제 사용 범위가 있는지 확인
' 인수   : ws - 대상 시트 (기본: ActiveSheet)
' 반환   : Boolean - True: 범위 있음 / False: 빈 시트
' 예시   : If GetUsedRange_IsValid(ActiveSheet) Then ...
Public Function GetUsedRange_IsValid(Optional ByVal ws As Worksheet) As Boolean
    GetUsedRange_IsValid = Application.Run(REF & "GetUsedRange_IsValid", ws)
End Function

' 목적   : 시트의 실제 사용 범위 행 수 반환
' 인수   : ws - 대상 시트 (기본: ActiveSheet)
' 반환   : Long - 행 수 (범위 없으면 0)
' 예시   : GetUsedRange_RowCount(ActiveSheet) → 100
Public Function GetUsedRange_RowCount(Optional ByVal ws As Worksheet) As Long
    GetUsedRange_RowCount = Application.Run(REF & "GetUsedRange_RowCount", ws)
End Function

' 목적   : 시트의 실제 사용 범위 열 수 반환
' 인수   : ws - 대상 시트 (기본: ActiveSheet)
' 반환   : Long - 열 수 (범위 없으면 0)
' 예시   : GetUsedRange_ColCount(ActiveSheet) → 10
Public Function GetUsedRange_ColCount(Optional ByVal ws As Worksheet) As Long
    GetUsedRange_ColCount = Application.Run(REF & "GetUsedRange_ColCount", ws)
End Function

' ── FindRange 스칼라 래퍼 ────────────────────────────────────

' 목적   : 범위 내 값 검색 결과가 존재하는지 확인
' 인수   : rngFind   - 검색 범위
'          strValue  - 검색 값
'          blnAddRow - True: 미발견 시 다음 빈 행도 유효로 처리
' 반환   : Boolean - True: 찾음 / False: 없음
' 예시   : FindRange_IsValid(tbl.ListColumns("ID").DataBodyRange, "001") → True
Public Function FindRange_IsValid(ByVal rngFind As Range, _
                                  ByVal strValue As String, _
                                  Optional ByVal blnAddRow As Boolean = False) As Boolean
    FindRange_IsValid = Application.Run(REF & "FindRange_IsValid", rngFind, strValue, blnAddRow)
End Function

' 목적   : 범위 내 값 검색 후 찾은 셀 값 반환
' 인수   : rngFind   - 검색 범위
'          strValue  - 검색 값
'          blnAddRow - True: 미발견 시 다음 빈 행 셀 값 반환
' 반환   : String - 찾은 셀 값 ("" = 없음)
' 예시   : FindRange_CellValue(tbl.ListColumns("이름").DataBodyRange, "홍길동")
Public Function FindRange_CellValue(ByVal rngFind As Range, _
                                    ByVal strValue As String, _
                                    Optional ByVal blnAddRow As Boolean = False) As String
    FindRange_CellValue = Application.Run(REF & "FindRange_CellValue", rngFind, strValue, blnAddRow)
End Function

' ── FindCellsByColor 스칼라 래퍼 ─────────────────────────────

' 목적   : 배경색 기준으로 범위 내 일치 셀 수 반환
' 인수   : lngColor  - 검색할 색상값
'          rng       - 검색 범위
'          blnUnion  - True: ColorIndex 기준 / False: Color(RGB) 기준
' 반환   : Long - 조건에 맞는 셀 수 (없으면 0)
' 예시   : FindCellsByColor_Count(xlNone, ws.UsedRange, True) → 50
Public Function FindCellsByColor_Count(ByVal lngColor As Long, _
                                       ByVal rng As Range, _
                                       Optional ByVal blnUnion As Boolean = True) As Long
    FindCellsByColor_Count = Application.Run(REF & "FindCellsByColor_Count", lngColor, rng, blnUnion)
End Function
