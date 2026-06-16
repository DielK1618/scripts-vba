Attribute VB_Name = "ref_Utils"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Utils                                              │
' │  역할 : am_Utils Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

' 참고 : 아래 프로시저는 래핑 불가 → 직접 Application.Run 사용
'        - IsTableRange   : ByRef Object 파라미터 (tbl, celTblHeader)
'        - CreateUniqueID : ByRef Collection 파라미터

Private Const REF As String = "corelib.xlam!am_Utils."

' ── 배열 ─────────────────────────────────────────────────────

' 목적   : Variant 배열을 동일 값의 새 배열로 복사하여 반환
'          (VBA 배열 대입 시 참조 공유 방지용)
' 인수   : arr - 복사할 원본 배열
' 반환   : Variant - 값이 복사된 새 배열
' 예시   : arrCopy = ConvertToArrData(arrSource)
Public Function ConvertToArrData(ByVal arr As Variant) As Variant
    ConvertToArrData = Application.Run(REF & "ConvertToArrData", arr)
End Function

' ── 검사 ─────────────────────────────────────────────────────

' 목적   : 배열이 비어 있는지 확인 (초기화 안 된 배열 포함)
' 인수   : arr - 확인할 배열 (Variant)
' 반환   : Boolean - True: 빈 배열 / False: 데이터 있음
' 예시   : If IsArrayEmpty(arrData) Then Exit Sub
Public Function IsArrayEmpty(ByVal arr As Variant) As Boolean
    IsArrayEmpty = Application.Run(REF & "IsArrayEmpty", arr)
End Function

' 목적   : 개체가 Range(Cells) 타입인지 확인
' 인수   : obj - 확인할 개체
' 반환   : Boolean - True: Range / False: 그 외
' 예시   : If IsCells(Selection) Then ...
Public Function IsCells(ByVal obj As Object) As Boolean
    IsCells = Application.Run(REF & "IsCells", obj)
End Function

' 목적   : Range 내 병합 셀이 있는지 확인
' 인수   : rng - 확인할 범위
' 반환   : Boolean - True: 병합 셀 포함 / False: 없음
' 예시   : If IsRangeMerged(Sheet1.Range("A1:D1")) Then ...
Public Function IsRangeMerged(ByVal rng As Range) As Boolean
    IsRangeMerged = Application.Run(REF & "IsRangeMerged", rng)
End Function

' 목적   : 파일명으로 사용할 수 없는 문자(\/:*?"<>|) 포함 여부 확인
' 인수   : strFileName - 검사할 파일명 문자열
' 반환   : Boolean - True: 유효한 파일명 / False: 사용 불가 문자 포함
' 예시   : IsValidFileName("report_2024.xlsx") → True
'          IsValidFileName("re:port") → False
Public Function IsValidFileName(ByVal strFileName As String) As Boolean
    IsValidFileName = Application.Run(REF & "IsValidFileName", strFileName)
End Function

' 목적   : 셀에 적용된 유효성 검사 유형 이름 반환
' 인수   : cel - 확인할 셀
' 반환   : String - 유효성 유형명 (예: "목록", "정수", "날짜", "없음")
' 예시   : GetValidationType(Sheet1.Range("B2")) → "목록"
Public Function GetValidationType(ByVal cel As Range) As String
    GetValidationType = Application.Run(REF & "GetValidationType", cel)
End Function

' ── 코드 생성 ────────────────────────────────────────────────

' 목적   : 기존 코드 목록과 중복되지 않는 랜덤 코드 생성
' 인수   : rngCodes - 기존 코드가 있는 범위 (중복 방지 기준)
'          intLen   - 생성할 코드 길이
' 반환   : String - 중복 없는 랜덤 영숫자 코드
' 예시   : GenerateRandomCode(tbl.ListColumns("코드").DataBodyRange, 8) → "X7K2P9MQ"
Public Function GenerateRandomCode(ByVal rngCodes As Range, _
                                   ByVal intLen As Integer) As String
    GenerateRandomCode = Application.Run(REF & "GenerateRandomCode", rngCodes, intLen)
End Function

' ── 날짜 / 정규식 ────────────────────────────────────────────

' 목적   : 날짜/시각 문자열을 Excel 날짜 일련번호로 변환
' 인수   : strDate - 날짜 문자열 (예: "2024-01-15")
'          strTime - 시각 문자열 (기본: "00:00:00")
' 반환   : Double - Excel 날짜 일련번호 (예: 45306.0)
' 예시   : ConvertToExcelSerialDate("2024-01-15", "09:30:00") → 45306.395833...
Public Function ConvertToExcelSerialDate(ByVal strDate As String, _
                                         Optional ByVal strTime As String = "00:00:00") As Double
    ConvertToExcelSerialDate = Application.Run(REF & "ConvertToExcelSerialDate", strDate, strTime)
End Function

' 목적   : 정규식 패턴으로 문자열에서 값 추출
' 인수   : strValue   - 원본 문자열
'          strPattern - VBScript 정규식 패턴
' 반환   : Variant - 일치한 값 배열 (없으면 빈 배열)
' 예시   : ExtractValues("ABC-123-XYZ", "\d+") → Array("123")
Public Function ExtractValues(ByVal strValue As String, _
                              ByVal strPattern As String) As Variant
    ExtractValues = Application.Run(REF & "ExtractValues", strValue, strPattern)
End Function

' ── 외부 앱 ──────────────────────────────────────────────────

' 목적   : 주소 문자열을 Google Maps 에서 열기 (기본 브라우저)
' 인수   : strAddress - 검색할 주소 문자열
' 예시   : OpenAddressInGoogleMaps("서울시 강남구 테헤란로 123")
Public Sub OpenAddressInGoogleMaps(ByVal strAddress As String)
    Application.Run REF & "OpenAddressInGoogleMaps", strAddress
End Sub

' 목적   : 동영상 파일 재생 시간 반환
' 인수   : strPath - 동영상 파일 경로
' 반환   : String - 재생 시간 문자열 (예: "01:23:45")
' 예시   : GetVideoLength("C:\Videos\intro.mp4") → "00:02:30"
Public Function GetVideoLength(ByVal strPath As String) As String
    GetVideoLength = Application.Run(REF & "GetVideoLength", strPath)
End Function

' ── 도구 ─────────────────────────────────────────────────────

' 목적   : 현재 Selection 의 타입 정보를 메시지박스로 표시 (디버그용)
' 예시   : CheckSelectionType
Public Sub CheckSelectionType()
    Application.Run REF & "CheckSelectionType"
End Sub

' 목적   : 밀리초 단위 대기
' 인수   : lngMs - 대기 시간 (밀리초)
' 예시   : WaitMs 500   ' 0.5초 대기
Public Sub WaitMs(ByVal lngMs As Long)
    Application.Run REF & "WaitMs", lngMs
End Sub

' ── 수식 / 유효성 ────────────────────────────────────────────

' 목적   : 수식 문자열을 평가하여 True/False 반환
' 인수   : strFormula - 평가할 수식 문자열 (예: "=A1>0")
' 반환   : Boolean - 수식 평가 결과
' 예시   : EvaluateFormula("=A1>B1") → True
Public Function EvaluateFormula(ByVal strFormula As String) As Boolean
    EvaluateFormula = Application.Run(REF & "EvaluateFormula", strFormula)
End Function

' 목적   : 셀에 유효성 검사를 통과하는 값만 입력
' 인수   : rng      - 입력할 셀
'          vntValue - 입력할 값
'          blnMsg   - True: 유효성 실패 시 메시지 표시 (기본: False)
' 예시   : SetIfValTrue Sheet1.Range("B2"), "완료"
Public Sub SetIfValTrue(ByVal rng As Range, _
                        ByVal vntValue As Variant, _
                        Optional ByVal blnMsg As Boolean = False)
    Application.Run REF & "SetIfValTrue", rng, vntValue, blnMsg
End Sub
