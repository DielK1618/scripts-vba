Attribute VB_Name = "ref_Format"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Format                                             │
' │  역할 : am_Format Application.Run 래퍼                 │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Format."

' ── 조건부 서식 ───────────────────────────────────────────────

' 목적   : 수식 기반 조건부 서식 추가
' 인수   : rng             - 서식 적용 범위
'          strFormula      - 조건 수식 (예: "=A1>0")
'          lngFontColor    - 글자색 RGB (-1: 변경 안 함)
'          lngBackColor    - 배경색 RGB (-1: 변경 안 함)
'          xlPtn           - 패턴 종류 (기본: xlPatternNone)
'          lngPatternColor - 패턴 색상 RGB (-1: 변경 안 함)
'          lngBorderColor  - 테두리 색상 RGB (-1: 변경 안 함)
'          blnBorderTop    - True: 상단 테두리 적용
'          blnBorderBottom - True: 하단 테두리 적용
'          blnBorderLeft   - True: 좌측 테두리 적용
'          blnBorderRight  - True: 우측 테두리 적용
'          blnStopIfTrue   - True: 조건 충족 시 이후 규칙 중지
'          intPriority     - 규칙 우선순위 (1 = 가장 높음)
' 예시   : ConditionalFormattingFormula rng, "=A1=""완료""", lngBackColor:=RGB(0,255,0)
Public Sub ConditionalFormattingFormula( _
        ByVal rng As Range, _
        ByVal strFormula As String, _
        Optional ByVal lngFontColor As Long = -1, _
        Optional ByVal lngBackColor As Long = -1, _
        Optional ByVal xlPtn As XlPattern = xlPatternNone, _
        Optional ByVal lngPatternColor As Long = -1, _
        Optional ByVal lngBorderColor As Long = -1, _
        Optional ByVal blnBorderTop As Boolean = False, _
        Optional ByVal blnBorderBottom As Boolean = False, _
        Optional ByVal blnBorderLeft As Boolean = False, _
        Optional ByVal blnBorderRight As Boolean = False, _
        Optional ByVal blnStopIfTrue As Boolean = False, _
        Optional ByVal intPriority As Integer = 1)
    Application.Run REF & "ConditionalFormattingFormula", _
                    rng, strFormula, lngFontColor, lngBackColor, xlPtn, lngPatternColor, _
                    lngBorderColor, blnBorderTop, blnBorderBottom, blnBorderLeft, blnBorderRight, _
                    blnStopIfTrue, intPriority
End Sub

' 목적   : 범위의 조건부 서식 모두 삭제
' 인수   : rng - 삭제할 범위
' 예시   : ClearConditionalFormatting(Sheet1.Range("A1:H100"))
Public Sub ClearConditionalFormatting(ByVal rng As Range)
    Application.Run REF & "ClearConditionalFormatting", rng
End Sub

' 목적   : 색상 스케일 조건부 서식 추가 (최소~최대 또는 최소~중간~최대)
' 인수   : rng          - 서식 적용 범위
'          lngMinColor  - 최솟값 셀 색상 RGB
'          lngMaxColor  - 최댓값 셀 색상 RGB
'          lngMidColor  - 중간값 셀 색상 RGB (-1: 2-색상 스케일)
'          intPriority  - 규칙 우선순위 (기본: 1)
' 예시   : ConditionalFormattingColorScale rng, RGB(255,0,0), RGB(0,255,0)
'          ConditionalFormattingColorScale rng, RGB(255,0,0), RGB(0,255,0), RGB(255,255,0)
Public Sub ConditionalFormattingColorScale( _
        ByVal rng As Range, _
        ByVal lngMinColor As Long, _
        ByVal lngMaxColor As Long, _
        Optional ByVal lngMidColor As Long = -1, _
        Optional ByVal intPriority As Integer = 1)
    Application.Run REF & "ConditionalFormattingColorScale", _
                    rng, lngMinColor, lngMaxColor, lngMidColor, intPriority
End Sub

' 목적   : 데이터 막대 조건부 서식 추가
' 인수   : rng          - 서식 적용 범위
'          lngBarColor  - 막대 색상 RGB
'          blnShowValue - True: 셀 값 텍스트 표시 (기본: True)
'          blnGradient  - True: 그라데이션 막대 / False: 단색 막대 (기본: True)
'          intPriority  - 규칙 우선순위 (기본: 1)
' 예시   : ConditionalFormattingDataBar rng, RGB(0,112,192), blnGradient:=False
Public Sub ConditionalFormattingDataBar( _
        ByVal rng As Range, _
        ByVal lngBarColor As Long, _
        Optional ByVal blnShowValue As Boolean = True, _
        Optional ByVal blnGradient As Boolean = True, _
        Optional ByVal intPriority As Integer = 1)
    Application.Run REF & "ConditionalFormattingDataBar", _
                    rng, lngBarColor, blnShowValue, blnGradient, intPriority
End Sub

' ── 유효성 검사 ───────────────────────────────────────────────

' 목적   : 드롭다운 목록 유효성 검사 적용 (배열 기반)
' 인수   : rngTarget - 유효성 적용 범위
'          arrValues - 목록 항목 배열
' 예시   : ValidationList(Sheet1.Range("B2:B100"), Array("완료","진행","대기"))
Public Sub ValidationList(ByVal rngTarget As Range, ByVal arrValues As Variant)
    Application.Run REF & "ValidationList", rngTarget, arrValues
End Sub

' 목적   : 범위에 다양한 유형의 유효성 검사 적용
' 인수   : rng              - 유효성 적용 범위
'          strFormula1      - 조건 수식 1 (기본: "")
'          strFormula2      - 조건 수식 2 (Between 시 상한값)
'          xlVldType        - 유효성 유형 (기본: xlValidateInputOnly)
'          xlAlertStyle     - 오류 알림 유형 (기본: xlValidAlertStop)
'          xlOpr            - 비교 연산자 (기본: xlBetween)
'          blnIgnoreBlank   - True: 빈 셀 허용 (기본: True)
'          blnInCellDropdown - True: 드롭다운 표시 (기본: True)
'          strInputTitle    - 입력 메시지 제목
'          strInputMessage  - 입력 메시지 내용
'          strErrorTitle    - 오류 메시지 제목
'          strErrorMessage  - 오류 메시지 내용
'          xlIME            - IME 입력 모드 (기본: xlIMEModeNoControl)
'          blnShowInput     - True: 입력 메시지 표시 (기본: True)
'          blnShowError     - True: 오류 메시지 표시 (기본: True)
' 예시   : SetValidation rng, "1", "100", xlValidateWholeNumber, , xlBetween
Public Sub SetValidation( _
        ByVal rng As Range, _
        Optional ByVal strFormula1 As String = "", _
        Optional ByVal strFormula2 As String = "", _
        Optional ByVal xlVldType As XlDVType = xlValidateInputOnly, _
        Optional ByVal xlAlertStyle As XlDVAlertStyle = xlValidAlertStop, _
        Optional ByVal xlOpr As XlFormatConditionOperator = xlBetween, _
        Optional ByVal blnIgnoreBlank As Boolean = True, _
        Optional ByVal blnInCellDropdown As Boolean = True, _
        Optional ByVal strInputTitle As String = "", _
        Optional ByVal strInputMessage As String = "", _
        Optional ByVal strErrorTitle As String = "", _
        Optional ByVal strErrorMessage As String = "", _
        Optional ByVal xlIME As XlIMEMode = xlIMEModeNoControl, _
        Optional ByVal blnShowInput As Boolean = True, _
        Optional ByVal blnShowError As Boolean = True)
    Application.Run REF & "SetValidation", _
                    rng, strFormula1, strFormula2, xlVldType, xlAlertStyle, xlOpr, _
                    blnIgnoreBlank, blnInCellDropdown, strInputTitle, strInputMessage, _
                    strErrorTitle, strErrorMessage, xlIME, blnShowInput, blnShowError
End Sub

' 목적   : 범위의 유효성 검사 삭제
' 인수   : rng - 유효성을 삭제할 범위
' 예시   : ClearValidation(Sheet1.Range("B2:B100"))
Public Sub ClearValidation(ByVal rng As Range)
    Application.Run REF & "ClearValidation", rng
End Sub
