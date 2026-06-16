Attribute VB_Name = "ref_Excel"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Excel                                              │
' │  역할 : am_Excel Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

' 참고 : 아래 프로시저는 래핑 불가 → 직접 Application.Run 사용
'        - ExecuteKeySequence          : ParamArray 파라미터
'        - ProcessRangeWithKeySequence : ParamArray 파라미터
'        - GetMousePosition            : ByRef Long 파라미터

Private Const REF As String = "corelib.xlam!am_Excel."

' ── Enum 재선언 (CWB 직접 참조용) ────────────────────────────
' xlam 에 정의된 Public Enum 은 CWB 에서 이름으로 접근 불가
' → ref_Excel 에 동일 값으로 재선언하여 CWB 코드에서 상수처럼 사용

Public Enum VirtualKeys
    VK_TAB     = &H9
    VK_RETURN  = &HD
    VK_CONTROL = &H11
    VK_MENU    = &H12
    VK_ESCAPE  = &H1B
    VK_SPACE   = &H20
    VK_LEFT    = &H25
    VK_UP      = &H26
    VK_RIGHT   = &H27
    VK_DOWN    = &H28
End Enum

Public Enum KeyActions
    ACTION_COPY        = 1
    ACTION_PASTE       = 2
    ACTION_TAB         = 3
    ACTION_ENTER       = 4
    ACTION_ALT_TAB     = 5
    ACTION_ESCAPE      = 6
    ACTION_ARROW_DOWN  = 7
    ACTION_ARROW_UP    = 8
    ACTION_ARROW_LEFT  = 9
    ACTION_ARROW_RIGHT = 10
End Enum

' ── 인쇄 / 내보내기 ──────────────────────────────────────────

' 목적   : 시트 인쇄 영역, 여백, 용지 크기, 맞춤 설정 적용
' 인수   : sht           - 대상 시트 (기본: ActiveSheet)
'          rngPrintArea  - 인쇄 영역 Range (기본: UsedRange)
'          rngTitleRows  - 반복 인쇄 행 Range (없으면 Nothing)
'          blnAddArea    - True: 기존 인쇄 영역에 추가
'          paperSize     - 용지 크기 (기본: xlPaperA4)
'          orientation   - 방향 (기본: xlPortrait 세로)
'          sngTopMargin  - 상 여백 cm (기본: 1.5)
'          sngBottomMargin - 하 여백 cm (기본: 1)
'          sngLeftMargin - 좌 여백 cm (기본: 1)
'          sngRightMargin - 우 여백 cm (기본: 1)
'          CenterH       - 가로 가운데 맞춤 (기본: True)
'          CenterV       - 세로 가운데 맞춤 (기본: False)
'          fitToPage     - 페이지 맞춤 사용 (기본: True)
'          intWide       - 가로 페이지 수 (기본: 1)
'          intTall       - 세로 페이지 수 (기본: 0 = 자동)
'          intPer        - 확대/축소 비율 % (기본: 100)
' 예시   : SetPrintPage sht:=ActiveSheet, rngPrintArea:=Sheet1.Range("A1:H50")
Public Sub SetPrintPage(Optional ByVal sht As Worksheet, _
                        Optional ByVal rngPrintArea As Range, _
                        Optional ByVal rngTitleRows As Range, _
                        Optional ByVal blnAddArea As Boolean = False, _
                        Optional ByVal paperSize As XlPaperSize = xlPaperA4, _
                        Optional ByVal orientation As XlPageOrientation = xlPortrait, _
                        Optional ByVal sngTopMargin As Double = 1.5, _
                        Optional ByVal sngBottomMargin As Double = 1, _
                        Optional ByVal sngLeftMargin As Double = 1, _
                        Optional ByVal sngRightMargin As Double = 1, _
                        Optional ByVal CenterH As Boolean = True, _
                        Optional ByVal CenterV As Boolean = False, _
                        Optional ByVal fitToPage As Boolean = True, _
                        Optional ByVal intWide As Integer = 1, _
                        Optional ByVal intTall As Integer = 0, _
                        Optional ByVal intPer As Integer = 100)
    Application.Run REF & "SetPrintPage", _
                    sht, rngPrintArea, rngTitleRows, blnAddArea, paperSize, orientation, _
                    sngTopMargin, sngBottomMargin, sngLeftMargin, sngRightMargin, _
                    CenterH, CenterV, fitToPage, intWide, intTall, intPer
End Sub

' 목적   : 시트를 PDF 파일로 내보내기
' 인수   : strFilePath  - 저장할 PDF 전체 파일 경로
'          sht          - 내보낼 시트 (기본: ActiveSheet)
'          blnOpenAfter - True: 내보낸 후 PDF 자동 열기
'          xlQual       - 품질 (기본: xlQualityStandard)
'          blnDocProps  - True: 문서 속성 포함
' 예시   : ExportPDF "C:\출력\report.pdf", ActiveSheet, blnOpenAfter:=True
Public Sub ExportPDF(ByVal strFilePath As String, _
                     Optional ByVal sht As Worksheet, _
                     Optional ByVal blnOpenAfter As Boolean = False, _
                     Optional ByVal xlQual As XlFixedFormatQuality = xlQualityStandard, _
                     Optional ByVal blnDocProps As Boolean = True)
    Application.Run REF & "ExportPDF", strFilePath, sht, blnOpenAfter, xlQual, blnDocProps
End Sub

' 목적   : 시트를 CSV 파일로 내보내기 (쉼표 구분, UTF-8)
' 인수   : strPath     - 저장 폴더 경로
'          strFileName - 저장 파일명 (확장자 포함)
'          sht         - 내보낼 시트 (기본: ActiveSheet)
' 예시   : ExportSheetToCSV "C:\출력", "data.csv"
'          ExportSheetToCSV "C:\출력", "data.csv", Sheet2
Public Sub ExportSheetToCSV(ByVal strPath As String, _
                            ByVal strFileName As String, _
                            Optional ByVal sht As Worksheet = Nothing)
    If sht Is Nothing Then Set sht = ActiveSheet
    Application.Run REF & "ExportSheetToCSV", strPath, strFileName, sht
End Sub

' ── 차트 ─────────────────────────────────────────────────────

' 목적   : 차트 개체의 데이터 범위 변경
' 인수   : strChart - 차트 개체 이름
'          rng      - 새 데이터 범위
'          ws       - 차트가 있는 시트 (기본: ActiveSheet)
' 예시   : SetChartDataRange("차트 1", Sheet1.Range("A1:C10"), Sheet1)
Public Sub SetChartDataRange(ByVal strChart As String, _
                             ByVal rng As Range, _
                             Optional ByVal ws As Worksheet)
    Application.Run REF & "SetChartDataRange", strChart, rng, ws
End Sub

' ── 도형 ─────────────────────────────────────────────────────

' 목적   : 도형에 연결된 매크로 실행
' 인수   : strShpName - 도형 이름
'          ws         - 도형이 있는 시트 (기본: ActiveSheet)
' 예시   : RunShpMacro("btn_실행", ActiveSheet)
Public Sub RunShpMacro(ByVal strShpName As String, _
                       Optional ByVal ws As Worksheet)
    Application.Run REF & "RunShpMacro", strShpName, ws
End Sub

' 목적   : 도형 텍스트 안전하게 읽기 (오류 시 "" 반환)
' 인수   : shp - 텍스트를 읽을 도형 개체
' 반환   : String - 도형 텍스트 ("" = 텍스트 없음 또는 오류)
' 예시   : GetShapeTextSafe(ActiveSheet.Shapes("lbl_Status")) → "완료"
Public Function GetShapeTextSafe(ByVal shp As Shape) As String
    GetShapeTextSafe = Application.Run(REF & "GetShapeTextSafe", shp)
End Function

' ── 키보드 ───────────────────────────────────────────────────

' 목적   : 단일 키보드 액션 실행 (KeyActions Enum 사용)
' 인수   : action    - 실행할 액션 (KeyActions Enum: ACTION_COPY 등)
'          waitAfter - 액션 후 대기 시간(ms, 기본: 100)
' 예시   : ExecuteKeyAction ACTION_COPY
'          ExecuteKeyAction ACTION_PASTE, 200
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Application.Run REF & "ExecuteKeyAction", CLng(action), waitAfter
End Sub

' ── 마우스 ───────────────────────────────────────────────────

' 목적   : 화면 절대 좌표 위치 클릭
' 인수   : lngX      - 클릭 X 좌표 (픽셀)
'          lngY      - 클릭 Y 좌표 (픽셀)
'          blnLeft   - True: 좌클릭 / False: 우클릭 (기본: True)
'          strDelay  - 클릭 전 대기 시간 문자열 (예: "00:00:01")
'          strWait   - 클릭 후 대기 시간 문자열
' 예시   : ClickAtPosition 500, 300
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    Application.Run REF & "ClickAtPosition", lngX, lngY, blnLeft, strDelay, strWait
End Sub

' 목적   : 특정 시각까지 대기 (Application.Wait 래퍼)
' 인수   : strTime - 대기 종료 시각 문자열 (예: "00:00:02" → 2초 후)
' 예시   : WaitTime "00:00:03"
Public Sub WaitTime(ByVal strTime As String)
    Application.Run REF & "WaitTime", strTime
End Sub
