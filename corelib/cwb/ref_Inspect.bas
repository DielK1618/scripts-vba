Attribute VB_Name = "ref_Inspect"
Option Explicit
' 버전 : 2026-10-16 16:32 · 신규 - am_Inspect 래퍼

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Inspect                                             │
' │  역할 : am_Inspect Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘
' 참고 : 결과 시트는 이 CWB(ThisWorkbook)에 추가됨 (wb 인수 자동 입력)
'        이 모듈 자체는 UIAutomationClient 참조가 필요 없음 (참조는 corelib.xlam 쪽에만 필요)
'
' 사용 예 (직접 실행 창)
'   DumpWindows                  ' 1. 열린 창 목록 → 대상 창 이름 확인
'   DumpElementTree "메신저"     ' 2. 그 창의 요소 트리
'   DumpElementAtCursor 5        ' 3. 5초 안에 마우스를 대상 요소 위에 올려 두기

Private Const REF As String = "corelib.xlam!am_Inspect."

' ── 창 목록 ───────────────────────────────────────────────────

' 목적   : 바탕화면의 최상위 창 목록 반환
' 인수   : strKeyword - 창 이름 또는 프로세스 이름 필터 (생략 시 전체)
' 반환   : Variant - 2차원 배열 (1행 머리글) / 실패 시 Empty
' 예시   : arr = ListWindows("메신저")
Public Function ListWindows(Optional ByVal strKeyword As String = "") As Variant
    ListWindows = Application.Run(REF & "ListWindows", strKeyword)
End Function

' 목적   : 최상위 창 목록을 이 CWB 의 새 시트에 기록
' 인수   : strKeyword - 창 이름 또는 프로세스 이름 필터 (생략 시 전체)
' 반환   : Long - 기록한 창 수 (실패 시 0)
' 예시   : DumpWindows
Public Function DumpWindows(Optional ByVal strKeyword As String = "") As Long
    DumpWindows = Application.Run(REF & "DumpWindows", strKeyword, ThisWorkbook)
End Function

' ── 요소 트리 ─────────────────────────────────────────────────

' 목적   : 지정한 창 안의 UI 요소를 트리 순서로 반환
' 인수   : strWindowKeyword - 창 이름에 포함된 문자열 (없으면 프로세스 이름으로 재탐색)
'          lngMaxDepth      - 최대 깊이 (기본 10)
'          lngMaxRows       - 최대 요소 수 (기본 3000)
'          blnRawView       - True=모든 요소 / False=컨트롤 요소만 (기본 True)
' 반환   : Variant - 2차원 배열 (1행 머리글) / 실패 시 Empty
' 예시   : arr = GetElementTree("메신저", 8)
Public Function GetElementTree(ByVal strWindowKeyword As String, _
                               Optional ByVal lngMaxDepth As Long = 10, _
                               Optional ByVal lngMaxRows As Long = 3000, _
                               Optional ByVal blnRawView As Boolean = True) As Variant
    GetElementTree = Application.Run(REF & "GetElementTree", strWindowKeyword, lngMaxDepth, lngMaxRows, blnRawView)
End Function

' 목적   : 지정한 창 안의 UI 요소 트리를 이 CWB 의 새 시트에 기록
' 인수   : GetElementTree 와 동일
' 반환   : Long - 기록한 요소 수 (실패 시 0)
' 예시   : DumpElementTree "메신저"
Public Function DumpElementTree(ByVal strWindowKeyword As String, _
                                Optional ByVal lngMaxDepth As Long = 10, _
                                Optional ByVal lngMaxRows As Long = 3000, _
                                Optional ByVal blnRawView As Boolean = True) As Long
    DumpElementTree = Application.Run(REF & "DumpElementTree", strWindowKeyword, lngMaxDepth, lngMaxRows, blnRawView, ThisWorkbook)
End Function

' ── 마우스 위치 요소 ──────────────────────────────────────────

' 목적   : 대기 후 마우스 커서 아래 요소와 상위 경로(최상위 창까지)를 반환
' 인수   : lngDelaySec - 커서를 대상 위에 올려 둘 대기 시간(초, 기본 3)
'          blnRawView  - True=모든 요소 기준 / False=컨트롤 요소 기준 (기본 True)
' 반환   : Variant - 2차원 배열 (마지막 행=커서 아래 요소) / 실패 시 Empty
Public Function GetElementAtCursor(Optional ByVal lngDelaySec As Long = 3, _
                                   Optional ByVal blnRawView As Boolean = True) As Variant
    GetElementAtCursor = Application.Run(REF & "GetElementAtCursor", lngDelaySec, blnRawView)
End Function

' 목적   : 대기 후 마우스 커서 아래 요소와 상위 경로를 이 CWB 의 새 시트에 기록
' 인수   : GetElementAtCursor 와 동일
' 반환   : Long - 기록한 행 수 (실패 시 0)
' 예시   : DumpElementAtCursor 5
Public Function DumpElementAtCursor(Optional ByVal lngDelaySec As Long = 3, _
                                    Optional ByVal blnRawView As Boolean = True) As Long
    DumpElementAtCursor = Application.Run(REF & "DumpElementAtCursor", lngDelaySec, blnRawView, ThisWorkbook)
End Function
