Attribute VB_Name = "ref_Credential"
Option Explicit
' 버전 : 2026-10-16 16:32 · 신규 - am_Credential 래퍼

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Credential                                          │
' │  역할 : am_Credential Application.Run 래퍼               │
' └─────────────────────────────────────────────────────────┘
' 반환값 형식: PromptCredential / ReadCredential → Array(아이디, 비밀번호), 취소·없음 → Empty (IsArray 로 확인)
' 사용 예는 xlam 의 am_Credential 모듈 상단 주석 참고

Private Const REF As String = "corelib.xlam!am_Credential."

' 목적   : Windows 자격 증명 입력 창으로 아이디·비밀번호를 받음 (비밀번호 ● 표시, 저장하지 않음)
' 인수   : strCaption     - 창 제목
'          strMessage     - 안내 문구
'          strDefaultUser - 아이디 칸에 미리 넣을 값 (생략 가능)
' 반환   : Variant - Array(아이디, 비밀번호) / 취소 시 Empty
' 예시   : vntCred = PromptCredential("WMC Messenger 로그인", "아이디와 비밀번호를 입력하세요.")
Public Function PromptCredential(Optional ByVal strCaption As String = "로그인", _
                                 Optional ByVal strMessage As String = "아이디와 비밀번호를 입력하세요.", _
                                 Optional ByVal strDefaultUser As String = "") As Variant
    PromptCredential = Application.Run(REF & "PromptCredential", strCaption, strMessage, strDefaultUser)
End Function

' 목적   : 자격 증명 관리자에 아이디·비밀번호 저장 (같은 이름이 있으면 덮어씀)
' 인수   : strTarget   - 저장 이름 (예: "corelib/WMC_Messenger")
'          strUser     - 아이디
'          strPassword - 비밀번호
' 반환   : Boolean - 저장 성공 여부
Public Function SaveCredential(ByVal strTarget As String, _
                               ByVal strUser As String, _
                               ByVal strPassword As String) As Boolean
    SaveCredential = Application.Run(REF & "SaveCredential", strTarget, strUser, strPassword)
End Function

' 목적   : 자격 증명 관리자에서 아이디·비밀번호 읽기
' 인수   : strTarget - 저장 이름
' 반환   : Variant - Array(아이디, 비밀번호) / 없으면 Empty
Public Function ReadCredential(ByVal strTarget As String) As Variant
    ReadCredential = Application.Run(REF & "ReadCredential", strTarget)
End Function

' 목적   : 자격 증명 관리자에 저장 이름이 있는지 확인
' 인수   : strTarget - 저장 이름
' 반환   : Boolean
Public Function HasCredential(ByVal strTarget As String) As Boolean
    HasCredential = Application.Run(REF & "HasCredential", strTarget)
End Function

' 목적   : 자격 증명 관리자에서 삭제
' 인수   : strTarget - 저장 이름
' 반환   : Boolean - 삭제 성공 여부 (없던 이름이면 False)
Public Function DeleteCredential(ByVal strTarget As String) As Boolean
    DeleteCredential = Application.Run(REF & "DeleteCredential", strTarget)
End Function
