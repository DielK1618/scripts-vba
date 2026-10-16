Attribute VB_Name = "am_Credential"
Option Explicit
' 버전 : 2026-10-16 16:32 · 신규 - 로그인 정보 입력 창·자격 증명 관리자

' ┌─────────────────────────────────────────────────────────┐
' │  am_Credential                                          │
' │  역할 : 로그인 정보 입력 창, Windows 자격 증명 관리자 저장·읽기 │
' └─────────────────────────────────────────────────────────┘
' ※ 외부 참조 없음 (Windows API 만 사용)
'
' ■ 두 가지 방식
'   1) 실행할 때마다 입력 - PromptCredential
'      Windows 기본 "자격 증명 입력" 창을 띄워 아이디·비밀번호를 받음 (비밀번호는 ● 로 표시, 어디에도 저장 안 함)
'   2) 자격 증명 관리자에 저장 - SaveCredential / ReadCredential / DeleteCredential / HasCredential
'      Windows 가 로그인 계정 기준으로 암호화해 보관 (제어판 > 자격 증명 관리자 > Windows 자격 증명 > 일반 자격 증명)
'      ※ 같은 Windows 계정으로 실행되는 프로그램은 읽을 수 있음. 회사 보안 솔루션이 "매크로가 저장된 비밀번호를
'        읽는 동작"을 경고·차단할 수 있으므로 실사용 전 IT 보안 정책 확인 (2026-10-16 사용자와 협의 - 기본은 1번 방식)
'
' ■ 반환값 형식 (PromptCredential / ReadCredential 공통)
'   성공 시 Array(아이디, 비밀번호) - vntCred(0) = 아이디, vntCred(1) = 비밀번호
'   취소·없음 시 Empty → IsArray(vntCred) 로 확인
'
' ■ 사용 예 (CWB 에서 ref_Credential 래퍼로 호출)
'   Dim vntCred As Variant
'   vntCred = PromptCredential("WMC Messenger 로그인", "아이디와 비밀번호를 입력하세요.")
'   If Not IsArray(vntCred) Then Exit Sub            ' 취소
'   ... vntCred(0), vntCred(1) 사용 ...
'   vntCred = Empty                                  ' 다 쓴 뒤 바로 비움
'
'   SaveCredential "corelib/WMC_Messenger", "hong", "비밀번호"   ' 저장 (같은 이름이면 덮어씀)
'   vntCred = ReadCredential("corelib/WMC_Messenger")             ' 읽기
'   DeleteCredential "corelib/WMC_Messenger"                      ' 삭제
'
' ■ 저장 이름(strTarget) 규칙: 다른 프로그램과 겹치지 않도록 "corelib/프로그램명" 형식 권장

' ── Windows API ───────────────────────────────────────────────
Private Declare PtrSafe Function CredReadW Lib "advapi32" _
        (ByVal TargetName As LongPtr, ByVal CredType As Long, ByVal Flags As Long, ByRef pCredential As LongPtr) As Long
Private Declare PtrSafe Function CredWriteW Lib "advapi32" _
        (ByRef Credential As CREDENTIALW, ByVal Flags As Long) As Long
Private Declare PtrSafe Function CredDeleteW Lib "advapi32" _
        (ByVal TargetName As LongPtr, ByVal CredType As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Sub CredFree Lib "advapi32" (ByVal Buffer As LongPtr)
Private Declare PtrSafe Function CredUIPromptForCredentialsW Lib "credui" _
        (ByRef pUiInfo As CREDUI_INFOW, ByVal pszTargetName As LongPtr, ByVal Reserved As LongPtr, _
         ByVal dwAuthError As Long, ByVal pszUserName As LongPtr, ByVal ulUserNameBufferSize As Long, _
         ByVal pszPassword As LongPtr, ByVal ulPasswordBufferSize As Long, ByRef pfSave As Long, _
         ByVal dwFlags As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Private Type FILETIME
    lngLow  As Long
    lngHigh As Long
End Type

Private Type CREDENTIALW
    Flags              As Long
    CredType           As Long
    TargetName         As LongPtr
    Comment            As LongPtr
    LastWritten        As FILETIME
    CredentialBlobSize As Long
    CredentialBlob     As LongPtr
    Persist            As Long
    AttributeCount     As Long
    Attributes         As LongPtr
    TargetAlias        As LongPtr
    UserName           As LongPtr
End Type

Private Type CREDUI_INFOW
    cbSize         As Long
    hwndParent     As LongPtr
    pszMessageText As LongPtr
    pszCaptionText As LongPtr
    hbmBanner      As LongPtr
End Type

Private Const CRED_TYPE_GENERIC          As Long = 1
Private Const CRED_PERSIST_LOCAL_MACHINE As Long = 2     ' 이 PC 의 이 사용자에게만 저장 (다른 PC 로 로밍 안 함)

Private Const CREDUI_FLAGS_DO_NOT_PERSIST       As Long = &H2
Private Const CREDUI_FLAGS_EXCLUDE_CERTIFICATES As Long = &H8
Private Const CREDUI_FLAGS_ALWAYS_SHOW_UI       As Long = &H80
Private Const CREDUI_FLAGS_GENERIC_CREDENTIALS  As Long = &H40000

Private Const CREDUI_MAX_USERNAME_LENGTH As Long = 513
Private Const CREDUI_MAX_PASSWORD_LENGTH As Long = 256
Private Const ERROR_CANCELLED            As Long = 1223

' ══════════════════════════════════════════════════════════
'  입력 창
' ══════════════════════════════════════════════════════════

' 목적   : Windows 자격 증명 입력 창으로 아이디·비밀번호를 받음 (비밀번호 ● 표시, 저장하지 않음)
' 인수   : strCaption     - 창 제목 (예: "WMC Messenger 로그인")
'          strMessage     - 안내 문구
'          strDefaultUser - 아이디 칸에 미리 넣을 값 (생략 가능)
' 반환   : Variant - Array(아이디, 비밀번호) / 취소·오류 시 Empty
' 예시   : vntCred = PromptCredential("WMC Messenger 로그인", "아이디와 비밀번호를 입력하세요.")
Public Function PromptCredential(Optional ByVal strCaption As String = "로그인", _
                                 Optional ByVal strMessage As String = "아이디와 비밀번호를 입력하세요.", _
                                 Optional ByVal strDefaultUser As String = "") As Variant
    On Error GoTo ErrHandler

    Dim udtInfo   As CREDUI_INFOW
    Dim strTarget As String
    Dim strUser   As String
    Dim strPass   As String
    Dim lngSave   As Long
    Dim lngResult As Long

    PromptCredential = Empty

    udtInfo.cbSize = LenB(udtInfo)
    udtInfo.hwndParent = GetActiveWindow()
    udtInfo.pszMessageText = StrPtr(strMessage)
    udtInfo.pszCaptionText = StrPtr(strCaption)

    ' 일반 자격 증명 입력 창은 대상 이름이 필요함 (저장은 하지 않으므로 표시용)
    strTarget = strCaption
    strUser = Left(strDefaultUser, CREDUI_MAX_USERNAME_LENGTH - 1) & _
              String(CREDUI_MAX_USERNAME_LENGTH - Len(Left(strDefaultUser, CREDUI_MAX_USERNAME_LENGTH - 1)), vbNullChar)
    strPass = String(CREDUI_MAX_PASSWORD_LENGTH + 1, vbNullChar)

    lngResult = CredUIPromptForCredentialsW(udtInfo, StrPtr(strTarget), 0, 0, _
                                            StrPtr(strUser), CREDUI_MAX_USERNAME_LENGTH, _
                                            StrPtr(strPass), CREDUI_MAX_PASSWORD_LENGTH + 1, _
                                            lngSave, _
                                            CREDUI_FLAGS_GENERIC_CREDENTIALS Or CREDUI_FLAGS_ALWAYS_SHOW_UI Or _
                                            CREDUI_FLAGS_DO_NOT_PERSIST Or CREDUI_FLAGS_EXCLUDE_CERTIFICATES)

    If lngResult = 0 Then
        PromptCredential = Array(prv_TrimNull(strUser), prv_TrimNull(strPass))
    ElseIf lngResult <> ERROR_CANCELLED Then
        Debug.Print "am_Credential.PromptCredential: 입력 창 오류 " & lngResult
    End If

    strPass = String(Len(strPass), vbNullChar)    ' 버퍼 비움
    Exit Function

ErrHandler:
    Debug.Print "am_Credential.PromptCredential: 오류 " & Err.Number & ": " & Err.Description
    PromptCredential = Empty
End Function

' ══════════════════════════════════════════════════════════
'  자격 증명 관리자
' ══════════════════════════════════════════════════════════

' 목적   : 자격 증명 관리자에 아이디·비밀번호 저장 (같은 이름이 있으면 덮어씀)
' 인수   : strTarget   - 저장 이름 (예: "corelib/WMC_Messenger")
'          strUser     - 아이디
'          strPassword - 비밀번호
' 반환   : Boolean - 저장 성공 여부
' 예시   : SaveCredential "corelib/WMC_Messenger", "hong", "비밀번호"
Public Function SaveCredential(ByVal strTarget As String, _
                               ByVal strUser As String, _
                               ByVal strPassword As String) As Boolean
    On Error GoTo ErrHandler

    Dim udtCred As CREDENTIALW

    If Len(strTarget) = 0 Then Exit Function

    udtCred.CredType = CRED_TYPE_GENERIC
    udtCred.TargetName = StrPtr(strTarget)
    udtCred.UserName = StrPtr(strUser)
    udtCred.CredentialBlobSize = LenB(strPassword)
    udtCred.CredentialBlob = StrPtr(strPassword)
    udtCred.Persist = CRED_PERSIST_LOCAL_MACHINE

    SaveCredential = (CredWriteW(udtCred, 0) <> 0)
    If Not SaveCredential Then Debug.Print "am_Credential.SaveCredential: 저장 실패 (Windows 오류 " & Err.LastDllError & ")"
    Exit Function

ErrHandler:
    Debug.Print "am_Credential.SaveCredential: 오류 " & Err.Number & ": " & Err.Description
    SaveCredential = False
End Function

' 목적   : 자격 증명 관리자에서 아이디·비밀번호 읽기
' 인수   : strTarget - 저장 이름
' 반환   : Variant - Array(아이디, 비밀번호) / 없으면 Empty
' 예시   : vntCred = ReadCredential("corelib/WMC_Messenger")
Public Function ReadCredential(ByVal strTarget As String) As Variant
    On Error GoTo ErrHandler

    Dim pCred    As LongPtr
    Dim udtCred  As CREDENTIALW
    Dim arrBlob() As Byte
    Dim strPass  As String
    Dim strUser  As String

    ReadCredential = Empty
    If Len(strTarget) = 0 Then Exit Function
    If CredReadW(StrPtr(strTarget), CRED_TYPE_GENERIC, 0, pCred) = 0 Then Exit Function

    CopyMemory udtCred, ByVal pCred, LenB(udtCred)

    If udtCred.CredentialBlobSize > 0 Then
        ReDim arrBlob(0 To udtCred.CredentialBlobSize - 1)
        CopyMemory arrBlob(0), ByVal udtCred.CredentialBlob, udtCred.CredentialBlobSize
        strPass = arrBlob                          ' 바이트(UTF-16) → 문자열
        Erase arrBlob
    End If
    strUser = prv_PtrToString(udtCred.UserName)

    CredFree pCred
    ReadCredential = Array(strUser, strPass)
    strPass = ""
    Exit Function

ErrHandler:
    Debug.Print "am_Credential.ReadCredential: 오류 " & Err.Number & ": " & Err.Description
    If pCred <> 0 Then CredFree pCred
    ReadCredential = Empty
End Function

' 목적   : 자격 증명 관리자에 저장 이름이 있는지 확인
' 인수   : strTarget - 저장 이름
' 반환   : Boolean
Public Function HasCredential(ByVal strTarget As String) As Boolean
    Dim pCred As LongPtr

    If Len(strTarget) = 0 Then Exit Function
    If CredReadW(StrPtr(strTarget), CRED_TYPE_GENERIC, 0, pCred) <> 0 Then
        CredFree pCred
        HasCredential = True
    End If
End Function

' 목적   : 자격 증명 관리자에서 삭제
' 인수   : strTarget - 저장 이름
' 반환   : Boolean - 삭제 성공 여부 (없던 이름이면 False)
' 예시   : DeleteCredential "corelib/WMC_Messenger"
Public Function DeleteCredential(ByVal strTarget As String) As Boolean
    If Len(strTarget) = 0 Then Exit Function
    DeleteCredential = (CredDeleteW(StrPtr(strTarget), CRED_TYPE_GENERIC, 0) <> 0)
End Function

' ══════════════════════════════════════════════════════════
'  내부
' ══════════════════════════════════════════════════════════

' 목적   : 메모리의 UTF-16 문자열 포인터 → VBA 문자열
Private Function prv_PtrToString(ByVal pStr As LongPtr) As String
    Dim lngLen As Long

    If pStr = 0 Then Exit Function
    lngLen = lstrlenW(pStr)
    If lngLen = 0 Then Exit Function

    prv_PtrToString = String(lngLen, vbNullChar)
    CopyMemory ByVal StrPtr(prv_PtrToString), ByVal pStr, lngLen * 2
End Function

' 목적   : 첫 Null 문자 앞까지만
Private Function prv_TrimNull(ByVal strText As String) As String
    Dim lngPos As Long

    lngPos = InStr(strText, vbNullChar)
    If lngPos > 0 Then
        prv_TrimNull = Left(strText, lngPos - 1)
    Else
        prv_TrimNull = strText
    End If
End Function
