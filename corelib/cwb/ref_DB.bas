Attribute VB_Name = "ref_DB"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_DB                                                 │
' │  역할 : am_DB Application.Run 래퍼                     │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_DB."

' ── 상수 재선언 (CWB 직접 참조용) ────────────────────────────
' xlam 에 정의된 Public Const 는 CWB 에서 이름으로 접근 불가
' → ref_DB 에서 동일 값으로 재선언하여 CWB 코드에서 사용

Public Const DB_ODBC_DRIVER As String = "MySQL ODBC 8.2 Unicode Driver"
Public Const DB_EXCEL_VER   As String = "12.0"
Public Const DB_ACCESS_VER  As String = "12.0"

Public Const DB_TYPE_SERVER As String = "서버"
Public Const DB_TYPE_EXCEL  As String = "엑셀"
Public Const DB_TYPE_ACCESS As String = "엑세스"

' ── 연결 문자열 ───────────────────────────────────────────────

' 목적   : DB 연결 문자열 반환
' 인수   : strType   - DB 종류 (DB_TYPE_SERVER / DB_TYPE_EXCEL / DB_TYPE_ACCESS)
'          strFile   - 파일 경로 (엑셀/Access 시)
'          strServer - 서버 주소 (서버 DB 시)
'          strPort   - 포트 번호
'          strDB     - DB 이름
'          strID     - 사용자 ID
'          strPW     - 비밀번호
'          blnHeader - True: 엑셀 첫 행을 헤더로 처리
' 반환   : String - ADODB 연결 문자열
' 예시   : conStr(DB_TYPE_ACCESS, "C:\data.accdb")
Public Function conStr(ByVal strType As String, _
                       Optional ByVal strFile As String = "", _
                       Optional ByVal strServer As String = "", _
                       Optional ByVal strPort As String = "", _
                       Optional ByVal strDB As String = "", _
                       Optional ByVal strID As String = "", _
                       Optional ByVal strPW As String = "", _
                       Optional ByVal blnHeader As Boolean = False) As String
    conStr = Application.Run(REF & "conStr", _
                             strType, strFile, strServer, strPort, strDB, strID, strPW, blnHeader)
End Function

' ── 쿼리 실행 ────────────────────────────────────────────────

' 목적   : SQL 쿼리 배열을 순서대로 실행 (INSERT/UPDATE/DELETE)
' 인수   : arrQuery  - 실행할 SQL 문자열 배열
'          strType   - DB 종류
'          strFile   - 파일 경로 (엑셀/Access 시)
'          strServer - 서버 주소 / strPort - 포트 / strDB - DB 이름
'          strID     - 사용자 ID / strPW - 비밀번호
' 반환   : Boolean - True: 전체 성공 / False: 오류 발생
' 예시   : ExecuteQueryArr(Array("DELETE FROM tbl WHERE ID=1"), DB_TYPE_ACCESS, "C:\db.accdb")
Public Function ExecuteQueryArr(ByVal arrQuery As Variant, _
                                ByVal strType As String, _
                                Optional ByVal strFile As String = "", _
                                Optional ByVal strServer As String = "", _
                                Optional ByVal strPort As String = "", _
                                Optional ByVal strDB As String = "", _
                                Optional ByVal strID As String = "", _
                                Optional ByVal strPW As String = "") As Boolean
    ExecuteQueryArr = Application.Run(REF & "ExecuteQueryArr", _
                                      arrQuery, strType, strFile, strServer, strPort, strDB, strID, strPW)
End Function

' 목적   : SELECT 쿼리 결과를 시트 셀에 출력
' 인수   : rngTarget    - 출력 시작 셀
'          arrQuery     - SELECT SQL 문자열 배열
'          strType      - DB 종류
'          strFile      - 파일 경로 / strServer / strPort / strDB / strID / strPW
'          blnTranspose - True: 행/열 전치 출력
'          intMoveCells - 헤더 이후 데이터 시작 열 오프셋
'          blnHeader    - True: 헤더 행 포함 출력
' 예시   : SelectQuery(Sheet1.Range("A1"), Array("SELECT * FROM tbl"), DB_TYPE_ACCESS, strFile)
Public Sub SelectQuery(ByVal rngTarget As Range, _
                       ByVal arrQuery As Variant, _
                       ByVal strType As String, _
                       Optional ByVal strFile As String = "", _
                       Optional ByVal strServer As String = "", _
                       Optional ByVal strPort As String = "", _
                       Optional ByVal strDB As String = "", _
                       Optional ByVal strID As String = "", _
                       Optional ByVal strPW As String = "", _
                       Optional ByVal blnTranspose As Boolean = False, _
                       Optional ByVal intMoveCells As Integer = 0, _
                       Optional ByVal blnHeader As Boolean = False)
    Application.Run REF & "SelectQuery", _
                    rngTarget, arrQuery, strType, strFile, strServer, strPort, strDB, strID, strPW, _
                    blnTranspose, intMoveCells, blnHeader
End Sub

' 목적   : SELECT 쿼리 결과를 2차원 Variant 배열로 반환
' 인수   : arrQuery     - SELECT SQL 문자열 배열
'          strType      - DB 종류
'          strFile      - 파일 경로 / strServer / strPort / strDB / strID / strPW
'          blnHeader    - True: 첫 행에 필드명 포함
'          blnTranspose - True: 행/열 전치
' 반환   : Variant - 2차원 배열 (행 × 열)
' 예시   : arr = SelectQueryArr(Array("SELECT * FROM tbl"), DB_TYPE_ACCESS, strFile)
Public Function SelectQueryArr(ByVal arrQuery As Variant, _
                               ByVal strType As String, _
                               Optional ByVal strFile As String = "", _
                               Optional ByVal strServer As String = "", _
                               Optional ByVal strPort As String = "", _
                               Optional ByVal strDB As String = "", _
                               Optional ByVal strID As String = "", _
                               Optional ByVal strPW As String = "", _
                               Optional ByVal blnHeader As Boolean = False, _
                               Optional ByVal blnTranspose As Boolean = False) As Variant
    SelectQueryArr = Application.Run(REF & "SelectQueryArr", _
                                     arrQuery, strType, strFile, strServer, strPort, strDB, strID, strPW, _
                                     blnHeader, blnTranspose)
End Function

' ── 쿼리 생성 ────────────────────────────────────────────────

' 목적   : INSERT / UPDATE SQL 문자열 생성
' 인수   : strQType  - "INSERT" 또는 "UPDATE"
'          strTable  - 테이블명
'          arrData   - 필드명=값 쌍 배열 (예: Array("이름", "홍길동", "나이", 30))
'          strWhere  - WHERE 조건 (UPDATE 시 필수)
'          blnParens - True: 테이블명을 대괄호로 감쌈
' 반환   : String - 완성된 SQL 문자열
' 예시   : UpsertQuery("INSERT", "tbl_User", Array("이름", "홍길동"))
Public Function UpsertQuery(ByVal strQType As String, _
                            ByVal strTable As String, _
                            Optional ByVal arrData As Variant, _
                            Optional ByVal strWhere As String = "", _
                            Optional ByVal blnParens As Boolean = False) As String
    UpsertQuery = Application.Run(REF & "UpsertQuery", _
                                  strQType, strTable, arrData, strWhere, blnParens)
End Function

' 목적   : Range 데이터를 SQL INSERT VALUES 절 문자열로 변환
' 인수   : rngTarget - 변환할 범위 (헤더 제외 데이터 행)
'          arrTypes  - 각 열의 DB 타입명 배열
' 반환   : String - "VALUES ('val1', 2, ...)" 형태의 SQL 절
' 예시   : ConvertRangeToSQL(Sheet1.Range("A2:C10"), Array("Text", "Long", "Text"))
Public Function ConvertRangeToSQL(ByVal rngTarget As Range, _
                                  ByVal arrTypes As Variant) As String
    ConvertRangeToSQL = Application.Run(REF & "ConvertRangeToSQL", rngTarget, arrTypes)
End Function

' ── 스키마 조회 ───────────────────────────────────────────────

' 목적   : DB 내 테이블 이름 목록 반환
' 인수   : strType     - DB 종류 (DB_TYPE_ACCESS / DB_TYPE_EXCEL)
'          strFilePath - DB 파일 경로
' 반환   : Variant - 테이블명 1차원 배열
' 예시   : arr = GetDbTables(DB_TYPE_ACCESS, "C:\data.accdb")
Public Function GetDbTables(ByVal strType As String, _
                            ByVal strFilePath As String) As Variant
    GetDbTables = Application.Run(REF & "GetDbTables", strType, strFilePath)
End Function

' 목적   : 테이블 필드 정보 조회 (필드명, 타입 등)
' 인수   : strType  - DB 종류
'          strTable - 테이블명
'          strFile / strServer / strPort / strDB / strID / strPW - 연결 정보
' 반환   : Variant - 필드 정보 배열 (필드명, 타입명, 타입코드 ...)
' 예시   : arr = GetFieldInfo(DB_TYPE_ACCESS, "tbl_User", "C:\data.accdb")
Public Function GetFieldInfo(ByVal strType As String, _
                             ByVal strTable As String, _
                             Optional ByVal strFile As String = "", _
                             Optional ByVal strServer As String = "", _
                             Optional ByVal strPort As String = "", _
                             Optional ByVal strDB As String = "", _
                             Optional ByVal strID As String = "", _
                             Optional ByVal strPW As String = "") As Variant
    GetFieldInfo = Application.Run(REF & "GetFieldInfo", _
                                   strType, strTable, strFile, strServer, strPort, strDB, strID, strPW)
End Function

' ── 타입 처리 ────────────────────────────────────────────────

' 목적   : ADODB DataTypeEnum 코드를 타입명 문자열로 변환
' 인수   : intTypeCode - ADODB 타입 코드 (예: 202 → "VarWChar")
' 반환   : String - 타입명 문자열
' 예시   : GetDataTypeName(202) → "VarWChar"
Public Function GetDataTypeName(ByVal intTypeCode As Integer) As String
    GetDataTypeName = Application.Run(REF & "GetDataTypeName", intTypeCode)
End Function

' 목적   : 타입명 문자열을 ADODB DataTypeEnum 코드로 변환
' 인수   : strTypeName - 타입명 (예: "VarWChar")
' 반환   : Integer - ADODB 타입 코드 (예: 202)
' 예시   : GetDataTypeCode("VarWChar") → 202
Public Function GetDataTypeCode(ByVal strTypeName As String) As Integer
    GetDataTypeCode = CInt(Application.Run(REF & "GetDataTypeCode", strTypeName))
End Function

' 목적   : 값을 DB 타입에 맞게 SQL 포함 가능한 문자열로 포맷
' 인수   : vntValue    - 포맷할 값
'          strTypeName - DB 타입명 (예: "VarWChar", "Long", "Date")
' 반환   : String - SQL 에 삽입 가능한 형태 (문자열은 따옴표 포함)
' 예시   : FormatValueForSQLByDBType("홍길동", "VarWChar") → "'홍길동'"
Public Function FormatValueForSQLByDBType(ByVal vntValue As Variant, _
                                          ByVal strTypeName As String) As String
    FormatValueForSQLByDBType = Application.Run(REF & "FormatValueForSQLByDBType", vntValue, strTypeName)
End Function

' 목적   : 값이 해당 DB 타입에 유효한지 검증
' 인수   : vntValue    - 검증할 값
'          strTypeName - DB 타입명
' 반환   : Boolean - True: 유효 / False: 타입 불일치
' 예시   : ValidateValueForDBType("2024-01-01", "Date") → True
Public Function ValidateValueForDBType(ByVal vntValue As Variant, _
                                       ByVal strTypeName As String) As Boolean
    ValidateValueForDBType = Application.Run(REF & "ValidateValueForDBType", vntValue, strTypeName)
End Function

' ── 엑셀 DB 유틸리티 ─────────────────────────────────────────

' 목적   : 범위 내 빈 ID 셀에 자동 번호 채우기
' 인수   : rng - ID 열 범위 (데이터 행 전체)
' 예시   : AutoIDs(tbl.ListColumns("ID").DataBodyRange)
Public Sub AutoIDs(ByVal rng As Range)
    Application.Run REF & "AutoIDs", rng
End Sub

' 목적   : 엑셀 DB 시트에서 특정 ID 레코드 삭제
' 인수   : wb           - 엑셀 DB 워크북
'          strSheetName - 대상 시트명
'          strID        - 삭제할 레코드 ID
' 예시   : DelExcelRecQuery(wbDB, "tbl_User", "U001")
Public Sub DelExcelRecQuery(ByVal wb As Workbook, _
                            ByVal strSheetName As String, _
                            ByVal strID As String)
    Application.Run REF & "DelExcelRecQuery", wb, strSheetName, strID
End Sub

' ── 테이블 분석 ───────────────────────────────────────────────

' 목적   : 테이블에서 조건에 맞는 필드명 배열 반환
' 인수   : lngFontColor - 글자색 필터 (-1: 무시)
'          lngBgColor   - 배경색 필터 (-1: 무시)
'          tbl          - 대상 ListObject (기본: 활성 테이블)
' 반환   : Variant - 조건에 맞는 필드명 배열
' 예시   : GetFields(lngBgColor:=RGB(255,255,0)) → 노란 배경 필드명 배열
Public Function GetFields(Optional ByVal lngFontColor As Long = -1, _
                          Optional ByVal lngBgColor As Long = -1, _
                          Optional ByVal tbl As ListObject) As Variant
    GetFields = Application.Run(REF & "GetFields", lngFontColor, lngBgColor, tbl)
End Function

' 목적   : 필드명 배열과 DB 스키마를 매칭하여 필드명+타입 쌍 배열 반환
' 인수   : arrFields - 필드명 배열
'          strType   - DB 종류
'          strTable  - 테이블명
'          strFile / strServer / strPort / strDB / strID / strPW - 연결 정보
'          blnParens - True: 필드명을 대괄호로 감쌈
' 반환   : Variant - {필드명, 타입명} 쌍 2차원 배열
' 예시   : GetFieldAndType(Array("이름","나이"), DB_TYPE_ACCESS, "tbl_User", strFile)
Public Function GetFieldAndType(ByVal arrFields As Variant, _
                                ByVal strType As String, _
                                ByVal strTable As String, _
                                Optional ByVal strFile As String = "", _
                                Optional ByVal strServer As String = "", _
                                Optional ByVal strPort As String = "", _
                                Optional ByVal strDB As String = "", _
                                Optional ByVal strID As String = "", _
                                Optional ByVal strPW As String = "", _
                                Optional ByVal blnParens As Boolean = False) As Variant
    GetFieldAndType = Application.Run(REF & "GetFieldAndType", _
                                      arrFields, strType, strTable, strFile, strServer, strPort, _
                                      strDB, strID, strPW, blnParens)
End Function

' 목적   : 테이블 전체 필드명 + 타입명 쌍 배열 반환 (GetFieldAndType 간편 버전)
' 인수   : strType  - DB 종류
'          strTable - 테이블명
'          strFile / strServer / strPort / strDB / strID / strPW - 연결 정보
' 반환   : Variant - {필드명, 타입명} 쌍 2차원 배열
' 예시   : GetFieldNameConnection(DB_TYPE_ACCESS, "tbl_User", "C:\data.accdb")
Public Function GetFieldNameConnection(ByVal strType As String, _
                                       ByVal strTable As String, _
                                       Optional ByVal strFile As String = "", _
                                       Optional ByVal strServer As String = "", _
                                       Optional ByVal strPort As String = "", _
                                       Optional ByVal strDB As String = "", _
                                       Optional ByVal strID As String = "", _
                                       Optional ByVal strPW As String = "") As Variant
    GetFieldNameConnection = Application.Run(REF & "GetFieldNameConnection", _
                                             strType, strTable, strFile, strServer, strPort, strDB, strID, strPW)
End Function

' ── Access DB 조작 ────────────────────────────────────────────

' 목적   : Access DB 에 새 테이블 생성 (SQL DDL 방식)
' 인수   : strFile          - Access 파일 경로
'          strTable         - 생성할 테이블명
'          arrFieldAndTypes - {필드명, 타입명} 쌍 배열
' 반환   : Boolean - True: 성공 / False: 실패
' 예시   : CreateAccessTable("C:\db.accdb", "tbl_New", Array("이름","VarWChar","나이","Long"))
Public Function CreateAccessTable(ByVal strFile As String, _
                                  ByVal strTable As String, _
                                  ByVal arrFieldAndTypes As Variant) As Boolean
    CreateAccessTable = Application.Run(REF & "CreateAccessTable", strFile, strTable, arrFieldAndTypes)
End Function

' 목적   : Access DB 에서 테이블 삭제 (SQL DDL 방식)
' 인수   : strFile  - Access 파일 경로
'          strTable - 삭제할 테이블명
' 반환   : Boolean - True: 성공 / False: 실패
' 예시   : DeleteAccessTable("C:\db.accdb", "tbl_Old")
Public Function DeleteAccessTable(ByVal strFile As String, _
                                  ByVal strTable As String) As Boolean
    DeleteAccessTable = Application.Run(REF & "DeleteAccessTable", strFile, strTable)
End Function

' 목적   : Access DB 에 새 테이블 생성 (ADOX 방식)
' 인수   : strFile          - Access 파일 경로
'          strTable         - 생성할 테이블명
'          arrFieldAndTypes - {필드명, 타입코드} 쌍 배열
' 반환   : Boolean - True: 성공 / False: 실패
' 예시   : CreateAccessTableADOX("C:\db.accdb", "tbl_New", Array("이름",202,"나이",3))
Public Function CreateAccessTableADOX(ByVal strFile As String, _
                                      ByVal strTable As String, _
                                      ByVal arrFieldAndTypes As Variant) As Boolean
    CreateAccessTableADOX = Application.Run(REF & "CreateAccessTableADOX", strFile, strTable, arrFieldAndTypes)
End Function

' 목적   : Access 테이블에서 지정 위치의 필드 삭제
' 인수   : strFile      - Access 파일 경로
'          strTable     - 테이블명
'          arrPosition  - 삭제할 필드 위치(1-based) 배열
' 반환   : Boolean - True: 성공 / False: 실패
' 예시   : DeleteAccessFields("C:\db.accdb", "tbl_User", Array(3, 5))
Public Function DeleteAccessFields(ByVal strFile As String, _
                                   ByVal strTable As String, _
                                   ByVal arrPosition As Variant) As Boolean
    DeleteAccessFields = Application.Run(REF & "DeleteAccessFields", strFile, strTable, arrPosition)
End Function
