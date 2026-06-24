# TROUBLESHOOTING.md — scripts-vba

오류 발생 시 이 문서를 먼저 확인한다. 해결책이 없으면 LOG.md / 웹 검색 순으로 진행하고, 해결 후 이 문서에 추가한다.

---

## xlam 로드 / 호출

### TS-V01 · `Workbooks("corelib.xlam")` — 개체를 찾을 수 없음 오류

**증상**: xlam이 열려 있는데도 `Workbooks("corelib.xlam")`가 오류 발생

**원인**: `Workbooks` 컬렉션에 xlam 파일이 포함되지 않는 Excel 동작

**해결**: `On Error Resume Next`로 접근 시도 후 `Is Nothing` 확인

```vba
On Error Resume Next
Set wbXlam = Workbooks("corelib.xlam")
On Error GoTo ErrHandler
If wbXlam Is Nothing Then Workbooks.Open strXlamPath
```

---

### TS-V02 · xlam 로드 직후 함수 호출 — 컴파일 오류

**증상**: `corelib.xlam` 로드 직후 xlam 함수를 직접 호출하면 컴파일 시점 오류

**원인**: VBA 컴파일러가 로드 전 참조를 스캔

**해결**: `Application.Run`으로 문자열 전달 (컴파일 시점 스캔 우회)

```vba
Application.Run "corelib.xlam!am_Path.ReplacePath", arg1, arg2
```

---

### TS-V03 · `Property Get` 함수 — `Application.Run` 호출 불가

**증상**: `Application.Run "corelib.xlam!am_Core.XlamPath"` 오류

**원인**: `Application.Run`은 `Property Get`을 지원하지 않음

**해결**: `Property Get` → `Public Function`으로 변경

```vba
' ❌
Public Property Get XlamPath() As String

' ✅
Public Function XlamPath() As String
```

**적용**: am_Core의 `XlamPath` / `XlamFullName` / `Version` / `IsReady` (반영 완료)

---

### TS-V04 · `Optional` 인수가 필수 인수 앞에 위치 — 컴파일 오류

**증상**: 프로시저 선언 시 "선택적 인수가 필수 인수 앞에 올 수 없습니다" 컴파일 오류

**원인**: VBA 규칙 — `Optional` 인수는 반드시 필수 인수 뒤에 위치해야 함

**해결**: 인수 순서 재정렬 — 필수 인수 먼저, Optional 인수는 뒤로

```vba
' ❌
Public Sub SheetLock(Optional ws As Worksheet, ByVal strPW As String)

' ✅
Public Sub SheetLock(ByVal strPW As String, Optional ws As Worksheet)
```

---

## 파일 인코딩

### TS-V05 · `.bas` / `.cls` 파일 한글 깨짐

**증상**: VBE에서 가져온 모듈 파일의 한글이 깨짐

**원인**: VBE 기본 내보내기가 EUC-KR로 저장됨

**해결**: 파일을 UTF-8(BOM 없이)로 재작성 후 VBE에서 다시 가져오기. 신규 모듈은 처음부터 UTF-8로 작성.

---

## 배열 / 타입 오류

### TS-V06 · `Application.Run` — `Variant` 배열 수신 오류

**증상**: `Application.Run`으로 배열을 전달받을 때 타입 불일치 오류

**원인**: `Application.Run`은 `String()` 배열 등 특정 타입 배열을 제대로 전달하지 못함

**해결**: 배열 파라미터는 `As Variant`로 선언

```vba
' ❌
Public Sub AutoTableFilter_Arr(arrWildCards() As String)

' ✅
Public Sub AutoTableFilter_Arr(arrWildCards As Variant)
```

---

### TS-V07 · `GetSheetNames` — 2차원 배열로 반환되어 인덱스 오류

**증상**: `arrNames(i)` 접근 시 "인덱스가 유효 범위에 없습니다" 오류

**원인**: `WorksheetFunction.Transpose`가 1차원 배열을 2차원 배열로 변환

**해결**: `Transpose` 제거 → 1차원 배열 직접 구성 후 반환

---

## 파일 / 폴더

### TS-V08 · `DelFolder` — 중첩 폴더 미삭제

**증상**: 하위 폴더가 있는 폴더 삭제 시 오류 또는 일부만 삭제됨

**원인**: `Kill + RmDir`는 빈 폴더만 삭제 가능, 중첩 구조 처리 불가

**해결**: FSO `DeleteFolder(True)`로 재귀 삭제

```vba
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strPath) Then
    objFSO.DeleteFolder strPath, True
End If
```

---

## Range / 검색

### TS-V09 · `Find(What:="")` — 이전 검색어 재사용

**증상**: `Find(What:="")`로 빈 셀 검색 시 이전 Find의 검색어로 실행됨 (예상치 못한 셀 반환)

**원인**: Excel `Find`는 `What:=""`일 때 이전 Find 다이얼로그의 검색어 재사용

**해결**: `Find(What:="*")` (내용 있는 셀) + `SpecialCells(xlCellTypeBlanks)` (빈 셀) 하이브리드 사용

---

## 시트 보호

### TS-V10 · `SheetLock` — `Interior.ColorIndex`로 입력 셀 감지 누락

**증상**: RGB 색상 또는 테마 색상으로 칠한 셀이 입력 가능 셀로 인식되지 않고 잠김

**원인**: `Interior.ColorIndex`는 테마 색상과 일부 RGB 색상을 감지하지 못함

**해결**: `Interior.Pattern = xlNone` 기준 사용 (배경 없음 = 입력 가능 셀)

```vba
If cel.Interior.Pattern = xlNone Then
    cel.Locked = False
End If
```

---

## 초기화

### TS-V11 · `am_Core.IsReady` — VBE 리셋 후 `False` 유지로 기능 불작동

**증상**: VBE 리셋(Stop/재시작) 후 xlam 함수 호출 시 "초기화되지 않음" 오류

**원인**: `m_blnReady = False`가 리셋 후에도 유지되나 `Initialize()`가 자동 호출되지 않음

**해결**: lazy-init 패턴 — `IsReady` 호출 시 미초기화 상태면 `Initialize()` 자동 호출

```vba
Public Function IsReady() As Boolean
    If Not m_blnReady Then Initialize
    IsReady = m_blnReady
End Function
```
