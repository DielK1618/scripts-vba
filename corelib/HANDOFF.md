# HANDOFF.md — 다음 세션 시작 브리핑

> **보관 기록 (2026-10-16 고정)** — 대표 문서이며 모델은 수정하지 않는다. 새 기록은 같은 폴더의 `HANDOFF.<모델>.md`(예: `HANDOFF.claude.md`, `HANDOFF.codex.md`)에 쓰고, 읽을 때는 이 문서와 `HANDOFF.*.md`를 모두 읽는다. 규칙: `Git/CLAUDE.md` 3절.

> 새 채팅창에서 이 파일을 먼저 읽고 시작

---

## 프로젝트 개요

Excel VBA 범용 라이브러리 `corelib.xlam` 제작 프로젝트.
CWB(Client WorkBook)에서 `Workbooks.Open` 방식으로 xlam을 로드하고
`Application.Run "corelib.xlam!모듈.함수"` 방식으로 호출.

**프로젝트 경로:** `%USERPROFILE%\DriveSync\DEVELOPMENT\Git\scripts-vba\corelib\`

---

## 필수 참조 파일

| 파일 | 용도 |
|---|---|
| `CLAUDE.md` | 코딩 규칙, 네이밍, 에러 패턴, 설계 원칙 전체 |
| `LOG.md` | 모듈 완성 현황, 이식 현황, 설계 결정 사항 |
| `MIGRATE.md` | 이식 작업 인덱스 (⬜/✅ 체크리스트) |
| `refer/SOURCES.md` | tpl_ 소스별 이식 여부·수정 사항 분석 |
| `refer/export/*.bas` | vba.xlsm 원본 소스 (직접 Read 가능, xlsm 열 필요 없음) |

---

## 현재 완성 상태 (2026-07-24 기준)

### 완성된 am_ 모듈

| 모듈 | 주요 기능 |
|---|---|
| `am_Core` | 상수, 프로퍼티, 초기화, DPUpdate/Event/Calculate/WB_Lock |
| `am_Path` | 경로 토큰 변환, 경로 정규화 |
| `am_File` | 파일/폴더 생성·삭제·복사·검색·다이얼로그 |
| `am_DB` | DB 연결·쿼리 실행·스키마·타입 처리·Access 조작 |
| `am_Error` | 공통 에러 핸들링, 로그 기록 |
| `am_Sheet` | 백업, 표시/숨김, 정렬, SheetLock/SheetUnLock |
| `am_Table` | 테이블 CRUD·필터·정렬·검색 |
| `am_Range` | FindRange, FindCellsByColor, GetUsedRange |
| `am_Format` | 조건부 서식, 유효성 검사 |
| `am_Utils` | 배열·검사·코드생성·날짜·외부앱·도구·수식 유틸리티 |
| `am_Excel` | 인쇄/내보내기, 차트, 도형 등 Excel 객체 모델 자동화 |
| `am_Automation` | 키보드/마우스 입력 시뮬레이션, 창 탐색/활성화/버튼 제어(Win32 API, UI Automation) — Excel 비종속, 2026-07-24 신규 |

### 완성된 cwb_ 모듈

| 모듈 | 주요 기능 |
|---|---|
| `tpl_Test` | 전 모듈 자동화 테스트 (31개 프로시저) |

### 미완성

없음 — 전체 Phase 완료, 테스트 코드 작성 완료

---

## 다음 작업

### CWB 인수 정리 반영 상태 (2026-09-15)

- ✅ 소스 수정 완료 — am_Path · am_File · ref_Path · ref_File · tpl_Path · tpl_Test, CLAUDE.md · 매뉴얼 · LOG.md
- ✅ 교체 임포트 · `tpl_Test.Test_Path` 실행 — 전 항목 OK (2026-09-15 사용자 실행)
  - ⚠️ `UNC 경로` 결과 `H:\\서버명\공유폴더` — 판정은 OK(비어있지 않음)지만 실제로는 오동작
- ✅ 2차 소스 수정 — UNC 드라이브 제외·공유 루트 판정, 대체 드라이브 CWB 기준, `blnCreate` 신규, 테스트 강화
- ✅ 2차 반영 · `Test_Path` 전 항목 OK (2026-09-15 사용자 실행)
  - 실제 UNC(`\\localhost\C$`) · 매핑 드라이브(H: = `\\172.17.107.251\wf10084`) 항목 모두 실행되어 통과
  - 테스트 CWB 는 네트워크 드라이브(H:) 위치 — 실사용 환경과 동일 조건에서 검증됨
- ✅ `TROUBLESHOOTING.md` TS-V18 기록 (UNC 드라이브 부착 버그)
- ✅ `Test_DriveMapping` 신규 케이스 "없는 드라이브(Z:) → CWB 드라이브(C:)로 대체" OK (2026-09-15 사용자 실행) — 대체 기준 xlam → CWB 변경 검증 완료
- ⚠️ `GetPath()` 수동 확인 → `{cpath}\vba program template.xlsm` (소문자화) — 기존 LCase 역치환 블록 미삭제가 원인, 소스 수정 완료
- ✅ `am_File.bas` 재임포트 후 `GetPath()` 재확인 — `{cPath}\VBA PROGRAM Template.xlsm` 원본 대소문자 반환 확인 (2026-09-15 사용자 실행)
- ✅ 네트워크 드라이브 위치 CWB 에서 `ReplacePath` 검증 — 테스트 CWB 자체가 H:(네트워크 드라이브)
- ⬜ `GetPath` 나머지 수동 확인: 다중 선택(`;` 연결·항목별 치환), CWB 밖 경로 선택 시 토큰 미치환, 폴더 선택 모드
- ⬜ 기존 배포 CWB 의 구 시그니처 호출부 확인 — 구 호출은 타입 불일치 오류로 중단됨
  - `ReplacePath(p, ThisWorkbook.Path, ThisWorkbook.FullName)` → `ReplacePath(p)`
  - `GetPath(ThisWorkbook.Path, ...)` → `GetPath(...)`
  - `ResolvePath(p)` (tpl_Path 삭제) → `ReplacePath(p)`

**이 섹션 바로 다음 작업: `GetPath` 나머지 수동 확인 → 기존 CWB 구 시그니처 호출부 정리**

### am_Automation 반영 상태 (2026-07-24 기준)

- ✅ `am_Automation.bas` corelib.xlam VBE 임포트 완료 (사용자 확인)
- ✅ `ref_Automation.bas` CWB 워크북 VBE 임포트 완료 (사용자 확인)
- ✅ `corelib.xlam` VBE 참조에 "UIAutomationClient" 라이브러리 체크 완료 (사용자 확인)
- ⬜ **컴파일 체크** — `디버그 > VBAProject 컴파일`로 Early Binding 타입(`IUIAutomationElement` 등) 정상 인식 확인 (다음 세션에서 결과 확인)
- ⬜ **`ClickButtonByUIA` 실제 클릭 기능 테스트** — Windows 계산기 등으로 검증 예정:
  ```vba
  ?Application.Run("corelib.xlam!am_Automation.ClickButtonByUIA", "계산기", "1")
  ```
  `True` 반환 + 계산기 화면에 값 입력되면 정상. 버튼 이름이 안 맞으면 임시 `Diag_ListButtons` 스니펫(모듈에는 미포함, 대화 로그 참고)으로 실제 이름 확인 후 재시도
- ⬜ `ClickButtonByText`/`ClickButtonByRect`/`ForceActivateWindow` 도 실제 대상 창으로 수동 확인 필요 — `tpl_Test` 자동화 테스트 대상 아님
- ⬜ (아직 미반영) `am_Excel.bas`(키보드/마우스 섹션 제거된 버전), `ref_Excel.bas`(키보드/마우스 래퍼 제거된 버전) 교체 임포트 — am_Automation/ref_Automation만 추가 임포트했다면 기존 am_Excel/ref_Excel에 키보드/마우스 함수가 중복 존재할 수 있어 확인 필요

**바로 다음 작업: 컴파일 체크 → `ClickButtonByUIA` 계산기 테스트 → 결과에 따라 버튼 이름 재확인**

### 준비 완료 상태 (2026-06-09 기준)

- `corelib.xlam` 파일 생성 완료 (12개 am_ 모듈 임포트)
- `cwb_01.xlsm` 파일 생성 완료

### 테스트 실행 절차

1. `corelib.xlam` 과 `cwb_01.xlsm` 을 같은 폴더에 배치
2. `cwb_01.xlsm` 열기 → `Workbook_Open` 이 자동으로 corelib.xlam 로드
3. VBE Immediate 창에서 `tpl_Test.RunAllTests` 실행
4. 출력에서 `FAIL` 항목 확인 → 버그 수정

### 수정 완료 항목

| 항목 | 결과 |
|---|---|
| `GetSheetNames` 2D 배열 버그 | ✅ 수정 완료 — `Transpose` 제거, 1D 배열 직접 반환 (2026-06-09) |
| Property Get via Application.Run | ✅ 수정 완료 — `am_Core` Property Get 4개 → Public Function 변환 (2026-06-10) |
| `am_Range` 스칼라 래퍼 신규 | ✅ 추가 완료 — `GetUsedRange_IsValid/RowCount/ColCount`, `FindRange_IsValid/CellValue`, `FindCellsByColor_Count` (2026-06-10) |
| `Test_Range` 전체 통과 확인 | ✅ 런타임 검증 완료 (2026-06-10) |
| `BackupWorkbook` 폴더 미생성 버그 | ✅ 수정 완료 — `prv_MkFolder` 호출 추가 (2026-06-11) |

### 런타임 확인 필요 항목

| 항목 | 확인 포인트 |
|---|---|
| `SortTable` / `SortTableCustomList` | `Range("T_TestData")` 가 cwb_01.xlsm 의 테이블을 참조하는지 확인 (active WB 기준이므로 동작 예상) |
| `BackupSheet` 활성 워크북 변경 | ws.Copy 후 ActiveWorkbook 타이밍 — 코드 구조상 정상, 런타임 확인 필요 |
| `am_Error` 테스트 없음 | `ENABLE_ERROR_LOG=False` 기본값으로 WriteLog 무동작 — 테스트 생략 상태 |

### 이후 우선순위

1. **버그 수정** — RunAllTests 실행 중 발견된 FAIL 항목 처리
2. **신규 기능 추가** — 필요 시 SOURCES.md 미이식 항목 재검토
3. **실 CWB 적용** — cwb_01.xlsm 에서 실제 업무 함수 호출 검증

### 완료된 Phase 이력

| Phase | 내용 | 완료일 |
|---|---|---|
| Phase 1 | 기본 구조 (am_Core/Path/File/DB/Sheet/Table/Range/Format) | 2026-05-21 ~ 06-02 |
| Phase 2 | 신규 모듈 (am_Error, am_DB 확장) | 2026-06-04 |
| Phase 3 | am_Utils 신규 | 2026-06-05 |
| Phase 4 | am_Excel 완성 | 2026-06-05 |
| Phase 5 | tpl_Test 전 모듈 테스트 작성 | 2026-06-08 |

---

## tpl_Test 구조 요약 (2026-06-08 기준)

### 진입점

| 프로시저 | 설명 |
|---|---|
| `RunAllTests` | 전체 순차 실행 |
| `Setup_TestSheet` | `__TEST__` 시트 + `T_TestData` 테이블 자동 생성 |
| `Teardown_TestSheet` | `__TEST__` 삭제 + `_test_tmp` 폴더 정리 |

### 테스트 실행 순서

```
Test_Core → Test_Path → Test_File → Test_Utils
→ Setup_TestSheet
→ Test_Range → Test_Format → Test_Table → Test_Sheet → Test_Excel
→ Teardown_TestSheet
```

### 모듈별 테스트 커버리지

| 모듈 | 테스트 프로시저 | 주요 검증 항목 |
|---|---|---|
| am_Core | `Test_Core` | XlamPath/Version/IsReady/IsXlamLoaded, DPUpdate/Event/Calculate Off-On |
| am_Path | `Test_Path` (×6) | 고정토큰, 커스텀토큰, 절대경로, UNC, 드라이브매핑, 엣지케이스 |
| am_File | `Test_File` | GetExt, CheckFolderExistence, MkFolder 중첩, CheckFileExistence, DelFolder |
| am_Utils | `Test_Utils` (×5) | ConvertToArrData, IsArrayEmpty, IsValidFileName, CreateUniqueID, GenerateRandomCode, ConvertToExcelSerialDate, ExtractValues, EvaluateFormula |
| am_Range | `Test_Range` | GetUsedRange, FindRange(있음/없음), FindCellsByColor(있음/없음) |
| am_Format | `Test_Format` | CF 추가/삭제(Formula/ColorScale/DataBar), ValidationList, SetValidation, ClearValidation ※am_Utils.GetValidationType 교차검증 |
| am_Table | `Test_Table` | GetTableNames, IsTable, TblFindVals_MC/One/Rng, intOffset, AutoTableFilter, SortTable, SortTableCustomList, ChangeTableValue, AddTableRows, DelTableFilteredRows, AddArrayColumns, DelTableColumns |
| am_Sheet | `Test_Sheet` | GetSheetNames, VisibleAllSheets, HideAllSheetsExceptOne, SortSheets(역순→복원), BackupSheet, BackupWorkbook, SheetLock/SheetUnLock |
| am_Excel | `Test_Excel` | SetPrintPage, ExportSheetToCSV, ExportPDF |
| am_Error | — | ENABLE_ERROR_LOG=False 기본값으로 자동 테스트 불가 (수동 확인 필요) |

---

## 작업 방식

1. `MIGRATE.md` 에서 ⬜ 항목 선택
2. `refer/SOURCES.md` 에서 해당 모듈 이식 여부·수정 사항 확인
3. `refer/export/tpl_XXX.bas` Read
4. 기존 am_ 모듈 Read
5. 코드 작성
6. `MIGRATE.md` 해당 항목 ✅ 업데이트
7. `LOG.md` 변경 이력 추가

---

## 핵심 설계 원칙 (자주 잊는 것)

- **모듈 간 직접 호출 금지** → 필요하면 `prv_` 로 내부 구현
- **Late Binding** → `CreateObject(...)` 사용, 참조 추가 안 함
- **Optional 인수** → 반드시 필수 인수 뒤에 위치
- **CWB 인수** → 필요한 정보별 1개(`strCwbPath` 또는 `strCwbFile`), 맨 끝 Optional, ref_ 래퍼가 자동 입력 (CLAUDE.md §6)
- **색상 센티널** → `-1` (0은 검정으로 유효값)
- **배열 인덱스** → 항상 `LBound` / `UBound`
- **에러 패턴** → `On Error GoTo ErrHandler` + `MsgBox ... am_Core.AM_NAME`
- **CleanUp 패턴** → 개체 해제 필요 시 `GoTo CleanUp` 사용
