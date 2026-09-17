# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

교직원 급식비 신청·징수 자동화 시스템 — Google Sheets + Google Apps Script(HTMLService)로 만든 웹앱. 직원이 웹 화면에서 본인의 급식 가능일을 체크해 제출하면 시트에 반영되고 급식비가 계산되어 안내 메일이 자동 발송된다. 구현된 코드는 `gas/` 폴더 전체이며, 나머지 최상위 파일(`0.작업지시-1.txt`, `1.챗지피티-프롬프트.txt`, `2.claude-계획1/2.txt`, `3.*`, `*.xlsx`, `html 화면-1.png`)은 요구사항·계획·참고자료이지 실행 코드가 아니다.

**중요**: `.planning/PROJECT.md`는 현재 구현(Google Apps Script + OAuth)을 로컬 파이썬 웹서버 + SQLite + exe 패키징 방식으로 전환하는 계획을 담고 있다. 이 전환은 아직 코드로 구현되지 않았다(저장소에 Python 소스 없음) — 작업 전 어느 단계(GAS 유지보수 vs 신규 로컬 서버 구축)를 요청받은 것인지 확인할 것.

## Commands

이 프로젝트는 별도 빌드/린트/테스트 러너가 없다(`package.json` 없음). Google Apps Script는 [clasp](https://github.com/google/clasp) CLI로 관리하며, 모든 clasp 명령은 `gas/` 디렉터리(스크립트 프로젝트 루트, `.clasp.json` 위치)에서 실행한다.

```bash
cd gas
clasp login        # 최초 1회, 브라우저 OAuth 인증
clasp push          # 로컬 .gs/.html 변경사항을 바인딩된 Apps Script 프로젝트로 업로드
clasp pull          # Apps Script 편집기에서 직접 수정한 내용을 로컬로 내려받기
clasp open          # 브라우저에서 Apps Script 편집기 열기
```

- **자동 테스트 없음.** 검증은 (1) `clasp push` 후 스프레드시트에서 "급식비 관리" 메뉴의 "구조 점검"·"테스트 메일(본인)" 실행, (2) `2.claude-계획1.txt`에 정리된 수동 테스트 체크리스트(T1~T30)로 수행한다.
- `runInitialSetup`/`runGuidedSetup`처럼 시트·메일 발급 권한이 필요한 함수는 `clasp run`으로 원격 실행할 수 없다(미검증 앱의 민감 스코프 요청이 Google 정책상 차단됨) — 반드시 Apps Script 편집기에서 직접 실행해 권한을 승인해야 한다. 이 제약은 코드 버그가 아니라 구조적 한계이므로 우회를 시도하지 말 것.
- 웹앱 자동 배포(`DeployService.gs`의 `deployNow`)는 사본의 기본 GCP 프로젝트가 접근 제한된 경우 `Apps Script API has not been used in project ...` 오류로 실패하는 것이 정상이다 — 이 경우 Apps Script 편집기의 수동 "배포 → 새 배포 → 웹 앱"이 기본 경로다.

## Architecture

### 계층 구조

- **`Config.gs`** — 시트명(`SHEET_NAMES`), 열 인덱스(`DATA_COL`/`MEAL_COL`/`PERSONAL_COL`), 헤더 행 번호 등 전역 상수를 한 곳에 모은 단일 진실 소스. 시트 구조 변경 시 이 파일만 고치면 나머지 서비스 파일은 상수를 참조한다.
- **서비스 파일 분리** — `EmployeeService.gs`(직원 마스터/직원ID/동기화), `MealService.gs`(기간 조회·제출 저장·월 생성/마감·초기화), `BillingService.gs`(개인별내역 재생성), `EmailService.gs`(안내 메일), `AdminService.gs`(토큰 검증·관리자 상세·구조 자기진단), `Setup.gs`(초기 설정 2가지 경로), `DeployService.gs`(자가 배포). 각 파일은 자신의 도메인 시트만 다루고, 다른 시트 접근은 `Config.gs`의 헬퍼(`getSheet_`, `getSpreadsheet_`)를 통한다.
- **클라이언트-서버 연결** — `Code.gs`의 `doGet`이 `Index.html`을 템플릿으로 렌더링(관리자 토큰은 URL 쿼리 `?admin=`으로 템플릿에 주입). 클라이언트(`JavaScript.html`)는 `google.script.run`으로 `MealService.gs`/`EmployeeService.gs`/`AdminService.gs`의 함수를 직접 호출한다 — REST API나 별도 라우터가 없다. 응답은 `Utils.gs`의 `ok_(data)`/`fail_(code, message)`로 통일된 `{ok, data|error}` 형태를 반환한다.
- **관리자 메뉴(`onOpen`)** — 스프레드시트를 열 때 생성되는 "급식비 관리" 메뉴가 각 서비스 함수의 유일한 수동 진입점이다. 새 관리 기능을 추가하면 `Code.gs`의 `onOpen`과 대응하는 `menu*_` 래퍼 함수도 함께 추가해야 UI에 노출된다.
- **`onEdit`** — `학교급식` 시트의 년/월 셀(`MEAL_YEAR_CELL`/`MEAL_MONTH_CELL`)을 직접 수정하면 값은 막지 않고 토스트 경고만 띄운다(월 전환은 반드시 "월 마감" 메뉴를 거치도록 유도).

### 시트 = 데이터베이스

이 시스템에는 별도 DB가 없고 5개 시트가 스키마 역할을 한다(정확한 열 매핑은 `Config.gs` 참조):

| 시트 | 역할 |
|---|---|
| 데이터 | 직원 마스터 (순/직급/성명/이메일/직원ID/재직여부) |
| 학교급식 | 월별 신청 원장 — B1=년도, D1=월, F1=급식단가, H2:AL2=1~31일 체크, AM~AP=제출일시/비고/메일상태/직원ID |
| 개인별내역 | 학교급식 기준으로 스크립트가 매번 전체 재생성하는 출력 전용 시트 |
| 설정 | `Setup.gs`가 자동 생성. A:B=운영설정(KEY/VALUE), D:F=날짜별 급식가능일(가능/불가/주의/사유) |
| 제출로그 | 제출 이력 append-only 감사 로그 |

직원ID(`EMP001...`)가 `데이터`·`학교급식` 두 시트를 잇는 매칭 키다. 새 열을 추가하거나 헤더 행을 옮길 때는 `Config.gs`의 상수와 `Setup.gs`의 `ensure*Headers_` 함수를 함께 갱신해야 한다.

### 배포 모델의 제약 (설계에 영향을 주는 핵심 사실)

- 웹앱은 Google Sheets 템플릿 "사본 만들기" + `SetupWizard.html`/`Setup.gs`의 `runGuidedSetup()` 1회 실행으로 다른 학교에 재배포한다. exe나 별도 설치 프로그램은 만들지 않았다 — Google이 시트/메일 접근 권한을 앱 형태와 무관하게 사람이 최소 1회 브라우저에서 직접 승인하도록 강제하기 때문이다. 이 제약을 우회하는 방향의 코드 변경(예: 자동 OAuth 승인 시도)은 요청하지 않는 한 하지 않는다.
- 비밀값(`BANK_ACCOUNT`, `ACCOUNT_HOLDER`)은 코드가 아니라 스크립트 속성(`PropertiesService`)에 저장한다. `Config.gs`의 `getScriptProp_`/`setScriptProp_`을 거치지 않고 하드코딩하지 않는다.
- 관리자 상세표는 `AdminService.gs`의 토큰 검증(`checkAdminToken_`)을 통해서만 접근한다 — 일반 직원 화면에는 집계 카운트만 노출된다(개인정보 보호 경계).

더 자세한 설치 절차·관리자 메뉴 전체 목록·현재 알려진 제약은 [README.md](README.md)를, 함수 명세와 테스트 체크리스트(T1~T30)는 `2.claude-계획1.txt`를 참조한다.
