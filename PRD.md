# PRD (제품 요구사항 문서) — excel-dbapi

> **버전**: 1.0.0  
> **최종 수정일**: 2026-03-15  
> **상태**: v1.0.0 Stable (기능 완성)  
> **작성자**: Yeongseon Choe

---

## 1. 제품 개요

### 1.1 한 줄 요약

Excel(.xlsx) 파일을 데이터베이스처럼 SQL로 조회·삽입·수정·삭제할 수 있는 **PEP 249 (DB-API 2.0) 호환 Python 드라이버**.

### 1.2 비전

스프레드시트 데이터를 다루기 위해 별도의 데이터 변환 파이프라인이나 데이터베이스 마이그레이션 없이, 표준 Python DB-API 인터페이스를 통해 **익숙한 SQL 구문으로 Excel 파일에 즉시 접근**할 수 있는 세상을 만든다.

### 1.3 핵심 가치 제안

| 가치 | 설명 |
|------|------|
| **표준 호환** | PEP 249 DB-API 2.0 완전 준수 — `connect()`, `cursor()`, `execute()`, `fetchall()` 등 표준 인터페이스 |
| **SQL 기반 접근** | SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE 지원 |
| **제로 러닝 커브** | Python DB-API를 아는 개발자라면 추가 학습 없이 바로 사용 가능 |
| **듀얼 엔진** | openpyxl(기본, 셀 레벨 접근) 또는 pandas(DataFrame 기반) 엔진 선택 |
| **트랜잭션 안전** | 스냅샷 기반 commit/rollback으로 데이터 무결성 보장 |
| **원자적 저장** | OpenpyxlEngine은 tempfile + os.replace를 이용한 원자적 파일 저장 |
| **하위 호환** | sqlalchemy-excel의 핵심 Excel I/O 레이어로 활용 |

---

## 2. 대상 사용자 (페르소나)

### 2.1 데이터 분석가 (Data Analyst)

> "Excel 파일에 있는 데이터를 CSV로 변환하거나 DB에 넣지 않고, 바로 SQL로 조회하고 싶다."

- **사용 시나리오**: 주간 보고서 Excel에서 특정 조건의 데이터 추출
- **핵심 요구**: SELECT + WHERE + ORDER BY + LIMIT
- **기대 효과**: pandas나 openpyxl API 학습 없이 SQL로 즉시 분석

### 2.2 Python 개발자 (Python Developer)

> "기존 DB 추상화 레이어(SQLAlchemy)와 동일한 인터페이스로 Excel 데이터를 다루고 싶다."

- **사용 시나리오**: 레거시 시스템의 Excel 데이터를 DB 마이그레이션 전 단계로 활용
- **핵심 요구**: DB-API 2.0 표준 인터페이스, 파라미터 바인딩
- **기대 효과**: 기존 DB 코드 패턴 재활용, SQLAlchemy 연동 가능

### 2.3 QA 엔지니어 (QA Engineer)

> "테스트 데이터가 담긴 Excel 파일을 SQL로 조회해서 자동 검증하고 싶다."

- **사용 시나리오**: Excel 기반 테스트 데이터의 자동화된 검증
- **핵심 요구**: 정확한 조건 필터링, 프로그래밍 방식 접근
- **기대 효과**: 수동 스프레드시트 검사 제거

### 2.4 sqlalchemy-excel (다운스트림 라이브러리)

> "Excel I/O를 위한 표준화된 DB-API 레이어가 필요하다."

- **사용 시나리오**: SQLAlchemy 모델 기반 Excel 처리의 저수준 I/O 엔진
- **핵심 요구**: `connect()` → `cursor()` → SQL 실행, `workbook` 속성 접근
- **기대 효과**: `ExcelWorkbookSession`이 SQL 채널과 openpyxl 워크북 채널을 동시 제공
- **통합 방식**:
  - `ExcelWorkbookSession`이 excel-dbapi `connect()`를 래핑
  - `ExcelDbapiReader`가 excel-dbapi `cursor()`를 통해 SQL 기반 데이터 읽기
  - `ExcelTemplate`과 `ExcelExporter`가 `ExcelWorkbookSession.open()`으로 워크북 생성

---

## 3. 기능 요구사항

### 3.1 v1.0.0 현재 기능 (모두 완성)

#### 3.1.1 DB-API 2.0 인터페이스

| ID | 기능 | 상태 |
|----|------|------|
| FR-01 | `connect(file_path, engine, autocommit, create, data_only)` 모듈 레벨 생성자 | ✅ |
| FR-02 | `ExcelConnection`: `cursor()`, `commit()`, `rollback()`, `close()` | ✅ |
| FR-03 | `ExcelCursor`: `execute()`, `executemany()` | ✅ |
| FR-04 | `ExcelCursor`: `fetchone()`, `fetchall()`, `fetchmany(size)` | ✅ |
| FR-05 | `description` 속성 (PEP 249 7-튜플 형식) | ✅ |
| FR-06 | `rowcount`, `lastrowid`, `arraysize` 속성 | ✅ |
| FR-07 | 컨텍스트 매니저 (`with` 문) 지원 | ✅ |
| FR-08 | PEP 249 예외 계층 구조 | ✅ |
| FR-09 | `check_closed` 데코레이터 (닫힌 커넥션/커서 보호) | ✅ |
| FR-10 | 모듈 레벨 상수: `apilevel="2.0"`, `threadsafety=1`, `paramstyle="qmark"` | ✅ |

#### 3.1.2 SQL 지원

| ID | 기능 | 상태 |
|----|------|------|
| FR-11 | `SELECT col1, col2 FROM Sheet WHERE ... ORDER BY col ASC/DESC LIMIT N` | ✅ |
| FR-12 | `SELECT *` — 전체 컬럼 조회 | ✅ |
| FR-13 | `INSERT INTO Sheet (col1, col2) VALUES (?, ?)` — 컬럼 지정/미지정 삽입 | ✅ |
| FR-14 | `UPDATE Sheet SET col = ? WHERE ...` — 다중 SET 할당 | ✅ |
| FR-15 | `DELETE FROM Sheet WHERE ...` — 조건부/전체 삭제 | ✅ |
| FR-16 | `CREATE TABLE SheetName (col1, col2)` — 새 워크시트 생성 | ✅ |
| FR-17 | `DROP TABLE SheetName` — 워크시트 제거 | ✅ |
| FR-18 | WHERE 절: `=`, `==`, `!=`, `<>`, `>`, `>=`, `<`, `<=` 비교 연산자 | ✅ |
| FR-19 | WHERE 절: `AND`, `OR` 논리 연산자 | ✅ |
| FR-20 | 파라미터 바인딩: `?` 플레이스홀더 (`qmark` paramstyle) | ✅ |
| FR-21 | ORDER BY 절: ASC/DESC 정렬 | ✅ |
| FR-22 | LIMIT 절: 결과 행 수 제한 | ✅ |

#### 3.1.3 엔진 아키텍처

| ID | 기능 | 상태 |
|----|------|------|
| FR-23 | OpenpyxlEngine (기본 엔진): 셀 레벨 읽기/쓰기 | ✅ |
| FR-24 | PandasEngine (대안 엔진): DataFrame 기반 처리 | ✅ |
| FR-25 | 연결 시 엔진 선택: `engine="openpyxl"` 또는 `engine="pandas"` | ✅ |
| FR-26 | `create=True` 플래그: 파일 미존재 시 자동 생성 | ✅ |
| FR-27 | `data_only=True` 플래그: 수식 캐시 값 읽기 제어 | ✅ |
| FR-28 | `workbook` 속성: openpyxl Workbook 객체 직접 접근 (OpenpyxlEngine만) | ✅ |

#### 3.1.4 트랜잭션 관리

| ID | 기능 | 상태 |
|----|------|------|
| FR-29 | `autocommit=True` (기본): 쓰기 작업 즉시 저장 | ✅ |
| FR-30 | `autocommit=False`: 수동 `commit()` / `rollback()` | ✅ |
| FR-31 | 스냅샷 기반 상태 관리 (OpenpyxlEngine: BytesIO, PandasEngine: deepcopy) | ✅ |
| FR-32 | `commit()` → 디스크 저장 + 새 스냅샷 생성 | ✅ |
| FR-33 | `rollback()` → 스냅샷 복원 (autocommit 모드에서는 `NotSupportedError`) | ✅ |
| FR-34 | `executemany()` + `autocommit=False`: 원자적 배치 — 실패 시 전체 롤백 | ✅ |
| FR-35 | 원자적 파일 저장: tempfile + `os.replace()` (OpenpyxlEngine) | ✅ |

#### 3.1.5 SQL 파서

| ID | 기능 | 상태 |
|----|------|------|
| FR-36 | 커스텀 재귀 하강 파서 (외부 의존성 없음) | ✅ |
| FR-37 | 테이블명 대소문자 무시 조회 (`table.lower()` 매핑) | ✅ |
| FR-38 | 비인용 테이블명 사용 (`Sheet1`, NOT `"Sheet1"`) | ✅ |
| FR-39 | 값 파싱: NULL, 문자열(따옴표), 정수, 실수 | ✅ |
| FR-40 | CSV 분리기: 따옴표 내 쉼표 처리 | ✅ |

---

## 4. 비기능 요구사항

### 4.1 성능

| ID | 요구사항 | 목표 |
|----|----------|------|
| NFR-01 | openpyxl 엔진: 10MB 이하 파일 초기 로드 | < 2초 |
| NFR-02 | pandas 엔진: 50MB 이하 파일 지원 | 메모리 < 500MB |
| NFR-03 | SELECT 쿼리 실행 (10,000행 이하) | < 1초 |
| NFR-04 | 읽기 전용 작업 시 메모리 사용 최적화 | < 파일 크기 2배 |

### 4.2 호환성

| ID | 요구사항 |
|----|----------|
| NFR-05 | Python 3.10 이상 지원 |
| NFR-06 | openpyxl 3.1.0 이상 |
| NFR-07 | pandas 2.0.0 이상 |
| NFR-08 | PEP 249 준수 테스트 통과 |
| NFR-09 | .xlsx 파일 형식 (Excel 2007+) 지원 |

### 4.3 품질

| ID | 요구사항 | 목표 |
|----|----------|------|
| NFR-10 | 코드 커버리지 | > 80% |
| NFR-11 | 공개 API 타입 힌트 | 100% |
| NFR-12 | mypy 정적 타입 검사 통과 | 0 에러 |
| NFR-13 | ruff/black 린팅/포매팅 통과 | 0 경고 |

### 4.4 신뢰성

| ID | 요구사항 |
|----|----------|
| NFR-14 | 손상된 Excel 파일에 대한 우아한 에러 처리 |
| NFR-15 | 빈 시트에 대한 안전한 처리 (크래시 없음) |
| NFR-16 | 쓰기 실패 시 원본 파일 보존 (원자적 저장) |
| NFR-17 | 일반적인 오류 상황에 대한 명확한 에러 메시지 |

### 4.5 사용성

| ID | 요구사항 |
|----|----------|
| NFR-18 | DB-API 2.0 관례를 따르는 발견 가능한 API |
| NFR-19 | 실행 가능한(actionable) 에러 메시지 |
| NFR-20 | 빠른 시작 가이드 포함 |
| NFR-21 | 일반적인 사용 사례를 다루는 예제 제공 |

---

## 5. sqlalchemy-excel과의 상호운용성

### 5.1 관계 정의

excel-dbapi는 [sqlalchemy-excel](https://github.com/yeongseon/sqlalchemy-excel)의 **전면 의존성(full dependency)**이다. sqlalchemy-excel은 excel-dbapi를 핵심 Excel I/O 레이어로 사용한다.

```
┌─────────────────────────────────────────────┐
│          sqlalchemy-excel                    │
│  ┌─────────────────┐  ┌──────────────────┐  │
│  │ExcelWorkbook    │  │ExcelDbapiReader  │  │
│  │Session          │  │                  │  │
│  │                 │  │  SQL 기반 데이터  │  │
│  │  SQL 채널       │  │  읽기            │  │
│  │  + 워크북 채널  │  │                  │  │
│  └────────┬────────┘  └────────┬─────────┘  │
│           │                    │             │
│           └────────┬───────────┘             │
│                    │                         │
│                    ▼                         │
│           ┌────────────────┐                 │
│           │  excel-dbapi   │                 │
│           │  connect()     │                 │
│           └────────────────┘                 │
└─────────────────────────────────────────────┘
```

### 5.2 통합 포인트

| sqlalchemy-excel 모듈 | excel-dbapi 사용 방식 |
|----------------------|----------------------|
| `ExcelWorkbookSession` | `connect(path, engine="openpyxl", create=True)` → SQL 실행 + `conn.workbook` 접근 |
| `ExcelDbapiReader` | `conn.cursor()` → `cursor.execute("SELECT ...")` → `cursor.fetchall()` |
| `ExcelTemplate` | `ExcelWorkbookSession.open()` → 워크북 생성 |
| `ExcelExporter` | `ExcelWorkbookSession.open()` → 쿼리 결과 Excel 내보내기 |

### 5.3 계약 (Contract)

excel-dbapi는 다음 계약을 sqlalchemy-excel에 보장한다:

1. **`connect()` 함수**는 `ExcelConnection` 객체를 반환한다
2. **`conn.workbook`** 속성은 openpyxl `Workbook` 객체를 반환한다 (openpyxl 엔진)
3. **`conn.cursor()`**는 표준 DB-API 2.0 `Cursor`를 반환한다
4. **`cursor.execute("SELECT * FROM SheetName")`**은 `ExecutionResult`를 생성한다
5. **`cursor.description`**은 PEP 249 7-튜플 시퀀스를 반환한다
6. **`create=True`** 플래그는 파일이 없을 때 빈 워크북을 생성한다
7. **예외**는 PEP 249 계층을 따른다

---

## 6. 제약 사항 및 한계

### 6.1 현재 한계

| 한계 | 설명 |
|------|------|
| SQL 서브셋만 지원 | 전체 SQL-92가 아닌, 기본 CRUD + DDL 서브셋만 지원 |
| 단일 파일 연산 | 여러 Excel 파일 간 JOIN 불가 |
| 동시 접근 불가 | 파일 레벨 잠금 미구현, 동시 쓰기 시 데이터 손실 위험 |
| Excel 고유 기능 미지원 | 수식, 차트, 서식, 피벗 테이블 등 비데이터 요소 조작 불가 |
| .xlsx만 지원 | CSV, ODS, .xls 등 다른 형식 미지원 |
| PandasEngine 포맷 손실 | PandasEngine으로 저장 시 서식, 차트, 수식이 제거됨 |
| 복합 WHERE 제한 | 괄호 그룹핑 미지원 (`(A AND B) OR C` 불가) |
| JOIN 미지원 | 시트 간 JOIN 미구현 |

### 6.2 의도적 비범위 (Non-Goals)

- 저장 프로시저, 트리거, 뷰, 인덱스 등 완전한 RDBMS 기능
- 네트워크 프로토콜 기반 접근 (원격 파일)
- 실시간 협업 또는 동시성 제어
- .xlsx 이외의 파일 형식 지원

---

## 7. 로드맵

### 7.1 완료된 마일스톤

| 버전 | 목표 | 상태 |
|------|------|------|
| v0.1.x | 기본 읽기 전용 지원 (SELECT, DB-API 인터페이스) | ✅ 완료 |
| v0.2.x | 쓰기 작업 & DDL (INSERT, CREATE TABLE, DROP TABLE) | ✅ 완료 |
| v0.3.x | 데이터 수정 (UPDATE, DELETE, 트랜잭션 시뮬레이션) | ✅ 완료 |
| v0.4.x | 고급 SQL (ORDER BY, LIMIT, 확장 WHERE) | ✅ 완료 |
| v1.0.0 | 프로덕션 릴리스 (기능 완성, 테스트, 문서화) | ✅ 완료 |

### 7.2 향후 계획 (v2.0.x)

| 기능 | 우선순위 | 설명 |
|------|----------|------|
| SQLAlchemy Dialect | 높음 | `create_engine("excel:///path.xlsx")` 지원 |
| 멀티시트 JOIN | 중간 | `SELECT a.id, b.name FROM Sheet1 a JOIN Sheet2 b ON a.id = b.id` |
| Polars 엔진 | 중간 | pandas 대안으로 Polars DataFrame 엔진 추가 |
| 비동기 쿼리 | 낮음 | `async` / `await` 기반 비동기 지원 |
| 대용량 파일 최적화 | 중간 | 스트리밍 읽기, 청크 처리 |
| 원격 파일 지원 | 낮음 | S3, HTTP URL 등에서 직접 읽기 |
| 서브쿼리 | 낮음 | `SELECT * FROM Sheet1 WHERE id IN (SELECT ...)` |
| GROUP BY / HAVING | 중간 | 집계 함수와 그룹핑 지원 |
| 함수 지원 | 낮음 | `COUNT()`, `SUM()`, `AVG()`, `MIN()`, `MAX()` |

---

## 8. 리스크 분석

### 8.1 기술적 리스크

| 리스크 | 확률 | 영향 | 완화 방안 |
|--------|------|------|-----------|
| Excel 파일 형식 제한 | 중 | 높 | 파일 크기/복잡도 한계 문서화, 명확한 에러 메시지 |
| SQL 파서 복잡도 증가 | 높 | 높 | 서브셋 SQL만 유지, 단계적 기능 추가 |
| 대용량 파일 성능 저하 | 높 | 중 | 스트리밍 읽기 구현, 벤치마크 제공 |
| 쓰기 작업 데이터 손상 | 중 | 높 | 원자적 저장 (tempfile + os.replace), 스냅샷 기반 롤백 |
| 트랜잭션 시뮬레이션 불일치 | 중 | 높 | 한계 명확히 문서화, DB-API 2.0 시맨틱스 준수 |

### 8.2 프로젝트 리스크

| 리스크 | 확률 | 영향 | 완화 방안 |
|--------|------|------|-----------|
| 범위 확대 (scope creep) | 높 | 중 | 엄격한 로드맵 유지, 고급 기능은 v2.0+로 연기 |
| 유지보수 대역폭 | 중 | 높 | 모듈식 코드 구조, 커뮤니티 기여 장려 |
| 호환성 파괴 변경 | 중 | 중 | 시맨틱 버저닝, 폐지(deprecation) 경고, 마이그레이션 가이드 |

### 8.3 의존성 리스크

| 리스크 | 확률 | 영향 | 완화 방안 |
|--------|------|------|-----------|
| openpyxl/pandas API 변경 | 낮 | 높 | 의존성 버전 고정, 업스트림 릴리스 모니터링, 자동 테스트 |
| Python 버전 비호환 | 낮 | 중 | 3.10+ 지원, 모든 지원 버전에서 테스트 |
| sqlalchemy-excel 인터페이스 변경 | 낮 | 중 | 명시적 계약(contract) 유지, 통합 테스트 |

---

## 9. 버전 정책

### 9.1 시맨틱 버저닝 (SemVer 2.0.0)

```
MAJOR.MINOR.PATCH (예: 1.0.0, 2.1.3)

MAJOR — 호환성을 깨뜨리는 API 변경
MINOR — 하위 호환 기능 추가
PATCH — 하위 호환 버그 수정
```

### 9.2 릴리스 기준

| 유형 | 기준 |
|------|------|
| PATCH (x.x.Z) | 버그 수정만, API 변경 없음, 체인지로그 필수 |
| MINOR (x.Y.0) | 하위 호환 신기능, 폐지 기능 경고 포함 가능 |
| MAJOR (X.0.0) | 호환성 파괴 변경, 마이그레이션 가이드 필수 |

### 9.3 폐지 정책

1. **공지**: 릴리스 노트에 폐지 안내
2. **경고 추가**: 폐지된 기능 사용 시 `DeprecationWarning`
3. **유지**: 최소 1개 마이너 버전 동안 유지
4. **제거**: 다음 메이저 버전에서 제거

### 9.4 지원 정책

| 버전 | 지원 범위 |
|------|-----------|
| 최신 메이저 | 활발한 개발 + 보안 패치 |
| 이전 메이저 | 보안 패치만 (새 메이저 출시 후 6개월) |
| 그 이전 | 지원 종료 |

---

## 10. 성공 지표

| 지표 | 목표 | 현재 |
|------|------|------|
| 테스트 통과율 | 100% | ✅ 85/85 통과 |
| 코드 커버리지 | > 80% | ✅ 달성 |
| PEP 249 준수 | 완전 | ✅ 모든 필수 인터페이스 구현 |
| SQL 문 지원 | 6종 | ✅ SELECT, INSERT, UPDATE, DELETE, CREATE, DROP |
| 엔진 지원 | 2종 | ✅ openpyxl, pandas |
| 문서화 | 완전 | ✅ README, USAGE, DEVELOPMENT, ROADMAP, AGENTS, PRD, ARCH, TDD |
| PyPI 배포 | 완료 | ✅ `pip install excel-dbapi` |
| sqlalchemy-excel 통합 | 완료 | ✅ 전면 의존성으로 연동 |

---

## 부록

### A. 용어 정의

| 용어 | 정의 |
|------|------|
| DB-API 2.0 | Python 데이터베이스 API 명세 (PEP 249) |
| Sheet | Excel 워크시트 — 데이터베이스의 테이블에 해당 |
| Engine | 쿼리 실행의 기반 구현체 (openpyxl 또는 pandas) |
| PEP | Python Enhancement Proposal |
| SemVer | Semantic Versioning — 버전 번호 체계 (MAJOR.MINOR.PATCH) |
| qmark | DB-API paramstyle — `?` 플레이스홀더 사용 |

### B. 참조

- [PEP 249 — Python Database API Specification v2.0](https://peps.python.org/pep-0249/)
- [openpyxl 공식 문서](https://openpyxl.readthedocs.io/)
- [pandas 공식 문서](https://pandas.pydata.org/)
- [Semantic Versioning 2.0.0](https://semver.org/)
- [sqlalchemy-excel GitHub](https://github.com/yeongseon/sqlalchemy-excel)

### C. 문서 변경 이력

| 버전 | 날짜 | 변경 사항 |
|------|------|-----------|
| 1.0 | 2026-01-21 | 초기 PRD 작성 (v0.1.x 기준) |
| 1.1 | 2026-01-21 | v0.4.x 완료 상태 반영 |
| 2.0 | 2026-03-15 | v1.0.0 완성 상태 반영, 전면 재작성. sqlalchemy-excel 상호운용성, 듀얼 엔진 아키텍처, 트랜잭션 관리, 비기능 요구사항, 로드맵, 리스크 분석, 버전 정책 추가 |
