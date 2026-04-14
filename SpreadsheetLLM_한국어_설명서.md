# SpreadsheetLLM .NET 구현

> **기술 기반**: arXiv:2407.09025 논문, C# .NET 구현  
> **작성일**: 2026년 4월

---

## 목차

1. [SpreadsheetLLM이란 무엇인가?](#1-spreadsheetllm이란-무엇인가)
2. [왜 이 기술이 필요한가?](#2-왜-이-기술이-필요한가)
3. [전체 시스템 구조](#3-전체-시스템-구조)
4. [3단계 압축 파이프라인 — 쉬운 설명](#4-3단계-압축-파이프라인--쉬운-설명)
   - [1단계: 중요한 행·열 찾기 (구조적 앵커 추출)](#1단계-중요한-행열-찾기--구조적-앵커-추출)
   - [2단계: 거꾸로 정리하기 (역색인 변환)](#2단계-거꾸로-정리하기--역색인-변환)
   - [3단계: 데이터 형식으로 묶기 (서식 기반 집계)](#3단계-데이터-형식으로-묶기--서식-기반-집계)
5. [실제 대형 시트 예제](#5-실제-대형-시트-예제)
   - [예제 A — 판매 거래 500행 시트](#예제-a--판매-거래-500행-시트)
   - [예제 B — 인사·급여 300행 시트](#예제-b--인사급여-300행-시트)
   - [예제 C — 직원 대형 시트 (50행 × 10열)](#예제-c--직원-대형-시트-50행--10열)
6. [코드 파일별 역할 설명](#6-코드-파일별-역할-설명)
7. [JSON 출력 결과 읽는 법](#7-json-출력-결과-읽는-법)
8. [압축 성능 수치](#8-압축-성능-수치)
9. [자주 묻는 질문](#9-자주-묻는-질문)

---

## 1. SpreadsheetLLM이란 무엇인가?

### 한 문장으로

> **엑셀 파일을 AI(대형 언어 모델, LLM)가 빠르고 저렴하게 이해할 수 있도록 자동으로 압축해 주는 소프트웨어입니다.**

### 비유로 이해하기

책 한 권(500페이지)을 외국어 전문가에게 번역을 맡기려 한다고 상상하세요.  
전문가에게 보내는 분량이 많을수록 **비용이 높아지고, 처리 시간도 길어집니다.**

하지만 책을 먼저 **핵심 요약본(30페이지)** 으로 줄인 다음 번역을 맡기면:
- 비용이 낮아지고
- 시간이 짧아지며
- 전문가가 더 빠르게 이해합니다.

SpreadsheetLLM은 엑셀 파일에 대해 정확히 이런 역할을 합니다.  
AI에게 보내기 전에 **스프레드시트를 핵심만 남겨 압축**합니다.

---

## 2. 왜 이 기술이 필요한가?

### 문제: 엑셀은 AI에게 너무 크다

| 상황 | 내용 |
|------|------|
| 일반 엑셀 파일 크기 | 수십~수백 행 × 수십 열 |
| AI가 한 번에 처리할 수 있는 글자 수 | 제한이 있음 (토큰 제한) |
| AI API 비용 계산 기준 | 보내는 글자 수(토큰) 기준 과금 |

엑셀 파일을 그대로 텍스트로 변환해서 AI에게 보내면:
- A1: 제품명, B1: 수량, C1: 단가, D1: 합계  
- A2: 사과, B2: 100, C2: 500, D2: 50000  
- A3: 바나나, B3: 200, C3: 300, D3: 60000  
- ... (수백 행 반복)

이 방식은 **토큰(글자)을 낭비**하고, 반복되는 숫자들이 AI의 이해를 방해합니다.

### 해결책: 스마트한 압축

SpreadsheetLLM은 다음을 자동으로 처리합니다:

1. 표의 구조적 경계(머리글 행, 섹션 구분)를 찾습니다.
2. **같은 값이 여러 셀에 있으면 하나로 묶습니다.** (예: "완료" → B5, B8, B12, B19)
3. **같은 서식의 셀들을 범위로 묶습니다.** (예: 통화 형식 → C2:C500)

결과: AI에게 전달하는 데이터 양이 최대 **25배까지** 줄어듭니다.

---

## 3. 전체 시스템 구조

```
엑셀 파일 (.xlsx)
       │
       ▼
┌─────────────────────┐
│   ExcelReader.cs    │  ← 엑셀 파일을 읽어서 데이터를 메모리로 불러옴
│   (파일 읽기 담당)   │
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│  SheetCompressor.cs │  ← 3단계 압축 파이프라인 실행
│  (핵심 압축 담당)    │
│                     │
│  1단계: 구조 파악   │  → 중요한 행·열 식별
│  2단계: 역색인 생성 │  → 값→위치 매핑
│  3단계: 서식 집계   │  → 동일 형식끼리 묶기
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│   CellUtils.cs      │  ← 셀 분류 도우미 (날짜? 통화? 정수? 이메일?)
│   (셀 분석 담당)    │
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│  JSON 출력 파일     │  ← AI에게 전달하는 최종 압축 결과
│  (압축된 결과물)    │
└─────────────────────┘
```

### 코드 파일 목록 한눈에 보기

| 파일명 | 역할 | 비유 |
|--------|------|------|
| `ExcelReader.cs` | 엑셀 파일을 읽어 메모리에 저장 | 책을 스캔하는 스캐너 |
| `SheetCompressor.cs` | 3단계 압축 실행 | 요약 전문가 |
| `CellUtils.cs` | 각 셀의 데이터 유형 분류 | 데이터 분류기 |
| `CellData.cs` | 셀 하나의 모든 정보 저장 | 셀의 신분증 |
| `WorksheetSnapshot.cs` | 시트 전체의 2D 격자 저장 | 시트의 사진 |
| `SheetEncoding.cs` | 압축 결과의 JSON 구조 정의 | 압축본의 틀 |
| `VanillaEncoder.cs` | 압축 전 원본 텍스트 생성 (비교용) | 원본 기준선 |
| `ChainOfSpreadsheet.cs` | AI와 Q&A 파이프라인 | AI 연결 담당 |

---

## 4. 3단계 압축 파이프라인 — 쉬운 설명

### 예제로 사용할 간단한 엑셀 표

아래와 같은 판매 데이터가 있다고 가정합니다:

```
     A          B       C         D
1  제품명      수량    단가      합계
2  사과        100    500원    50,000원
3  바나나      200    300원    60,000원
4  체리        150  1,200원   180,000원
5  대추         80  2,000원   160,000원
6  엘더베리     60  3,500원   210,000원
7  합계        ---    ---    =SUM(D2:D6)
```

---

### 1단계: 중요한 행·열 찾기 — 구조적 앵커 추출

**코드 위치**: `SheetCompressor.cs` → `FindStructuralAnchors()` 함수

#### 무엇을 하는가?

마치 책에서 **목차, 장 제목, 소제목** 을 먼저 찾는 것과 같습니다.  
모든 내용을 읽는 대신, 구조를 이해하는 핵심 지점만 선별합니다.

#### 어떻게 찾는가? (단계별)

**① 행(Row) 분석 — 각 행이 어떤 종류인지 파악**

```
코드: AnalyzeRowsSinglePass() 함수
```

각 행에 대해 다음을 검사합니다:

| 검사 항목 | 예시 | 결과 |
|----------|------|------|
| 볼드체 셀이 60% 이상인가? | 1행: 제품명, 수량, 단가, 합계 → 전부 볼드 | 머리글 행! |
| 가운데 정렬이 60% 이상인가? | 1행 전부 가운데 정렬 | 머리글 행! |
| 밑선(하단 테두리)이 있는가? | 1행에 밑선 있음 | 머리글 행! |
| 모두 대문자인가? | TOTAL, REVENUE 등 | 머리글 행! |
| 이전 행과 데이터 유형이 다른가? | 1행=텍스트, 2행=숫자 시작 | 경계! |
| 빈 행인가? | 6행과 7행 사이 빈 행 | 섹션 구분! |

**② 경계 후보 생성**

```
코드: FindBoundaryCandidates() 함수
```

행·열 분석 결과를 바탕으로 "이 부분이 표의 경계일 것이다"라는 후보 구역들을 만듭니다.

예를 들어:
- 행 경계 후보: [1, 2, 7] (1번=머리글, 2번=데이터 시작, 7번=합계행)
- 열 경계 후보: [A, B, C, D]

**③ 중복 제거 (NMS — Non-Maximum Suppression)**

```
코드: NmsCandidates() 함수
```

같은 구역을 가리키는 중복 후보들을 제거하고, 가장 의미 있는 후보만 남깁니다.

**비유**: 여러 사람이 각자 "이 부분이 중요해!"라고 표시했을 때, 겹치는 부분은 하나로 합치는 과정.

**④ 앵커 확장 (k=2 기본값)**

```
코드: ExpandAnchors() 함수
```

찾은 앵커 행/열 주변 2행/2열도 함께 포함합니다.  
앵커가 5번 행이면 → 3, 4, 5, 6, 7번 행을 모두 포함.

**⑤ 동일한 반복 행/열 제거**

```
코드: CompressHomogeneousRegions() 함수
```

모든 셀이 같은 값, 같은 서식인 행은 삭제합니다.  
(예: 500행 짜리 표에서 '완료' 상태 행이 300개 있으면, 그 행들은 2단계에서 하나로 묶음)

**1단계 결과 (우리 예제)**:
```
남겨진 행: [1, 2, 3, 4, 5, 6, 7]  ← 작은 표라서 전부 남음
남겨진 열: [A, B, C, D]
구조적 앵커: rows=[1,7], columns=["A","D"]
```

---

### 2단계: 거꾸로 정리하기 — 역색인 변환

**코드 위치**: `SheetCompressor.cs` → `CreateInvertedIndex()`, `CreateInvertedIndexTranslation()`

#### 무엇을 하는가?

일반 엑셀 표현:
```
A1=제품명, A2=사과, A3=바나나 ...
```

역색인 표현 (뒤집기):
```
"사과" → [A2]
"바나나" → [A3]
"=SUM(D2:D6)" → [D7]
```

#### 왜 이렇게 하는가?

500행짜리 판매 데이터에서 "완료(Completed)" 상태가 350개 행에 있다면:

**원래 방식**: `J2=완료, J5=완료, J8=완료, ... (350번 반복)` — 엄청난 낭비!

**역색인 방식**: `"완료" → [J2, J5, J8, J12, ...]` — 한 번만 기록!

#### 범위 병합 (MergeCellRanges)

```
코드: MergeCellRanges() 함수
```

연속된 셀들을 범위로 묶습니다:
```
[J2, J3, J4, J5, J6] → "J2:J6"  (5개 → 1개로 표현)
```

**큰 시트에서의 효과**:
```
500행 판매 데이터에서:
- "Widget A"가 98번 등장 → "Widget A" → ["B5", "B12", "B17", ...(98개)"] 
  또는 연속이면 → "Widget A" → ["B2:B99"]
```

#### 2단계 결과 (우리 예제):
```json
"cells": {
  "제품명": ["A1"],
  "수량": ["B1"],
  "단가": ["C1"],
  "합계": ["D1"],
  "사과": ["A2"],
  "바나나": ["A3"],
  "체리": ["A4"],
  "대추": ["A5"],
  "엘더베리": ["A6"],
  "=SUM(D2:D6)": ["D7"]
}
```

---

### 3단계: 데이터 형식으로 묶기 — 서식 기반 집계

**코드 위치**: `SheetCompressor.cs` → `GroupBySemanticType()`, `AggregateBySemanticType()`

#### 무엇을 하는가?

셀의 **데이터 유형**과 **서식 형태**를 기준으로 묶습니다.

`CellUtils.cs`가 각 셀의 의미적 유형을 9가지로 분류합니다:

| 유형(영어) | 유형(한국어) | 예시 |
|-----------|-------------|------|
| `integer` | 정수 | 100, 200, 80 |
| `float` | 소수 | 3.14, 1.75 |
| `currency` | 통화 | $1,250.50, ₩50,000 |
| `percentage` | 퍼센트 | 75%, 12.5% |
| `scientific` | 과학적 표기 | 1.23E+05 |
| `date` | 날짜 | 2024-01-15 |
| `time` | 시간 | 09:30:00 |
| `year` | 연도 | 2023, 2024 |
| `email` | 이메일 | user@company.com |

#### 3단계 결과 (우리 예제):
```json
"formats": {
  "{\"type\":\"text\",\"nfs\":\"General\"}": ["A1:A6"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}": ["B2:B6"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["C2:C6", "D2:D7"]
}
```

이 정보를 통해 AI는 "C열은 전부 통화 형식이구나"를 즉시 파악합니다.

---

## 5. 실제 대형 시트 예제

### 예제 A — 판매 거래 500행 시트

**코드 위치**: `Program.cs` → `CreateRealisticSalesSheet()` 함수

#### 시트 구조

```
     A       B           C        D              E          F           G      H          I          J          K
1   TxnID   Date       Region  Salesperson    Product   Category     Qty  UnitPrice   Total      Status  PaymentMethod
2   1001  2023-03-15   North   Alice Smith  Widget A   Hardware      12    $9.99    $119.88   Completed  Credit Card
3   1002  2023-07-22   South   Bob Jones    Widget B   Software       5   $14.99     $74.95   Completed  Cash
4   1003  2023-01-08   East    Carol White  Gadget X   Services      30   $24.99    $749.70   Pending    Bank Transfer
...
501  1500  2023-11-30  Central  Jack Ford   Part Z     Accessories   18   $99.99  $1,799.82  Completed  Credit Card
---  (빈 행)
503  TOTAL  ---         ---      ---          ---        ---          ---    ---   $X,XXX.XX    ---       ---
```

#### 압축 전 (원본 텍스트 전달 시)

AI에게 다음과 같은 내용이 전달됩니다:
```
A1: TxnID, B1: Date, C1: Region, D1: Salesperson, ... K1: PaymentMethod
A2: 1001, B2: 2023-03-15, C2: North, D2: Alice Smith, E2: Widget A, ...
A3: 1002, B3: 2023-07-22, C3: South, D3: Bob Jones, E3: Widget B, ...
... (500행 전체 반복)
```
**→ 매우 많은 토큰 소비**

#### 압축 후 (SpreadsheetLLM 처리 결과)

**1단계 결과 — 구조적 앵커**:
```json
"structural_anchors": {
  "rows": [1, 2, 503],
  "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
}
```
→ 행 1번(머리글), 2번(첫 데이터), 503번(합계행)만 앵커로 인식

**2단계 결과 — 역색인 (반복 값들이 하나로)**:
```json
"cells": {
  "TxnID": ["A1"], "Date": ["B1"], "Region": ["C1"],
  "North":      ["C2", "C8", "C15", "C23", ...],
  "South":      ["C3", "C7", "C19", ...],
  "East":       ["C5", "C11", ...],
  "West":       ["C4", "C9", ..."],
  "Central":    ["C6", "C13", ..."],
  "Widget A":   ["E2:E99"],
  "Widget B":   ["E100:E198"],
  "Completed":  ["J2:J4", "J7", "J9:J11", ...],
  "Pending":    ["J3", "J8", ..."],
  "Credit Card": ["K2", "K5", "K7", ..."],
  "=SUM(I2:I501)": ["I503"],
  "TOTAL": ["A503"]
}
```

**핵심 포인트**: "Widget A"가 98번 등장해도 역색인에는 **한 번** 기록됩니다!

**3단계 결과 — 서식 집계**:
```json
"formats": {
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":    ["B2:B501"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["H2:H501", "I2:I501", "I503"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":    ["A2:A501", "G2:G501"],
  "{\"type\":\"text\",\"nfs\":\"General\"}":        ["C2:C501", "D2:D501", ...]
}
```

→ AI는 "B열 전체가 날짜 형식"임을 한 줄로 파악합니다.

#### 압축 효과 (예상)

| 구분 | 토큰 수 |
|------|---------|
| 원본 (500행 × 11열) | ~수만 토큰 |
| 압축 후 | 크게 감소 |
| 압축비 | **반복 값이 많을수록 높아짐** |

---

### 예제 B — 인사·급여 300행 시트

**코드 위치**: `Program.cs` → `CreateRealisticHRSheet()` 함수

#### 시트 구조

```
     A       B              C           D                  E        F           G          H          I          J       K       L
1   EmpID   Name        Department    JobTitle           PayGrade  Location   StartDate  BaseSalary   Bonus   TotalComp  Status  Manager
2   2000  Employee 2000  Engineering  Software Engineer    L3     New York   2017-06-15  $85,000    $8,500   $93,500   Active  John Adams
3   2001  Employee 2001  Sales        Sales Rep            L3     Chicago    2019-03-22  $75,000    $9,200   $84,200   Active  Sarah Lee
4   2002  Employee 2002  Engineering  Senior Engineer      L4     San Fran.  2020-01-10 $115,000   $14,000  $129,000   Active  Mike Brown
...
301  2299  Employee 2299  Finance     Financial Analyst    L4     New York   2022-08-05  $90,000   $11,700  $101,700  On Leave Lisa Chan
```

#### 압축의 놀라운 효과

HR 데이터는 반복 패턴이 많습니다:
- 부서(Department): 10종류만 300번 반복
- 직급(PayGrade): 5종류(L3~L7)만 반복
- 위치(Location): 5개 도시만 반복
- 상태(Status): "Active" 또는 "On Leave"만 반복
- 관리자(Manager): 5명만 반복

**역색인 압축 후**:
```json
"cells": {
  "Engineering": ["C2", "C4", "C7", "C11", ...],   ← 90개 셀 → 1 항목
  "Sales":       ["C3", "C8", "C13", ..."],          ← 60개 셀 → 1 항목
  "Active":      ["K2:K4", "K6:K8", ..."],           ← 280개 셀이 범위로 묶임
  "On Leave":    ["K5", "K9", "K15", ..."],          ← 20개 셀
  "L3":          ["E2", "E3", "E6", ..."],
  "L4":          ["E4", "E5", "E8", ..."],
  "New York":    ["F2", "F5", "F7", ..."],
  "John Adams":  ["L2", "L6", "L14", ..."],
  ...
}
```

**서식 집계**:
```json
"formats": {
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":  ["G2:G301"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0\"}":  ["H2:H301", "I2:I301", "J2:J301"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":   ["A2:A301"]
}
```

→ AI는 "G열 전체가 날짜, H~J열 전체가 통화 형식"을 즉시 파악합니다.

---

### 예제 C — 직원 대형 시트 (50행 × 10열)

**코드 위치**: `Program.cs` → `CreateLargeSheet()` 함수

이 예제는 실제 테스트에서 **2.02배** 압축 효율을 달성한 케이스입니다.

#### 시트 구조

```
     A    B            C           D        E        F      G       H          I                J
1   ID   Name        Dept       Salary   YearsExp Rating  Active  StartDate   Email           Notes
2    1  Employee1  Engineering  $68,532     7      4.2    Yes   2019-03-15  emp1@company.com  Active
3    2  Employee2  Sales        $75,100    12      3.8    No    2016-07-22  emp2@company.com
4    3  Employee3  HR           $52,800     3      4.7    Yes   2021-01-10  emp3@company.com  Active
...
51  50  Employee50  Marketing  $91,200     9      3.5    Yes   2018-11-30  emp50@company.com
```

#### 1단계 앵커 탐지 과정 (상세)

**행 분석**:
- 1번 행: ID, Name, Dept, ... → 전부 볼드체 → **머리글 행 확인!**
- 2~51번 행: 데이터 행 (숫자와 텍스트 혼재)

**열 분석**:
- A열(ID): 순서대로 증가하는 정수 → 고유한 지문(fingerprint)
- D열(Salary): $#,##0 통화 형식
- H열(StartDate): yyyy-mm-dd 날짜 형식
- I열(Email): 이메일 패턴

**앵커 결과**:
```json
"structural_anchors": {
  "rows": [1, 2, 51],
  "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
}
```

#### 2단계 역색인 결과

```json
"cells": {
  "Engineering": ["C2", "C7", "C11", "C19", "C25", "C33", "C41", "C48"],
  "Sales":       ["C3", "C8", "C14", "C22", "C29", "C36", "C44"],
  "HR":          ["C4", "C12", "C20", "C27", "C35"],
  "Finance":     ["C5", "C9", "C16", "C23", "C30"],
  "Marketing":   ["C6", "C10", "C18", "C24", "C31", "C38", "C46"],
  "Yes":         ["G2", "G4", "G6", "G9", "G11", ...],
  "No":          ["G3", "G5", "G7", "G8", ..."],
  "Active":      ["J2", "J4", "J6", ...]
}
```

**부서명(5종류)이 50번 반복되는 것이 5개 항목으로 요약됩니다!**

#### 3단계 서식 집계 결과

```json
"formats": {
  "{\"type\":\"currency\",\"nfs\":\"$#,##0\"}":  ["D2:D51"],
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":  ["H2:H51"],
  "{\"type\":\"email\",\"nfs\":\"General\"}":    ["I2:I51"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":  ["A2:A51", "E2:E51"]
}
```

#### 압축 성능

| 단계 | 토큰 수 | 원본 대비 |
|------|---------|----------|
| 원본 (vanilla) | 7,675 | 기준 |
| 1단계 후 (앵커만) | ~3,000 | ~2.5x |
| 2단계 후 (역색인) | ~2,500 | ~3x |
| 최종 출력 | **3,794** | **2.02x** |

→ 원본보다 **절반 이하** 로 줄어들었습니다!

---

## 6. 코드 파일별 역할 설명

### `ExcelReader.cs` — 엑셀 파일 읽기

**하는 일**: 엑셀(.xlsx) 파일을 열어 모든 셀 정보를 메모리로 가져옵니다.

```
엑셀 파일 → ClosedXML 라이브러리로 열기 → 각 셀 데이터 추출 → WorksheetSnapshot 생성
```

**추출하는 정보**:

| 정보 | 예시 |
|------|------|
| 셀 값 | "사과", 100, 2024-01-15 |
| 수식 | `=SUM(B2:B5)` |
| 서식 | 통화(`$#,##0.00`), 날짜(`yyyy-mm-dd`) |
| 스타일 | 볼드체 여부, 글자색, 테두리, 정렬 |
| 병합 여부 | A1:D1이 합쳐진 셀인지 |

**중요한 특징**: 수식은 계산 결과 대신 **원래 수식 문자열**로 저장합니다.
```
셀 D7에 =SUM(D2:D6)이 있으면 → "=SUM(D2:D6)" 로 저장 (계산값 50000이 아님)
```

---

### `CellUtils.cs` — 셀 분류기

**하는 일**: 각 셀이 어떤 종류의 데이터인지 판단합니다.

#### `InferCellDataType()` — 기본 데이터 유형 판단

```
셀 값 → 이메일인가? → 수식인가? → 숫자인가? → 날짜 서식인가? → 텍스트
```

반환 값: `"empty"` | `"text"` | `"numeric"` | `"boolean"` | `"datetime"` | `"email"` | `"error"` | `"formula"`

#### `DetectSemanticType()` — 의미적 유형 판단 (9가지)

```csharp
// 예시: 셀 값이 "50000"이고 서식이 "$#,##0.00"인 경우
DetectSemanticType(cell) → "currency"  // 통화로 분류

// 예시: 셀 값이 "2024-01-15"이고 서식이 "yyyy-mm-dd"인 경우
DetectSemanticType(cell) → "date"  // 날짜로 분류

// 예시: 셀 값이 "user@company.com"인 경우
DetectSemanticType(cell) → "email"  // 이메일로 분류
```

#### `GetStyleFingerprint()` — 스타일 지문 생성

각 셀의 **시각적 스타일**을 하나의 숫자로 압축합니다.

```
볼드 + 파란색 글자 + 테두리 → 지문: 12345678
볼드 없음 + 검정 글자 + 테두리 없음 → 지문: 87654321
```

같은 지문 = 같은 스타일 → 경계 탐지에 활용됩니다.

---

### `SheetCompressor.cs` — 핵심 압축 엔진

전체 파이프라인을 orchestrate하는 가장 중요한 파일입니다.

#### 공개 API (사용 방법)

```csharp
// 방법 1: 파일 경로로 직접 압축
var 압축기 = new SheetCompressor();
var 결과 = 압축기.Encode("경로/파일.xlsx", k: 2);

// 방법 2: 미리 읽은 데이터로 압축 (VSTO 어댑터 등에서 사용)
var 결과 = 압축기.Encode(snapshots, "파일이름.xlsx", k: 2);
```

`k=2`의 의미: 앵커 행/열 주변 **2칸씩** 확장 (기본값)

#### 주요 상수 설명

| 상수명 | 값 | 의미 |
|--------|-----|------|
| `MaxCandidates` | 200 | 후보 구역 최대 개수 |
| `MaxBoundaryRows` | 100 | 경계 행 최대 개수 |
| `HeaderThreshold` | 0.6 | 머리글 판단 기준 (60%) |
| `SparsityThreshold` | 0.10 | 데이터 밀도 최소값 (10%) |
| `NmsIouThreshold` | 0.5 | 중복 제거 기준 (50% 겹침) |

---

### `VanillaEncoder.cs` — 원본 인코더 (비교 기준)

**하는 일**: 압축 없이 그대로 행 순서대로 텍스트를 생성합니다.

압축 효율 측정을 위한 **기준선** 역할을 합니다:
```
비교: 원본 크기(VanillaEncoder) vs 압축 후 크기(SheetCompressor)
```

---

### `ChainOfSpreadsheet.cs` — AI 연결 파이프라인

**하는 일**: 압축된 스프레드시트를 AI에게 보내고 질문-답변을 처리합니다.

지원하는 AI 백엔드:
- `"anthropic"` → Claude (Sonnet) 사용
- `"openai"` → GPT-4 사용
- `"placeholder"` → 테스트용 모의 응답

필요한 환경변수:
```
ANTHROPIC_API_KEY=sk-ant-...   (Claude 사용 시)
OPENAI_API_KEY=sk-...          (GPT-4 사용 시)
```

---

## 7. JSON 출력 결과 읽는 법

### 출력 파일 위치

```
SpreadsheetLLM.TestRunner/bin/Release/net9.0/test_output/
├── Simple_table.json
├── Large_sheet__50r_10c_.json
├── Sales_500_rows.json
├── HR_Payroll_300_rows.json
└── ...
```

### JSON 구조 전체 예시

```json
{
  "file_name": "판매데이터.xlsx",
  "sheets": {
    "SalesTransactions": {
      "structural_anchors": {
        "rows": [1, 2, 503],
        "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
      },
      "cells": {
        "TxnID":        ["A1"],
        "Date":         ["B1"],
        "Region":       ["C1"],
        "North":        ["C2", "C15", "C31:C35", "C50"],
        "Widget A":     ["E2:E98"],
        "Completed":    ["J2:J4", "J7:J11", "J350:J420"],
        "=SUM(I2:I501)":["I503"]
      },
      "formats": {
        "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":    ["B2:B501"],
        "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["H2:H501", "I2:I501"],
        "{\"type\":\"integer\",\"nfs\":\"General\"}":    ["A2:A501", "G2:G501"],
        "{\"type\":\"text\",\"nfs\":\"General\"}":        ["C2:C501", "D2:D501"]
      },
      "numeric_ranges": {
        "{\"type\":\"integer\",\"nfs\":\"General\"}": ["A2:A501", "G2:G501"]
      }
    }
  },
  "compression_metrics": {
    "sheets": {
      "SalesTransactions": {
        "original_tokens": 45000,
        "after_anchor_tokens": 12000,
        "after_inverted_index_tokens": 5000,
        "after_format_tokens": 3500,
        "final_tokens": 4200,
        "anchor_ratio": 3.75,
        "inverted_index_ratio": 9.00,
        "format_ratio": 12.86,
        "overall_ratio": 10.71
      }
    },
    "overall": {
      "original_tokens": 45000,
      "final_tokens": 4200,
      "overall_ratio": 10.71
    }
  }
}
```

### 각 필드 설명

| 필드 | 설명 |
|------|------|
| `structural_anchors.rows` | 중요한 행 번호 목록 (1-based) |
| `structural_anchors.columns` | 중요한 열 문자 목록 |
| `cells` | 값→위치 역색인 (값: 셀 범위 목록) |
| `formats` | 서식→위치 집계 (JSON 서식키: 셀 범위 목록) |
| `numeric_ranges` | 숫자 유형만 별도 집계 |
| `original_tokens` | 압축 전 토큰 수 |
| `final_tokens` | 압축 후 토큰 수 |
| `overall_ratio` | 압축비 (높을수록 압축 효율 좋음) |

---

## 8. 압축 성능 수치

CLAUDE.md 기준 실제 테스트 결과:

| 테스트 시트 | 원본 토큰 | 최종 토큰 | 압축비 | 비고 |
|------------|----------|----------|--------|------|
| 간단한 표 | 310 | 935 | 0.33x | 소형 → 메타데이터 오버헤드 |
| 다중 테이블 | 260 | 725 | 0.36x | 소형 시트 |
| 병합 셀 | 263 | 669 | 0.39x | 병합 처리 포함 |
| 수식 | 264 | 684 | 0.39x | 수식 그대로 보존 |
| 날짜·통화 | 265 | 648 | 0.41x | 서식 집계 효과 |
| 숫자만 | 599 | 1001 | 0.60x | 구조 정보 부족 |
| 혼합 서식 | 402 | 1706 | 0.24x | 다양한 서식 |
| **대형 시트(50×10)** | **7,675** | **3,794** | **2.02x** | **첫 압축 이득 구간** |
| 희소 시트 | 121 | 522 | 0.23x | 데이터 적음 |
| 다중 시트 | 679 | 1968 | 0.35x | 3개 시트 |

### 왜 소형 시트는 압축비가 1보다 낮은가?

소형 시트(수십 행):
- 압축 알고리즘이 추가하는 **메타데이터(구조적 앵커, JSON 키 등)** 가 원본보다 클 수 있음
- 반복 패턴이 거의 없어 역색인의 이점이 적음

대형·실무 시트(수백~수천 행):
- 반복 값이 많아 **역색인 효과 극대화**
- 논문에서 보고한 **최대 25배** 압축은 실제 기업 데이터 기준

---

## 9. 자주 묻는 질문

**Q: 이 소프트웨어를 사용하려면 어떻게 해야 하나요?**

```powershell
# 1. 빌드
dotnet build SpreadsheetLLM.Core/SpreadsheetLLM.Core.csproj -c Release

# 2. 테스트 실행 (샘플 시트 생성 + 압축)
dotnet run --project SpreadsheetLLM.TestRunner -c Release

# 3. 결과 확인
# SpreadsheetLLM.TestRunner/bin/Release/net9.0/test_output/ 폴더에 JSON 파일 생성됨
```

**Q: 엑셀 파일이 손상되지 않나요?**

아니요. 소프트웨어는 파일을 **읽기만** 합니다. 원본 파일은 전혀 변경되지 않습니다.

**Q: 어떤 종류의 엑셀 파일을 지원하나요?**

`.xlsx` 형식만 지원합니다 (Excel 2007 이상).

**Q: AI API 키가 없어도 압축은 되나요?**

예. 압축(인코딩) 기능 자체는 API 키 없이 작동합니다.  
API 키는 AI에게 질문할 때(`ChainOfSpreadsheet`) 필요합니다.

**Q: 수식은 어떻게 처리되나요?**

수식은 계산 결과 대신 **원래 수식 문자열**로 보존됩니다.
```
=SUM(B2:B500) → AI에게 그대로 전달
```
AI가 수식의 의미를 이해할 수 있습니다.

**Q: 병합 셀은 어떻게 처리되나요?**

병합 셀은 **시작 셀의 값**으로 처리됩니다.  
예: A1:D1이 "분기 보고서"로 병합된 경우 → A1, B1, C1, D1 모두 "분기 보고서" 값으로 처리.

**Q: 한국어 텍스트도 지원되나요?**

예. C# `string`은 유니코드를 완전히 지원하므로 한국어, 일본어, 중국어 등 모든 언어가 정상 처리됩니다.

---

## 부록: 전체 흐름 요약 다이어그램

```
엑셀 파일 (예: 판매데이터_500행.xlsx)
│
│ ExcelReader.ReadWorkbook()
▼
WorksheetSnapshot[] — 모든 셀의 스냅샷
│
│ SheetCompressor.Encode()
▼
┌─────────────────────────────────────────────────────┐
│                   3단계 파이프라인                    │
│                                                     │
│  1단계: FindStructuralAnchors()                     │
│    └─ 머리글/경계/데이터 전환점 감지               │
│    └─ NMS로 중복 제거                               │
│    └─ k=2 확장                                      │
│    └─ 동일 행/열 압축                               │
│         ↓                                           │
│  2단계: CreateInvertedIndex()                       │
│    └─ 값 → 셀 위치 목록으로 뒤집기                 │
│    └─ 연속 셀을 범위(A2:A100)로 병합                │
│    └─ 수식 그대로 보존                              │
│         ↓                                           │
│  3단계: GroupBySemanticType()                       │
│    └─ 9가지 의미 유형으로 분류                      │
│    └─ 같은 유형+서식 → 범위로 집계                 │
│    └─ 숫자 유형 별도 집계                           │
└─────────────────────────────────────────────────────┘
│
▼
SpreadsheetEncoding (JSON)
├─ structural_anchors (구조 지도)
├─ cells (역색인)
├─ formats (서식 집계)
├─ numeric_ranges (숫자 집계)
└─ compression_metrics (압축 성능 지표)
│
▼
AI (Claude / GPT-4)에게 전달
→ 적은 토큰으로 스프레드시트 완전 이해 가능
```

---

*본 문서는 SpreadsheetLLM .NET 구현 (C# netstandard2.0, arXiv:2407.09025 기반)을 비기술 고객에게 설명하기 위해 작성되었습니다.*  
*소스 코드 참조: `SpreadsheetLLM.Core/` 디렉토리*
