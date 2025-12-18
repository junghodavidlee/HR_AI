# 지원자 이력서 자동화 시스템
# Applicant Resume Automation System

이 시스템은 이력서에서 추출된 JSON 데이터를 검증하고 Excel 스프레드시트에 자동으로 추가합니다.

## 필요한 패키지 설치

```bash
pip install openpyxl --break-system-packages
```

## 파일 구조

```
.
├── applicant_excel_writer.py  # Excel 파일 생성 및 데이터 추가
├── validator.py                # 데이터 검증 및 정제
├── main.py                     # 실행 스크립트
└── README.md                   # 이 파일
```

## 사용 방법

### 1. 단일 이력서 처리

```python
from applicant_excel_writer import ApplicantExcelWriter
from validator import process_applicant_resume

# Excel writer 초기화
writer = ApplicantExcelWriter("applicants.xlsx")

# JSON 데이터
resume_data = {
    "applicant_name": "홍길동",
    "application_date": "2024-12-19",
    "affiliation": "서울대학교",
    "application_field": "소프트웨어 개발",
    "basic_info": {
        "birth_year": "1990",
        "gender": "남",
        "final_education_school": "고려대학교",
        "final_education_degree": "석사"
    },
    "work_experience": [
        {
            "start_date": "2020-03",
            "end_date": "재직중",
            "company_name": "네이버",
            "final_department": "AI Lab",
            "final_position": "선임연구원",
            "salary": 85000
        }
    ]
}

# 처리 (검증 + Excel 추가)
process_applicant_resume(resume_data, writer, strict_mode=False)
```

### 2. 여러 이력서 일괄 처리

```python
from main import batch_process_resumes
import os

# JSON 파일들이 있는 디렉토리
json_directory = "resume_json_outputs"
json_files = [
    os.path.join(json_directory, f) 
    for f in os.listdir(json_directory) 
    if f.endswith('.json')
]

# 일괄 처리
batch_process_resumes(
    json_files=json_files,
    excel_path="applicants.xlsx",
    strict_mode=False
)
```

### 3. main.py 직접 실행

```bash
python main.py
```

## JSON 데이터 스키마

```json
{
  "applicant_name": "김정호",
  "application_date": "2024-12-19",
  "affiliation": "한국증권",
  "application_field": "AI 솔루션 팀",
  "basic_info": {
    "birth_year": "1995",
    "gender": "남",
    "final_education_school": "서울대학교",
    "final_education_degree": "학사"
  },
  "work_experience": [
    {
      "start_date": "2022-01",
      "end_date": "재직중",
      "company_name": "한국증권",
      "final_department": "AI솔루션팀",
      "final_position": "매니저",
      "salary": 80000
    }
  ]
}
```

### 필수 필드 (Mandatory Fields)
- `applicant_name` (지원자명)
- `application_date` (지원일)
- `affiliation` (소속)
- `application_field` (지원분야/공고)

**참고**: `applicant_number` (지원자 번호)는 자동으로 순차적으로 생성됩니다 (1, 2, 3, ...).

### 선택 필드 (Optional Fields)
- `basic_info` (기본정보)
- `work_experience` (경력사항)

## 주요 기능

### 1. 지원자 번호 자동 생성
- `applicant_number`는 JSON에 포함하지 않아도 됩니다
- Excel의 행 번호에 따라 자동으로 순차 생성 (1, 2, 3, ...)
- 첫 번째 지원자는 1번, 두 번째는 2번...

### 2. 자동 검증
- 필수 필드 확인 (지원자명, 지원일, 소속, 지원분야/공고)
- 날짜 형식 검증 (YYYY-MM-DD, YYYY-MM)
- 데이터 타입 확인
- 논리적 오류 감지 (예: 입사일 > 퇴사일)

### 3. 데이터 정제
- 공백 제거
- 날짜 형식 정규화
- 성별 표준화 (남/여/기타)
- 경력 최신순 정렬
- 최대 5개 경력만 유지

### 4. Excel 자동 추가
- 기존 파일에 새 행 추가
- 자동 열 매핑
- 템플릿이 없으면 자동 생성

## 검증 모드

### strict_mode=False (권장)
- 경고가 있어도 데이터 추가
- 오류만 거부

### strict_mode=True
- 경고나 오류가 있으면 데이터 추가 안함
- 더 엄격한 검증

## Excel 컬럼 구조

| 컬럼 | 항목 | 비고 |
|------|------|------|
| A | 지원자 번호 | |
| B | 지원자명 | |
| C | 지원일 | YYYY-MM-DD |
| D | 소속 | |
| E | 지원분야/공고 | |
| F | 출생년도 | YYYY |
| G | 성별 | 남/여/기타 |
| H | 최종학력(학교) | |
| I | 최종학력(학사 등) | |
| J-O | 경력 1 | 입사년월, 퇴사년월, 회사명, 최종부서명, 최종직위, 연봉 |
| P-U | 경력 2 | 동일 |
| V-AA | 경력 3 | 동일 |
| AB-AG | 경력 4 | 동일 |
| AH-AM | 경력 5 | 동일 |

## 문제 해결

### 패키지 설치 오류
```bash
pip install openpyxl --break-system-packages
```

### Excel 파일이 열려있을 때 오류
- Excel 파일을 닫고 다시 실행

### JSON 파싱 오류
- JSON 형식이 올바른지 확인
- UTF-8 인코딩 확인

## 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.
