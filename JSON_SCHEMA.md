# JSON Schema for Resume Data Extraction

## Complete JSON Schema

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "ApplicantData",
  "type": "object",
  "required": ["applicant_name", "application_date", "affiliation", "application_field"],
  "properties": {
    "applicant_name": {
      "type": "string",
      "description": "지원자명 - Full name of applicant (REQUIRED)"
    },
    "application_date": {
      "type": "string",
      "pattern": "^\\d{4}-\\d{2}-\\d{2}$",
      "description": "지원일 - Application date in YYYY-MM-DD format (REQUIRED)"
    },
    "affiliation": {
      "type": "string",
      "description": "소속 - Current affiliation/organization (REQUIRED)"
    },
    "application_field": {
      "type": "string",
      "description": "지원분야/공고 - Position/job posting applied for (REQUIRED)"
    },
    
    "basic_info": {
      "type": "object",
      "description": "기본정보 - Basic personal information (OPTIONAL)",
      "properties": {
        "birth_year": {
          "type": "string",
          "pattern": "^\\d{4}$",
          "description": "출생년도 - Birth year (YYYY)"
        },
        "gender": {
          "type": "string",
          "enum": ["남", "여", "기타"],
          "description": "성별 - Gender"
        },
        "final_education_school": {
          "type": "string",
          "description": "최종학력(학교) - Name of final educational institution"
        },
        "final_education_degree": {
          "type": "string",
          "description": "최종학력(학사 등) - Degree level (e.g., 고졸, 학사, 석사, 박사)"
        }
      }
    },
    
    "work_experience": {
      "type": "array",
      "maxItems": 5,
      "description": "경력 - Work experience, max 5 most recent (OPTIONAL)",
      "items": {
        "type": "object",
        "properties": {
          "start_date": {
            "type": "string",
            "pattern": "^\\d{4}-\\d{2}$",
            "description": "입사년월 - Start date (YYYY-MM)"
          },
          "end_date": {
            "type": ["string", "null"],
            "pattern": "^(\\d{4}-\\d{2}|재직중)$",
            "description": "퇴사년월 - End date (YYYY-MM) or '재직중'"
          },
          "company_name": {
            "type": "string",
            "description": "회사명 - Company name"
          },
          "final_department": {
            "type": ["string", "null"],
            "description": "최종부서명 - Final department name"
          },
          "final_position": {
            "type": ["string", "null"],
            "description": "최종직위 - Final position/title"
          },
          "salary": {
            "type": ["integer", "null"],
            "description": "연봉(천원 단위) - Annual salary in thousands of KRW"
          }
        }
      }
    }
  }
}
```

## Important Notes

### Applicant Number
- **DO NOT include `applicant_number` in the JSON output**
- This field is automatically generated sequentially (1, 2, 3, ...) by the Excel writer
- The number is based on the row position in the Excel file

### Required Fields (필수 필드)
Only these 4 fields are mandatory:
1. `applicant_name` (지원자명)
2. `application_date` (지원일) - Format: YYYY-MM-DD
3. `affiliation` (소속)
4. `application_field` (지원분야/공고)

### Optional Fields (선택 필드)
All other fields are optional:
- `basic_info` object and all its sub-fields
- `work_experience` array and all experience entries

If optional fields are not available in the resume, simply omit them or set to `null`.

## Example Minimal JSON (Only Required Fields)

```json
{
  "applicant_name": "홍길동",
  "application_date": "2024-12-19",
  "affiliation": "서울대학교",
  "application_field": "소프트웨어 엔지니어"
}
```

## Example Complete JSON (With Optional Fields)

```json
{
  "applicant_name": "김정호",
  "application_date": "2024-12-19",
  "affiliation": "한국증권",
  "application_field": "AI 솔루션 팀 매니저",
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
    },
    {
      "start_date": "2020-03",
      "end_date": "2021-12",
      "company_name": "삼성전자",
      "final_department": "소프트웨어개발팀",
      "final_position": "주임",
      "salary": 60000
    }
  ]
}
```

## LLM Extraction Prompt Template

```
당신은 한국 기업의 지원자 이력서를 분석하는 전문가입니다. 
다음 이력서 텍스트에서 정보를 추출하여 JSON 형식으로만 출력하세요.

이력서 텍스트:
{ocr_text}

필수 추출 규칙:
1. 필수 필드 (반드시 추출):
   - applicant_name: 지원자 이름
   - application_date: 지원일 (YYYY-MM-DD 형식)
   - affiliation: 현재 소속 또는 최종 학력
   - application_field: 지원 직무/분야

2. 선택 필드 (가능한 경우만 추출):
   - basic_info: 출생년도, 성별, 최종학력
   - work_experience: 경력사항 (최신순으로 최대 5개)

3. 날짜 형식:
   - application_date: YYYY-MM-DD (예: 2024-12-19)
   - start_date/end_date: YYYY-MM (예: 2022-01)
   - 현재 재직중: "재직중"

4. 연봉: 천원 단위로 숫자만 (예: 60,000,000원 → 60000)

5. 성별: "남", "여", "기타" 중 하나

6. 학력: "고졸", "전문학사", "학사", "석사", "박사" 등

7. 경력 정렬: 최신 경력이 첫 번째

8. applicant_number는 절대 포함하지 마세요 (자동 생성됨)

중요: 
- 정보가 없는 선택 필드는 생략하거나 null로 설정
- 마크다운 코드 블록이나 설명 없이 JSON 객체만 반환
- 필수 필드가 누락되면 안 됨

JSON 출력:
```

## Validation Rules

The validator will check:

1. **Required Fields**: Must have all 4 mandatory fields
2. **Date Formats**: 
   - `application_date`: YYYY-MM-DD
   - `start_date`/`end_date`: YYYY-MM or "재직중"
3. **Gender Values**: Must be "남", "여", or "기타"
4. **Birth Year**: Must be 4-digit year (1940-current year)
5. **Work Experience**: 
   - Maximum 5 entries
   - start_date must be before end_date
   - salary must be non-negative integer
6. **Date Logic**: Application date cannot be in the future

Warnings (non-blocking):
- Unusual birth years
- Non-standard education degrees
- Very high salaries (>1,000,000천원)
- More than 5 work experiences (only first 5 will be used)
