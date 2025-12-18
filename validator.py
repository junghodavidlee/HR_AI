import json
from typing import Dict, Any, List, Tuple
from datetime import datetime
import re


class ApplicantDataValidator:
    """Validates JSON data before writing to Excel"""
    
    def __init__(self, strict_mode: bool = False):
        """
        Args:
            strict_mode: If True, reject data with any validation errors.
                        If False, allow data through with warnings.
        """
        self.strict_mode = strict_mode
        self.errors = []
        self.warnings = []
    
    def validate(self, data: Dict[str, Any]) -> Tuple[bool, List[str], List[str]]:
        """
        Validate applicant data
        
        Returns:
            Tuple of (is_valid, errors, warnings)
        """
        self.errors = []
        self.warnings = []
        
        # Check required fields
        self._validate_required_fields(data)
        
        # Validate basic info
        self._validate_basic_info(data)
        
        # Validate work experience
        self._validate_work_experience(data)
        
        # Validate dates
        self._validate_dates(data)
        
        # Check data types
        self._validate_data_types(data)
        
        is_valid = len(self.errors) == 0 if self.strict_mode else True
        
        return is_valid, self.errors, self.warnings
    
    def _validate_required_fields(self, data: Dict[str, Any]):
        """Check that all required fields are present"""
        required_fields = [
            'applicant_number',
            'applicant_name',
            'application_date',
            'basic_info.birth_year',
            'basic_info.gender',
            'basic_info.final_education_school',
            'basic_info.final_education_degree'
        ]
        
        for field_path in required_fields:
            if not self._get_nested_value(data, field_path):
                self.errors.append(f"필수 필드 누락: {field_path}")
    
    def _validate_basic_info(self, data: Dict[str, Any]):
        """Validate basic information fields"""
        basic_info = data.get('basic_info', {})
        
        # Validate birth year
        birth_year = basic_info.get('birth_year', '')
        if birth_year:
            if not re.match(r'^\d{4}$', str(birth_year)):
                self.errors.append(f"출생년도 형식 오류: '{birth_year}' (YYYY 형식이어야 함)")
            else:
                year = int(birth_year)
                current_year = datetime.now().year
                if year < 1940 or year > current_year:
                    self.warnings.append(f"출생년도가 비정상적임: {birth_year}")
        
        # Validate gender
        gender = basic_info.get('gender', '')
        valid_genders = ['남', '여', '기타']
        if gender and gender not in valid_genders:
            self.errors.append(f"성별 값 오류: '{gender}' (허용: {valid_genders})")
        
        # Validate education degree
        degree = basic_info.get('final_education_degree', '')
        valid_degrees = ['고졸', '전문학사', '학사', '석사', '박사', '기타']
        if degree and degree not in valid_degrees:
            self.warnings.append(f"학력 값이 표준과 다름: '{degree}' (권장: {valid_degrees})")
    
    def _validate_work_experience(self, data: Dict[str, Any]):
        """Validate work experience entries"""
        experiences = data.get('work_experience', [])
        
        if not isinstance(experiences, list):
            self.errors.append("work_experience는 배열이어야 합니다")
            return
        
        if len(experiences) > 5:
            self.warnings.append(f"경력이 5개를 초과함 ({len(experiences)}개). 최신 5개만 사용됩니다.")
        
        for idx, exp in enumerate(experiences[:5]):
            self._validate_single_experience(exp, idx + 1)
    
    def _validate_single_experience(self, exp: Dict[str, Any], exp_num: int):
        """Validate a single work experience entry"""
        # Check required fields
        if not exp.get('start_date'):
            self.errors.append(f"경력 {exp_num}: 입사년월 누락")
        if not exp.get('company_name'):
            self.errors.append(f"경력 {exp_num}: 회사명 누락")
        
        # Validate start_date format
        start_date = exp.get('start_date', '')
        if start_date and not re.match(r'^\d{4}-\d{2}$', str(start_date)):
            self.errors.append(f"경력 {exp_num}: 입사년월 형식 오류 '{start_date}' (YYYY-MM 형식이어야 함)")
        
        # Validate end_date format
        end_date = exp.get('end_date')
        if end_date and end_date != '재직중':
            if not re.match(r'^\d{4}-\d{2}$', str(end_date)):
                self.errors.append(f"경력 {exp_num}: 퇴사년월 형식 오류 '{end_date}' (YYYY-MM 또는 '재직중'이어야 함)")
        
        # Validate date logic (start before end)
        if start_date and end_date and end_date != '재직중':
            try:
                start = datetime.strptime(start_date, '%Y-%m')
                end = datetime.strptime(end_date, '%Y-%m')
                if start > end:
                    self.errors.append(f"경력 {exp_num}: 입사일이 퇴사일보다 늦음")
            except ValueError:
                pass  # Format errors already caught above
        
        # Validate salary
        salary = exp.get('salary')
        if salary is not None:
            if not isinstance(salary, (int, float)):
                self.errors.append(f"경력 {exp_num}: 연봉은 숫자여야 함 (천원 단위)")
            elif salary < 0:
                self.errors.append(f"경력 {exp_num}: 연봉은 0 이상이어야 함")
            elif salary > 1000000:  # 10억 이상
                self.warnings.append(f"경력 {exp_num}: 연봉이 매우 높음 ({salary:,}천원). 확인 필요.")
    
    def _validate_dates(self, data: Dict[str, Any]):
        """Validate all date fields"""
        # Validate application_date
        app_date = data.get('application_date', '')
        if app_date:
            if not re.match(r'^\d{4}-\d{2}-\d{2}$', str(app_date)):
                self.errors.append(f"지원일 형식 오류: '{app_date}' (YYYY-MM-DD 형식이어야 함)")
            else:
                try:
                    date_obj = datetime.strptime(app_date, '%Y-%m-%d')
                    if date_obj > datetime.now():
                        self.warnings.append(f"지원일이 미래 날짜임: {app_date}")
                except ValueError:
                    self.errors.append(f"지원일 날짜 값 오류: '{app_date}'")
    
    def _validate_data_types(self, data: Dict[str, Any]):
        """Validate data types of all fields"""
        # Check string fields
        string_fields = [
            'applicant_number',
            'applicant_name',
            'application_date',
            'affiliation',
            'application_field'
        ]
        
        for field in string_fields:
            value = data.get(field)
            if value is not None and not isinstance(value, str):
                self.errors.append(f"'{field}' 필드는 문자열이어야 함")
        
        # Check basic_info fields
        basic_info = data.get('basic_info', {})
        if not isinstance(basic_info, dict):
            self.errors.append("basic_info는 객체여야 함")
    
    def _get_nested_value(self, data: Dict, path: str) -> Any:
        """Extract value from nested dictionary using dot notation"""
        keys = path.split('.')
        value = data
        
        for key in keys:
            if isinstance(value, dict):
                value = value.get(key)
            else:
                return None
                
            if value is None:
                return None
                
        return value
    
    def print_validation_report(self, data: Dict[str, Any]):
        """Print a detailed validation report"""
        is_valid, errors, warnings = self.validate(data)
        
        applicant_name = data.get('applicant_name', 'Unknown')
        applicant_number = data.get('applicant_number', 'Unknown')
        
        print(f"\n{'='*70}")
        print(f"검증 보고서: {applicant_name} ({applicant_number})")
        print(f"{'='*70}")
        
        if is_valid and not warnings:
            print("✓ 모든 검증 통과")
        else:
            if errors:
                print(f"\n오류 ({len(errors)}개):")
                for error in errors:
                    print(f"  ✗ {error}")
            
            if warnings:
                print(f"\n경고 ({len(warnings)}개):")
                for warning in warnings:
                    print(f"  ⚠ {warning}")
        
        print(f"\n상태: {'통과' if is_valid else '실패'}")
        print(f"{'='*70}\n")
        
        return is_valid


class DataCleaner:
    """Clean and normalize data before validation"""
    
    @staticmethod
    def clean(data: Dict[str, Any]) -> Dict[str, Any]:
        """Clean and normalize applicant data"""
        cleaned = data.copy()
        
        # Trim all string fields
        DataCleaner._trim_strings(cleaned)
        
        # Normalize dates
        DataCleaner._normalize_dates(cleaned)
        
        # Normalize gender
        if 'basic_info' in cleaned and 'gender' in cleaned['basic_info']:
            gender = cleaned['basic_info']['gender']
            gender_map = {
                '남성': '남', 'male': '남', 'M': '남',
                '여성': '여', 'female': '여', 'F': '여'
            }
            cleaned['basic_info']['gender'] = gender_map.get(gender, gender)
        
        # Sort work experience by start_date (most recent first)
        if 'work_experience' in cleaned:
            experiences = cleaned['work_experience']
            if isinstance(experiences, list):
                # Sort by start_date descending
                try:
                    experiences.sort(
                        key=lambda x: x.get('start_date', '0000-00'),
                        reverse=True
                    )
                    # Keep only top 5
                    cleaned['work_experience'] = experiences[:5]
                except Exception:
                    pass  # Keep original order if sorting fails
        
        # Convert salary to integer if it's a string
        if 'work_experience' in cleaned:
            for exp in cleaned['work_experience']:
                if 'salary' in exp and isinstance(exp['salary'], str):
                    try:
                        exp['salary'] = int(exp['salary'].replace(',', ''))
                    except ValueError:
                        exp['salary'] = None
        
        return cleaned
    
    @staticmethod
    def _trim_strings(obj: Any):
        """Recursively trim all strings in the object"""
        if isinstance(obj, dict):
            for key, value in obj.items():
                if isinstance(value, str):
                    obj[key] = value.strip()
                else:
                    DataCleaner._trim_strings(value)
        elif isinstance(obj, list):
            for item in obj:
                DataCleaner._trim_strings(item)
    
    @staticmethod
    def _normalize_dates(data: Dict[str, Any]):
        """Normalize date formats"""
        # Normalize application_date
        if 'application_date' in data:
            date_str = data['application_date']
            # Handle various formats: 2024.12.15, 2024/12/15, 20241215
            date_str = date_str.replace('.', '-').replace('/', '-')
            if len(date_str) == 8 and date_str.isdigit():
                date_str = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            data['application_date'] = date_str
        
        # Normalize work experience dates
        if 'work_experience' in data:
            for exp in data['work_experience']:
                for date_field in ['start_date', 'end_date']:
                    if date_field in exp and exp[date_field]:
                        date_str = str(exp[date_field])
                        if date_str not in ['재직중', 'Present', '현재']:
                            date_str = date_str.replace('.', '-').replace('/', '-')
                            if len(date_str) == 6 and date_str.isdigit():
                                date_str = f"{date_str[:4]}-{date_str[4:]}"
                            exp[date_field] = date_str
                        else:
                            exp[date_field] = '재직중'


def process_applicant_resume(json_data: Dict[str, Any], 
                             excel_writer,
                             strict_mode: bool = False) -> bool:
    """
    Complete pipeline: Clean → Validate → Write to Excel
    
    Args:
        json_data: Raw JSON data from LLM
        excel_writer: Instance of ApplicantExcelWriter
        strict_mode: If True, reject data with any errors
    
    Returns:
        True if data was successfully written to Excel
    """
    # Step 1: Clean the data
    print("1. 데이터 정제 중...")
    cleaner = DataCleaner()
    cleaned_data = cleaner.clean(json_data)
    
    # Step 2: Validate the data
    print("2. 데이터 검증 중...")
    validator = ApplicantDataValidator(strict_mode=strict_mode)
    is_valid = validator.print_validation_report(cleaned_data)
    
    # Step 3: Write to Excel (if valid or if not in strict mode)
    if is_valid or not strict_mode:
        print("3. Excel에 데이터 추가 중...")
        try:
            row_num = excel_writer.append_applicant(cleaned_data)
            print(f"✓ 성공: {cleaned_data['applicant_name']} (행 {row_num})")
            return True
        except Exception as e:
            print(f"✗ Excel 작성 실패: {e}")
            return False
    else:
        print("✗ 검증 실패로 인해 Excel에 추가되지 않음")
        return False
