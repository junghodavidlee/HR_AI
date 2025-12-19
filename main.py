import json
import os
import sys
from pathlib import Path
from applicant_excel_writer import ApplicantExcelWriter
from validator import process_applicant_resume, ApplicantDataValidator, DataCleaner


def process_single_json_dict(json_data: dict, excel_path: str = "applicants.xlsx", strict_mode: bool = False):
    """
    Process a single applicant from a Python dictionary
    
    Args:
        json_data: Dictionary containing applicant data
        excel_path: Path to Excel file
        strict_mode: If True, reject data with validation errors
        
    Returns:
        True if successful
    """
    writer = ApplicantExcelWriter(excel_path)
    
    # Create template if needed
    if not os.path.exists(excel_path):
        print(f"Excel 파일을 새로 생성합니다: {excel_path}")
        writer.create_template()
    
    return process_applicant_resume(json_data, writer, strict_mode)


def process_single_json_file(json_file_path: str, excel_path: str = "applicants.xlsx", strict_mode: bool = False):
    """
    Process a single applicant from a JSON file
    
    Args:
        json_file_path: Path to JSON file
        excel_path: Path to Excel file
        strict_mode: If True, reject data with validation errors
        
    Returns:
        True if successful
    """
    print(f"\n{'='*70}")
    print(f"처리 중: {json_file_path}")
    print(f"{'='*70}")
    
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        # Check if data is a list (array of applicants)
        if isinstance(json_data, list):
            print(f"⚠ JSON 파일에 {len(json_data)}개의 지원자가 배열로 있습니다.")
            print(f"첫 번째 지원자만 처리합니다. 모든 지원자를 처리하려면 batch_process를 사용하세요.")
            if len(json_data) > 0:
                json_data = json_data[0]
            else:
                print("✗ 빈 배열입니다.")
                return False
        
        # Check if data is a dict
        if not isinstance(json_data, dict):
            print(f"✗ JSON 데이터가 올바른 형식이 아닙니다 (타입: {type(json_data).__name__})")
            print(f"딕셔너리 형식이어야 합니다: {{'applicant_name': '...', ...}}")
            return False
        
        return process_single_json_dict(json_data, excel_path, strict_mode)
        
    except FileNotFoundError:
        print(f"✗ 파일을 찾을 수 없음: {json_file_path}")
        return False
    except json.JSONDecodeError as e:
        print(f"✗ JSON 파싱 오류: {json_file_path}")
        print(f"  상세: {e}")
        return False
    except Exception as e:
        print(f"✗ 예상치 못한 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_json_string(json_string: str, excel_path: str = "applicants.xlsx", strict_mode: bool = False):
    """
    Process a single applicant from a JSON string
    
    Args:
        json_string: JSON string containing applicant data
        excel_path: Path to Excel file
        strict_mode: If True, reject data with validation errors
        
    Returns:
        True if successful
    """
    try:
        json_data = json.loads(json_string)
        return process_single_json_dict(json_data, excel_path, strict_mode)
    except json.JSONDecodeError as e:
        print(f"✗ JSON 파싱 오류: {e}")
        return False


def batch_process_json_files(json_files: list, excel_path: str = "applicants.xlsx", strict_mode: bool = False):
    """
    Process multiple resume JSON files and append to Excel
    
    Args:
        json_files: List of paths to JSON files
        excel_path: Path to Excel file (will be created if doesn't exist)
        strict_mode: If True, reject any data with validation errors
        
    Returns:
        Dictionary with results summary
    """
    # Initialize writer
    writer = ApplicantExcelWriter(excel_path)
    
    # Create template if needed
    if not os.path.exists(excel_path):
        print(f"Excel 파일이 없습니다. 새로 생성합니다: {excel_path}")
        writer.create_template()
    else:
        print(f"기존 Excel 파일에 추가합니다: {excel_path}")
    
    # Process each resume
    results = {
        'success': [],
        'failed': [],
        'warnings': []
    }
    
    for json_file in json_files:
        print(f"\n{'='*70}")
        print(f"처리 중: {json_file}")
        print(f"{'='*70}")
        
        try:
            # Load JSON
            with open(json_file, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # Process (clean, validate, write)
            if process_applicant_resume(json_data, writer, strict_mode):
                results['success'].append(json_file)
                
                # Check for warnings
                validator = ApplicantDataValidator()
                is_valid, errors, warnings = validator.validate(
                    DataCleaner.clean(json_data)
                )
                if warnings:
                    results['warnings'].append((json_file, warnings))
            else:
                results['failed'].append(json_file)
                
        except FileNotFoundError:
            print(f"✗ 파일을 찾을 수 없음: {json_file}")
            results['failed'].append(json_file)
        except json.JSONDecodeError as e:
            print(f"✗ JSON 파싱 오류: {json_file}")
            print(f"  상세: {e}")
            results['failed'].append(json_file)
        except Exception as e:
            print(f"✗ 예상치 못한 오류: {json_file}")
            print(f"  상세: {e}")
            results['failed'].append(json_file)
    
    # Print summary
    print(f"\n{'='*70}")
    print("처리 요약")
    print(f"{'='*70}")
    print(f"✓ 성공: {len(results['success'])}개")
    print(f"✗ 실패: {len(results['failed'])}개")
    print(f"⚠ 경고 있음: {len(results['warnings'])}개")
    
    if results['failed']:
        print(f"\n실패한 파일:")
        for file in results['failed']:
            print(f"  - {file}")
    
    if results['warnings']:
        print(f"\n경고가 있는 파일:")
        for file, warnings in results['warnings']:
            print(f"  - {file}: {len(warnings)}개 경고")
    
    print(f"\nExcel 파일 위치: {os.path.abspath(excel_path)}")
    print(f"{'='*70}\n")
    
    return results


def batch_process_from_directory(directory: str, excel_path: str = "applicants.xlsx", strict_mode: bool = False):
    """
    Process all JSON files in a directory
    
    Args:
        directory: Directory containing JSON files
        excel_path: Path to Excel file
        strict_mode: If True, reject data with validation errors
        
    Returns:
        Dictionary with results summary
    """
    if not os.path.exists(directory):
        print(f"✗ 디렉토리를 찾을 수 없음: {directory}")
        return None
    
    json_files = [
        os.path.join(directory, f)
        for f in os.listdir(directory)
        if f.endswith('.json')
    ]
    
    if not json_files:
        print(f"✗ {directory}에 JSON 파일이 없습니다")
        return None
    
    print(f"📁 {len(json_files)}개의 JSON 파일을 발견했습니다")
    return batch_process_json_files(json_files, excel_path, strict_mode)


# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

def main_cli():
    """Command line interface for processing resumes"""
    import argparse
    
    parser = argparse.ArgumentParser(description='지원자 이력서 데이터를 Excel로 변환')
    
    parser.add_argument(
        'input',
        help='JSON 파일 경로, JSON 파일이 있는 디렉토리, 또는 JSON 문자열'
    )
    parser.add_argument(
        '-o', '--output',
        default='applicants.xlsx',
        help='출력 Excel 파일 경로 (기본값: applicants.xlsx)'
    )
    parser.add_argument(
        '-s', '--strict',
        action='store_true',
        help='엄격 모드 (경고가 있으면 데이터 추가 안함)'
    )
    parser.add_argument(
        '-d', '--directory',
        action='store_true',
        help='입력을 디렉토리로 처리 (모든 JSON 파일 처리)'
    )
    
    args = parser.parse_args()
    
    if args.directory:
        # Process directory
        batch_process_from_directory(args.input, args.output, args.strict)
    elif os.path.isfile(args.input):
        # Process single file
        process_single_json_file(args.input, args.output, args.strict)
    elif os.path.isdir(args.input):
        # Auto-detect directory
        batch_process_from_directory(args.input, args.output, args.strict)
    else:
        # Try to parse as JSON string
        try:
            process_json_string(args.input, args.output, args.strict)
        except:
            print(f"✗ 입력을 인식할 수 없습니다: {args.input}")
            print("파일 경로, 디렉토리 경로, 또는 JSON 문자열을 입력하세요")
            sys.exit(1)


# ============================================================================
# USAGE EXAMPLES
# ============================================================================

if __name__ == "__main__":
    # Check if running from command line with arguments
    if len(sys.argv) > 1:
        main_cli()
    else:
        # Interactive examples
        print("=== 지원자 데이터 처리 예제 ===\n")
        
        # Example 1: Process from Python dictionary (직접 딕셔너리로 입력)
        print("예제 1: Python 딕셔너리로 직접 입력")
        print("-" * 70)
        
        applicant_data = {
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
        
        process_single_json_dict(applicant_data, "applicants.xlsx")
        
        print("\n" + "="*70 + "\n")
        
        # Example 2: Process from JSON file (JSON 파일에서 읽기)
        print("예제 2: JSON 파일에서 읽기")
        print("-" * 70)
        print("사용법:")
        print('  process_single_json_file("applicant_001.json", "applicants.xlsx")')
        
        print("\n" + "="*70 + "\n")
        
        # Example 3: Process from JSON string (JSON 문자열로 입력)
        print("예제 3: JSON 문자열로 입력")
        print("-" * 70)
        
        json_str = '''
        {
            "applicant_name": "김영희",
            "application_date": "2024-12-19",
            "affiliation": "연세대학교",
            "application_field": "데이터 분석"
        }
        '''
        
        print("사용법:")
        print('  process_json_string(json_string, "applicants.xlsx")')
        
        print("\n" + "="*70 + "\n")
        
        # Example 4: Batch process from directory (디렉토리의 모든 JSON 파일 처리)
        print("예제 4: 디렉토리의 모든 JSON 파일 일괄 처리")
        print("-" * 70)
        print("사용법:")
        print('  batch_process_from_directory("json_outputs", "applicants.xlsx")')
        
        print("\n" + "="*70 + "\n")
        
        # Example 5: Command line usage
        print("예제 5: 커맨드 라인에서 실행")
        print("-" * 70)
        print("단일 파일:")
        print('  python main.py applicant_001.json')
        print('  python main.py applicant_001.json -o output.xlsx')
        print()
        print("디렉토리의 모든 파일:")
        print('  python main.py json_outputs/ -o applicants.xlsx')
        print('  python main.py -d json_outputs/')
        print()
        print("엄격 모드 (경고도 거부):")
        print('  python main.py applicant_001.json --strict')
        
        print("\n" + "="*70 + "\n")
        
        print("✓ 예제 실행 완료!")
        print(f"Excel 파일 생성됨: {os.path.abspath('applicants.xlsx')}")