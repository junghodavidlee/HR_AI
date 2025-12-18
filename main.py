import json
import os
from pathlib import Path
from applicant_excel_writer import ApplicantExcelWriter
from validator import process_applicant_resume, ApplicantDataValidator, DataCleaner


def batch_process_resumes(json_files: list, 
                          excel_path: str = "applicants.xlsx",
                          strict_mode: bool = False):
    """
    Process multiple resume JSON files and append to Excel
    
    Args:
        json_files: List of paths to JSON files
        excel_path: Path to Excel file (will be created if doesn't exist)
        strict_mode: If True, reject any data with validation errors
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


# Example usage
if __name__ == "__main__":
    # Example 1: Process single resume
    writer = ApplicantExcelWriter("applicants.xlsx")
    
    single_resume = {
        # applicant_number will be auto-generated sequentially (1, 2, 3, ...)
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
    
    process_applicant_resume(single_resume, writer, strict_mode=False)
    
    # Example 2: Batch process multiple resumes from a directory
    # Uncomment the lines below to use batch processing
    # json_directory = "resume_json_outputs"
    # if os.path.exists(json_directory):
    #     json_files = [
    #         os.path.join(json_directory, f) 
    #         for f in os.listdir(json_directory) 
    #         if f.endswith('.json')
    #     ]
    #     
    #     batch_process_resumes(
    #         json_files=json_files,
    #         excel_path="applicants.xlsx",
    #         strict_mode=False
    #     )
