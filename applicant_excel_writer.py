import openpyxl
from openpyxl.styles import Font, Alignment
import json
from datetime import datetime
from typing import Dict, Any

class ApplicantExcelWriter:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.column_mapping = self._create_column_mapping()
        
    def _create_column_mapping(self) -> Dict[str, str]:
        """Define the mapping between JSON paths and Excel columns"""
        mapping = {
            # Basic information
            'applicant_number': 'A',
            'applicant_name': 'B',
            'application_date': 'C',
            'affiliation': 'D',
            'application_field': 'E',
            
            # Basic info section
            'basic_info.birth_year': 'F',
            'basic_info.gender': 'G',
            'basic_info.final_education_school': 'H',
            'basic_info.final_education_degree': 'I',
        }
        
        # Work experience mapping (경력 1-5)
        base_columns = ['J', 'P', 'V', 'AB', 'AH']  # Starting column for each experience
        exp_fields = ['start_date', 'end_date', 'company_name', 
                      'final_department', 'final_position', 'salary']
        
        for exp_idx in range(5):
            base_col_idx = self._column_to_index(base_columns[exp_idx])
            for field_offset, field in enumerate(exp_fields):
                col = self._index_to_column(base_col_idx + field_offset)
                mapping[f'work_experience[{exp_idx}].{field}'] = col
                
        return mapping
    
    def _column_to_index(self, col: str) -> int:
        """Convert Excel column (A, B, AA) to 0-based index"""
        result = 0
        for char in col:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def _index_to_column(self, idx: int) -> str:
        """Convert 0-based index to Excel column (A, B, AA)"""
        col = ''
        idx += 1
        while idx > 0:
            idx -= 1
            col = chr(idx % 26 + ord('A')) + col
            idx //= 26
        return col
    
    def _get_nested_value(self, data: Dict, path: str) -> Any:
        """Extract value from nested dictionary using dot notation and array indices"""
        keys = path.replace('[', '.').replace(']', '').split('.')
        value = data
        
        for key in keys:
            if key.isdigit():
                idx = int(key)
                if isinstance(value, list) and idx < len(value):
                    value = value[idx]
                else:
                    return None
            else:
                value = value.get(key) if isinstance(value, dict) else None
                
            if value is None:
                return None
                
        return value
    
    def create_template(self):
        """Create a new Excel file with headers"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "지원자 데이터"
        
        # Define headers
        headers = {
            'A1': '지원자 번호',
            'B1': '지원자명',
            'C1': '지원일',
            'D1': '소속',
            'E1': '지원분야/공고',
            'F1': '출생년도',
            'G1': '성별',
            'H1': '최종학력(학교)',
            'I1': '최종학력(학사 등)',
        }
        
        # Add work experience headers
        exp_labels = ['경력 1', '경력 2', '경력 3', '경력 4', '경력 5']
        exp_fields = ['입사년월', '퇴사년월', '회사명', '최종부서명', '최종직위', '연봉(천원)']
        base_columns = ['J', 'P', 'V', 'AB', 'AH']
        
        for exp_idx, (label, base_col) in enumerate(zip(exp_labels, base_columns)):
            base_idx = self._column_to_index(base_col)
            for field_offset, field_name in enumerate(exp_fields):
                col = self._index_to_column(base_idx + field_offset)
                headers[f'{col}1'] = f'{label} {field_name}'
        
        # Write headers
        header_font = Font(bold=True)
        for cell_ref, header_text in headers.items():
            ws[cell_ref] = header_text
            ws[cell_ref].font = header_font
            ws[cell_ref].alignment = Alignment(horizontal='center')
        
        # Adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column].width = adjusted_width
        
        wb.save(self.excel_path)
        print(f"Template created: {self.excel_path}")
    
    def append_applicant(self, json_data: Dict[str, Any]) -> int:
        """Append a new applicant to the Excel file"""
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
        except FileNotFoundError:
            print("Excel file not found. Creating new template...")
            self.create_template()
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
        
        # Find next empty row
        next_row = ws.max_row + 1
        
        # Auto-generate applicant_number if not provided
        if not json_data.get('applicant_number'):
            # Sequential number based on row number (row 2 = applicant 1)
            json_data['applicant_number'] = str(next_row - 1)
        
        # Write data to appropriate cells
        for json_path, excel_col in self.column_mapping.items():
            value = self._get_nested_value(json_data, json_path)
            cell_ref = f'{excel_col}{next_row}'
            
            # Handle None values
            if value is None:
                ws[cell_ref] = ''
            else:
                ws[cell_ref] = value
                
            # Center align all cells
            ws[cell_ref].alignment = Alignment(horizontal='center', vertical='center')
        
        wb.save(self.excel_path)
        print(f"Applicant added at row {next_row}: {json_data.get('applicant_name', 'Unknown')}")
        return next_row
    
    def batch_append(self, json_data_list: list):
        """Append multiple applicants at once (more efficient)"""
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
        except FileNotFoundError:
            print("Excel file not found. Creating new template...")
            self.create_template()
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
        
        start_row = ws.max_row + 1
        
        for idx, json_data in enumerate(json_data_list):
            current_row = start_row + idx
            
            # Auto-generate applicant_number if not provided
            if not json_data.get('applicant_number'):
                json_data['applicant_number'] = str(current_row - 1)
            
            for json_path, excel_col in self.column_mapping.items():
                value = self._get_nested_value(json_data, json_path)
                cell_ref = f'{excel_col}{current_row}'
                
                ws[cell_ref] = value if value is not None else ''
                ws[cell_ref].alignment = Alignment(horizontal='center', vertical='center')
        
        wb.save(self.excel_path)
        print(f"Batch added {len(json_data_list)} applicants (rows {start_row}-{start_row + len(json_data_list) - 1})")
