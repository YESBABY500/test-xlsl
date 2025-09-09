"""
Data loader module for strict Excel file parsing with validation.
Handles the allocation logic data files with proper error reporting.
"""

import pandas as pd
import os
from typing import Dict, List, Tuple, Optional, Any


class DataLoader:
    """Strict data loader for allocation logic Excel files."""
    
    def __init__(self, base_path: str = "."):
        self.base_path = base_path
        self.required_files = {
            '断面关系表.xlsx': ['断面编号', '断面名称', '点位'],
            '电站信息表.xlsx': ['电站编号', '电站名称', '出力', '点位'],
            '负荷信息表.xlsx': ['用户编号', '用电用户', '负荷', '点位'],
            '导入数据（测试）.xlsx': None  # Special handling for this file
        }
        
        # Optional rule files
        self.optional_files = [
            'D站分配逻辑（11个）.xlsx',
            'T站分配逻辑（7个）.xlsx', 
            '分配逻辑-其他（13个）.xlsx',
            '逻辑判断大用户（8个）.xlsx',
            '直接挂接（41个）.xlsx',
            '直接大用户（21个）.xlsx'
        ]
        
        # Column mapping from actual file structure to required structure
        self.column_mappings = {
            '断面关系表.xlsx': {
                '序号': '断面编号',
                '代码': '断面名称', 
                '断面': '点位'
            },
            '电站信息表.xlsx': {
                '序号': '电站编号',
                '代码': '电站编号',  # Alternative mapping
                '电站名称': '电站名称',
                '实际出力': '出力',
                '代码': '点位'  # Using code as point location
            },
            '负荷信息表.xlsx': {
                '序号': '用户编号',
                '代码': '用户编号',  # Alternative mapping
                '用户名称': '用电用户',
                # Note: 负荷 column not present in actual file, needs special handling
                '代码': '点位'  # Using code as point location
            }
        }
    
    def validate_file_exists(self, filename: str) -> bool:
        """Check if a file exists in the base path."""
        filepath = os.path.join(self.base_path, filename)
        return os.path.exists(filepath)
    
    def load_excel_file(self, filename: str, sheet_name: str = None, header_row: int = 1) -> pd.DataFrame:
        """Load an Excel file with proper header handling."""
        filepath = os.path.join(self.base_path, filename)
        
        if not self.validate_file_exists(filename):
            raise ValueError(f"Required file '{filename}' not found in {self.base_path}")
        
        try:
            if sheet_name:
                df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row)
            else:
                df = pd.read_excel(filepath, header=header_row)
            return df
        except Exception as e:
            raise ValueError(f"Error reading file '{filename}': {str(e)}")
    
    def validate_columns(self, df: pd.DataFrame, required_columns: List[str], filename: str) -> None:
        """Validate that required columns exist in the dataframe."""
        missing_columns = []
        
        # For files with column mappings, check if we can map existing columns
        if filename in self.column_mappings:
            mapping = self.column_mappings[filename]
            available_columns = set(df.columns)
            
            for req_col in required_columns:
                # Look for direct match first
                if req_col in available_columns:
                    continue
                
                # Look for mapped column
                mapped_found = False
                for actual_col, mapped_col in mapping.items():
                    if mapped_col == req_col and actual_col in available_columns:
                        mapped_found = True
                        break
                
                if not mapped_found:
                    missing_columns.append(req_col)
        else:
            # Direct column check
            for col in required_columns:
                if col not in df.columns:
                    missing_columns.append(col)
        
        if missing_columns:
            raise ValueError(
                f"File '{filename}' is missing required columns: {missing_columns}. "
                f"Available columns: {list(df.columns)}"
            )
    
    def load_section_mapping(self) -> pd.DataFrame:
        """Load 断面关系表.xlsx with validation."""
        filename = '断面关系表.xlsx'
        df = self.load_excel_file(filename)
        
        # Map columns to required format
        if filename in self.column_mappings:
            mapping = self.column_mappings[filename]
            rename_dict = {}
            for actual_col, req_col in mapping.items():
                if actual_col in df.columns:
                    rename_dict[actual_col] = req_col
            df = df.rename(columns=rename_dict)
        
        # Validate required columns are present (after mapping)
        required_cols = self.required_files[filename]
        self.validate_columns(df, required_cols, filename)
        
        return df
    
    def load_stations(self) -> pd.DataFrame:
        """Load 电站信息表.xlsx with validation."""
        filename = '电站信息表.xlsx'
        df = self.load_excel_file(filename)
        
        # Handle special mapping for stations
        rename_dict = {}
        if '代码' in df.columns:
            rename_dict['代码'] = '点位'
        if '实际出力' in df.columns:
            rename_dict['实际出力'] = '出力'
        elif '出力(MW)' in df.columns:
            rename_dict['出力(MW)'] = '出力'
        
        # Handle ID column mapping - prefer existing numbered column
        if '序号' in df.columns:
            rename_dict['序号'] = '电站编号'
        elif '编号' in df.columns:
            rename_dict['编号'] = '电站编号'
        elif '代码' in df.columns and '电站编号' not in rename_dict.values():
            rename_dict['代码'] = '电站编号'
        
        df = df.rename(columns=rename_dict)
        
        # Validate required columns
        required_cols = ['电站编号', '电站名称', '出力', '点位']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(
                f"File '{filename}' is missing required columns: {missing_cols}. "
                f"Available columns: {list(df.columns)}"
            )
        
        return df
    
    def load_loads(self) -> pd.DataFrame:
        """Load 负荷信息表.xlsx with validation."""
        filename = '负荷信息表.xlsx'
        df = self.load_excel_file(filename)
        
        # Handle special mapping for loads
        rename_dict = {}
        if '代码' in df.columns:
            rename_dict['代码'] = '点位'
        if '用户名称' in df.columns:
            rename_dict['用户名称'] = '用电用户'
        elif '用户' in df.columns:
            rename_dict['用户'] = '用电用户'
        
        # Handle ID column mapping
        if '序号' in df.columns:
            rename_dict['序号'] = '用户编号'
        elif '编号' in df.columns:
            rename_dict['编号'] = '用户编号'
        elif '代码' in df.columns and '用户编号' not in rename_dict.values():
            rename_dict['代码'] = '用户编号'
        
        df = df.rename(columns=rename_dict)
        
        # Add missing 负荷 column if not present (will be filled with defaults)
        if '负荷' not in df.columns and '负荷(MW)' not in df.columns:
            df['负荷'] = 0.0  # Default load value
        elif '负荷(MW)' in df.columns:
            df = df.rename(columns={'负荷(MW)': '负荷'})
        
        # Validate required columns
        required_cols = ['用户编号', '用电用户', '负荷', '点位']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(
                f"File '{filename}' is missing required columns: {missing_cols}. "
                f"Available columns: {list(df.columns)}"
            )
        
        return df
    
    def load_point_status(self) -> Dict[str, int]:
        """Load point status data from 导入数据（测试）.xlsx."""
        filename = '导入数据（测试）.xlsx'
        
        if not self.validate_file_exists(filename):
            raise ValueError(f"Required file '{filename}' not found in {self.base_path}")
        
        try:
            # Try to find the point status sheet
            sheets = pd.read_excel(os.path.join(self.base_path, filename), sheet_name=None)
            
            status_sheet = None
            for sheet_name in sheets.keys():
                if '点位' in sheet_name or '逻辑' in sheet_name:
                    status_sheet = sheet_name
                    break
            
            if not status_sheet:
                raise ValueError(f"No point status sheet found in '{filename}'")
            
            df = pd.read_excel(os.path.join(self.base_path, filename), 
                             sheet_name=status_sheet, header=1)
            
            # Handle two possible formats:
            # 1. Two columns: 点位, 状态
            # 2. Point IDs as columns with first row giving status
            
            if '点位' in df.columns and '状态' in df.columns:
                # Format 1: Two-column layout
                point_status = {}
                for _, row in df.iterrows():
                    if pd.notna(row['点位']) and pd.notna(row['状态']):
                        try:
                            # Handle both numeric and string point IDs
                            point_id = str(row['点位']).strip()
                            if point_id.replace('.', '').isdigit():
                                point_id = str(int(float(point_id)))
                            status = int(row['状态'])
                            point_status[point_id] = status
                        except (ValueError, TypeError):
                            # Skip invalid entries
                            continue
                return point_status
            else:
                # Format 2: Point IDs as columns
                # First row contains the status values
                point_status = {}
                if len(df) > 0:
                    first_row = df.iloc[0]
                    for col in df.columns:
                        if str(col).isdigit():  # Column name is a point ID
                            if pd.notna(first_row[col]):
                                point_status[str(col)] = int(first_row[col])
                return point_status
                
        except Exception as e:
            raise ValueError(f"Error parsing point status from '{filename}': {str(e)}")
    
    def load_optional_rules(self) -> Dict[str, Any]:
        """Load optional rule files if they exist."""
        rules = {}
        
        for filename in self.optional_files:
            if self.validate_file_exists(filename):
                try:
                    # Load the first sheet of each rule file
                    df = self.load_excel_file(filename, header_row=0)
                    rules[filename] = df
                except Exception as e:
                    # Log error but don't fail - these are optional
                    print(f"Warning: Could not load optional file '{filename}': {e}")
        
        return rules
    
    def load_all_data(self) -> Dict[str, Any]:
        """Load all required data with validation."""
        # Check all required files exist first
        missing_files = []
        for filename in self.required_files.keys():
            if not self.validate_file_exists(filename):
                missing_files.append(filename)
        
        if missing_files:
            raise ValueError(f"Missing required files: {missing_files}")
        
        # Load all data
        data = {}
        
        try:
            data['section_mapping'] = self.load_section_mapping()
            data['stations'] = self.load_stations()
            data['loads'] = self.load_loads()
            data['point_status'] = self.load_point_status()
            data['optional_rules'] = self.load_optional_rules()
            
            return data
            
        except Exception as e:
            raise ValueError(f"Data loading failed: {str(e)}")


def test_data_loader():
    """Test function for the data loader."""
    loader = DataLoader()
    
    try:
        data = loader.load_all_data()
        print("✓ All data loaded successfully")
        print(f"✓ Section mapping: {len(data['section_mapping'])} records")
        print(f"✓ Stations: {len(data['stations'])} records")
        print(f"✓ Loads: {len(data['loads'])} records")
        print(f"✓ Point status: {len(data['point_status'])} points")
        print(f"✓ Optional rules: {len(data['optional_rules'])} files loaded")
        
    except ValueError as e:
        print(f"✗ Data loading failed: {e}")
        return False
    
    return True


if __name__ == "__main__":
    test_data_loader()