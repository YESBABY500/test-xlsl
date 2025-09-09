"""
Test script to validate the application components without GUI.
"""

import sys
import os

def test_imports():
    """Test that all modules can be imported."""
    try:
        from data_loader import DataLoader
        print("✓ DataLoader import successful")
        
        from allocation_engine import AllocationEngine
        print("✓ AllocationEngine import successful")
        
        # Test GUI imports without creating QApplication
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtCore import Qt
        print("✓ PyQt5 imports successful")
        
        return True
    except Exception as e:
        print(f"✗ Import failed: {e}")
        return False

def test_data_flow():
    """Test the complete data flow without GUI."""
    try:
        # Test data loading
        from data_loader import DataLoader
        loader = DataLoader(".")
        data = loader.load_all_data()
        print("✓ Data loading successful")
        
        # Test allocation engine
        from allocation_engine import AllocationEngine
        engine = AllocationEngine(data)
        results = engine.run_full_simulation()
        print("✓ Allocation simulation successful")
        
        # Test report generation
        report = engine.get_summary_report()
        print("✓ Report generation successful")
        print(f"✓ Report length: {len(report)} characters")
        
        return True
    except Exception as e:
        print(f"✗ Data flow test failed: {e}")
        return False

def test_error_handling():
    """Test error handling for missing files."""
    try:
        from data_loader import DataLoader
        
        # Test with non-existent directory
        loader = DataLoader("/non/existent/path")
        try:
            data = loader.load_all_data()
            print("✗ Should have failed with missing files")
            return False
        except ValueError as e:
            print(f"✓ Proper error handling for missing files: {str(e)[:100]}...")
            return True
            
    except Exception as e:
        print(f"✗ Error handling test failed: {e}")
        return False

def validate_required_files():
    """Validate that all required files exist."""
    required_files = [
        '断面关系表.xlsx',
        '电站信息表.xlsx', 
        '负荷信息表.xlsx',
        '导入数据（测试）.xlsx'
    ]
    
    missing_files = []
    for filename in required_files:
        if not os.path.exists(filename):
            missing_files.append(filename)
    
    if missing_files:
        print(f"✗ Missing required files: {missing_files}")
        return False
    else:
        print("✓ All required files present")
        return True

def main():
    """Run all tests."""
    print("电力分配仿真系统 - 组件测试")
    print("=" * 40)
    
    all_passed = True
    
    # Test file existence
    if not validate_required_files():
        all_passed = False
    
    # Test imports
    if not test_imports():
        all_passed = False
    
    # Test data flow
    if not test_data_flow():
        all_passed = False
    
    # Test error handling
    if not test_error_handling():
        all_passed = False
    
    print("\n" + "=" * 40)
    if all_passed:
        print("✓ 所有测试通过！应用程序准备就绪。")
        print("\n使用方法:")
        print("python app.py  # 启动GUI应用程序")
        return 0
    else:
        print("✗ 某些测试失败，请检查配置。")
        return 1

if __name__ == "__main__":
    sys.exit(main())