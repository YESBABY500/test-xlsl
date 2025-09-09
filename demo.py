"""
Demo script showing the complete application workflow.
"""

from data_loader import DataLoader
from allocation_engine import AllocationEngine

def demo_complete_workflow():
    """Demonstrate the complete application workflow."""
    print("电力分配仿真系统演示")
    print("=" * 50)
    
    print("\n1. 正在检查所需文件...")
    required_files = [
        '断面关系表.xlsx',
        '电站信息表.xlsx', 
        '负荷信息表.xlsx',
        '导入数据（测试）.xlsx'
    ]
    
    import os
    for filename in required_files:
        if os.path.exists(filename):
            print(f"   ✓ {filename}")
        else:
            print(f"   ✗ {filename}")
    
    print("\n2. 正在导入数据...")
    try:
        loader = DataLoader(".")
        data = loader.load_all_data()
        
        print(f"   ✓ 断面关系表: {len(data['section_mapping'])} 条记录")
        print(f"   ✓ 电站信息表: {len(data['stations'])} 条记录") 
        print(f"   ✓ 负荷信息表: {len(data['loads'])} 条记录")
        print(f"   ✓ 点位状态: {len(data['point_status'])} 个点位")
        print(f"   ✓ 可选规则文件: {len(data['optional_rules'])} 个")
        
        # Show active points
        active_points = sum(1 for status in data['point_status'].values() if status == 1)
        print(f"   ✓ 激活点位: {active_points}/{len(data['point_status'])}")
        
    except ValueError as e:
        print(f"   ✗ 数据导入失败: {e}")
        return
    
    print("\n3. 正在运行仿真...")
    try:
        engine = AllocationEngine(data)
        results = engine.run_full_simulation()
        
        summary = results['summary']
        print(f"   ✓ 仿真完成")
        print(f"   ✓ 总发电量: {summary['total_active_generation']:.2f} MW")
        print(f"   ✓ 总负荷量: {summary['total_active_load']:.2f} MW")
        print(f"   ✓ 电力平衡: {summary['overall_balance']:.2f} MW")
        
        active_stations = sum(1 for s in results['stations'] if s['status'] == 'active')
        active_loads = sum(1 for l in results['loads'] if l['status'] == 'active')
        
        print(f"   ✓ 运行中电站: {active_stations}个")
        print(f"   ✓ 活动负荷: {active_loads}个")
        
    except Exception as e:
        print(f"   ✗ 仿真失败: {e}")
        return
    
    print("\n4. 生成详细报告...")
    report = engine.get_summary_report()
    
    print("\n" + "=" * 50)
    print("完整仿真报告:")
    print("=" * 50)
    print(report)
    
    print("\n" + "=" * 50)
    print("演示完成！")
    print("\n在有GUI环境下，可以运行: python app.py")
    print("这将启动完整的PyQt桌面应用程序。")

if __name__ == "__main__":
    demo_complete_workflow()