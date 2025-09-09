"""
Allocation logic engine for power distribution simulation.
Implements the table-driven allocation logic based on Excel data.
"""

import pandas as pd
from typing import Dict, List, Tuple, Any
from data_loader import DataLoader


class AllocationEngine:
    """Power allocation simulation engine."""
    
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.results = {}
        
    def get_active_points(self) -> List[str]:
        """Get list of active points (status = 1) from point status data."""
        active_points = []
        for point_id, status in self.data['point_status'].items():
            if status == 1:
                active_points.append(point_id)
        return active_points
    
    def get_station_allocation(self, station_id: str) -> Dict[str, Any]:
        """Get allocation for a specific station based on current point status."""
        stations_df = self.data['stations']
        station_data = stations_df[stations_df['电站编号'].astype(str) == station_id]
        
        if station_data.empty:
            return {'status': 'not_found', 'allocation': 0, 'reason': f'Station {station_id} not found'}
        
        station_info = station_data.iloc[0]
        point_id = str(station_info['点位'])
        
        # Check if station's point is active
        point_status = self.data['point_status'].get(point_id, 0)
        
        if point_status == 1:
            # Station is active - calculate allocation
            try:
                output = float(station_info['出力']) if pd.notna(station_info['出力']) and str(station_info['出力']).replace('.', '').isdigit() else 0.0
            except (ValueError, TypeError):
                output = 0.0
            
            return {
                'status': 'active',
                'allocation': output,
                'reason': f'Point {point_id} is active',
                'station_name': station_info['电站名称'],
                'point_id': point_id
            }
        else:
            return {
                'status': 'inactive',
                'allocation': 0,
                'reason': f'Point {point_id} is inactive',
                'station_name': station_info['电站名称'],
                'point_id': point_id
            }
    
    def get_load_allocation(self, load_id: str) -> Dict[str, Any]:
        """Get allocation for a specific load/user based on current point status."""
        loads_df = self.data['loads']
        load_data = loads_df[loads_df['用户编号'].astype(str) == load_id]
        
        if load_data.empty:
            return {'status': 'not_found', 'allocation': 0, 'reason': f'Load {load_id} not found'}
        
        load_info = load_data.iloc[0]
        point_id = str(load_info['点位'])
        
        # Check if load's point is active
        point_status = self.data['point_status'].get(point_id, 0)
        
        if point_status == 1:
            # Load is active - calculate allocation
            try:
                demand = float(load_info['负荷']) if pd.notna(load_info['负荷']) and str(load_info['负荷']).replace('.', '').isdigit() else 0.0
            except (ValueError, TypeError):
                demand = 0.0
            
            return {
                'status': 'active',
                'allocation': demand,
                'reason': f'Point {point_id} is active',
                'user_name': load_info['用电用户'],
                'point_id': point_id
            }
        else:
            return {
                'status': 'inactive',
                'allocation': 0,
                'reason': f'Point {point_id} is inactive',
                'user_name': load_info['用电用户'],
                'point_id': point_id
            }
    
    def calculate_section_allocation(self, section_name: str) -> Dict[str, Any]:
        """Calculate total allocation for a specific section."""
        section_df = self.data['section_mapping']
        section_points = section_df[section_df['点位'] == section_name]
        
        if section_points.empty:
            return {'total_generation': 0, 'total_load': 0, 'balance': 0, 'active_points': []}
        
        total_generation = 0
        total_load = 0
        active_points = []
        
        # Check all stations in this section
        stations_df = self.data['stations']
        loads_df = self.data['loads']
        
        for _, point_row in section_points.iterrows():
            point_id = str(point_row['断面名称'])  # Using section name as identifier
            
            # Find stations and loads for this point
            station_matches = stations_df[stations_df['点位'].astype(str) == point_id]
            load_matches = loads_df[loads_df['点位'].astype(str) == point_id]
            
            for _, station in station_matches.iterrows():
                allocation = self.get_station_allocation(str(station['电站编号']))
                if allocation['status'] == 'active':
                    total_generation += allocation['allocation']
                    active_points.append(f"Station {station['电站编号']}")
            
            for _, load in load_matches.iterrows():
                allocation = self.get_load_allocation(str(load['用户编号']))
                if allocation['status'] == 'active':
                    total_load += allocation['allocation']
                    active_points.append(f"Load {load['用户编号']}")
        
        balance = total_generation - total_load
        
        return {
            'section_name': section_name,
            'total_generation': total_generation,
            'total_load': total_load,
            'balance': balance,
            'active_points': active_points
        }
    
    def run_full_simulation(self) -> Dict[str, Any]:
        """Run complete allocation simulation."""
        active_points = self.get_active_points()
        
        # Get all station allocations
        station_results = []
        stations_df = self.data['stations']
        for _, station in stations_df.iterrows():
            allocation = self.get_station_allocation(str(station['电站编号']))
            station_results.append(allocation)
        
        # Get all load allocations
        load_results = []
        loads_df = self.data['loads']
        for _, load in loads_df.iterrows():
            allocation = self.get_load_allocation(str(load['用户编号']))
            load_results.append(allocation)
        
        # Calculate section-wise totals
        section_results = []
        sections_df = self.data['section_mapping']
        unique_sections = sections_df['点位'].unique()
        
        for section in unique_sections:
            if pd.notna(section):
                section_allocation = self.calculate_section_allocation(section)
                section_results.append(section_allocation)
        
        # Calculate overall totals
        total_active_generation = sum(s['allocation'] for s in station_results if s['status'] == 'active')
        total_active_load = sum(l['allocation'] for l in load_results if l['status'] == 'active')
        overall_balance = total_active_generation - total_active_load
        
        results = {
            'summary': {
                'total_active_points': len(active_points),
                'total_active_generation': total_active_generation,
                'total_active_load': total_active_load,
                'overall_balance': overall_balance,
                'active_points': active_points
            },
            'stations': station_results,
            'loads': load_results,
            'sections': section_results,
            'optional_rules_applied': len(self.data['optional_rules']) > 0
        }
        
        self.results = results
        return results
    
    def get_summary_report(self) -> str:
        """Generate a text summary of the allocation results."""
        if not self.results:
            return "No simulation results available. Run simulation first."
        
        summary = self.results['summary']
        
        report = f"""
电力分配仿真结果报告
=====================

总体概况:
- 激活点位数量: {summary['total_active_points']}
- 总发电量: {summary['total_active_generation']:.2f} MW
- 总负荷量: {summary['total_active_load']:.2f} MW
- 电力平衡: {summary['overall_balance']:.2f} MW

电站状态:
"""
        
        active_stations = [s for s in self.results['stations'] if s['status'] == 'active']
        inactive_stations = [s for s in self.results['stations'] if s['status'] == 'inactive']
        
        report += f"- 运行中电站: {len(active_stations)}个\n"
        report += f"- 停运电站: {len(inactive_stations)}个\n\n"
        
        if active_stations:
            report += "运行中电站详情:\n"
            for station in active_stations[:10]:  # Show first 10
                report += f"  - {station.get('station_name', 'Unknown')}: {station['allocation']:.2f} MW (点位: {station.get('point_id', 'N/A')})\n"
            if len(active_stations) > 10:
                report += f"  ... 还有 {len(active_stations) - 10} 个电站\n"
        
        report += f"\n负荷状态:\n"
        active_loads = [l for l in self.results['loads'] if l['status'] == 'active']
        inactive_loads = [l for l in self.results['loads'] if l['status'] == 'inactive']
        
        report += f"- 活动负荷: {len(active_loads)}个\n"
        report += f"- 非活动负荷: {len(inactive_loads)}个\n"
        
        if self.results['optional_rules_applied']:
            report += f"\n已加载可选规则文件: {len(self.data['optional_rules'])}个\n"
        
        return report


def test_allocation_engine():
    """Test function for the allocation engine."""
    try:
        # Load data
        loader = DataLoader()
        data = loader.load_all_data()
        
        # Create and run allocation engine
        engine = AllocationEngine(data)
        results = engine.run_full_simulation()
        
        print("✓ Allocation simulation completed successfully")
        print(f"✓ Total active points: {results['summary']['total_active_points']}")
        print(f"✓ Total generation: {results['summary']['total_active_generation']:.2f} MW")
        print(f"✓ Total load: {results['summary']['total_active_load']:.2f} MW")
        print(f"✓ Balance: {results['summary']['overall_balance']:.2f} MW")
        
        # Print summary report
        print("\n" + engine.get_summary_report())
        
        return True
        
    except Exception as e:
        print(f"✗ Allocation simulation failed: {e}")
        return False


if __name__ == "__main__":
    test_allocation_engine()