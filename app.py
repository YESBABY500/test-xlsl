"""
PyQt desktop application for power allocation simulation.
Provides GUI interface for loading Excel data and running simulations.
"""

import sys
import os
import traceback
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                           QWidget, QPushButton, QTextEdit, QLabel, QMessageBox,
                           QProgressBar, QGroupBox, QGridLayout, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap

from data_loader import DataLoader
from allocation_engine import AllocationEngine


class SimulationWorker(QThread):
    """Worker thread for running allocation simulation."""
    
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    
    def __init__(self, base_path="."):
        super().__init__()
        self.base_path = base_path
    
    def run(self):
        try:
            self.progress.emit("正在加载数据文件...")
            
            # Load data
            loader = DataLoader(self.base_path)
            data = loader.load_all_data()
            
            self.progress.emit("数据加载完成，正在运行分配仿真...")
            
            # Run allocation simulation
            engine = AllocationEngine(data)
            results = engine.run_full_simulation()
            
            # Generate report
            report = engine.get_summary_report()
            results['report'] = report
            
            self.progress.emit("仿真完成！")
            self.finished.emit(results)
            
        except Exception as e:
            self.error.emit(str(e))


class AllocationApp(QMainWindow):
    """Main application window for power allocation simulation."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("电力分配仿真系统")
        self.setGeometry(100, 100, 1000, 700)
        
        # Initialize UI
        self.init_ui()
        
        # Check for auto-import
        self.check_auto_import()
    
    def init_ui(self):
        """Initialize the user interface."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title_label = QLabel("电力分配仿真系统")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Control panel
        control_group = QGroupBox("控制面板")
        control_layout = QHBoxLayout(control_group)
        
        self.import_btn = QPushButton("导入数据")
        self.import_btn.clicked.connect(self.import_data)
        control_layout.addWidget(self.import_btn)
        
        self.simulate_btn = QPushButton("运行仿真")
        self.simulate_btn.clicked.connect(self.run_simulation)
        self.simulate_btn.setEnabled(False)
        control_layout.addWidget(self.simulate_btn)
        
        self.clear_btn = QPushButton("清空结果")
        self.clear_btn.clicked.connect(self.clear_results)
        control_layout.addWidget(self.clear_btn)
        
        control_layout.addStretch()
        main_layout.addWidget(control_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("准备就绪")
        main_layout.addWidget(self.status_label)
        
        # Results area
        results_group = QGroupBox("仿真结果")
        results_layout = QVBoxLayout(results_group)
        
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setFont(QFont("Consolas", 10))
        results_layout.addWidget(self.results_text)
        
        main_layout.addWidget(results_group)
        
        # Initialize data
        self.data = None
        self.results = None
    
    def check_auto_import(self):
        """Check if test data file exists and auto-import if present."""
        test_file = "导入数据（测试）.xlsx"
        if os.path.exists(test_file):
            self.show_info("发现测试数据文件", f"发现文件 '{test_file}'，将自动导入数据并运行仿真。")
            self.auto_import_and_simulate()
    
    def auto_import_and_simulate(self):
        """Automatically import data and run simulation."""
        try:
            self.import_data()
            if self.data:
                self.run_simulation()
        except Exception as e:
            self.show_error("自动导入失败", f"自动导入和仿真失败：{str(e)}")
    
    def import_data(self):
        """Import data from Excel files."""
        try:
            self.status_label.setText("正在导入数据...")
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            
            # Load data
            loader = DataLoader(".")
            self.data = loader.load_all_data()
            
            self.progress_bar.setVisible(False)
            self.status_label.setText("数据导入成功")
            self.simulate_btn.setEnabled(True)
            
            # Show data summary
            summary = self.get_data_summary()
            self.results_text.setText(summary)
            
            self.show_info("导入成功", "Excel数据文件导入成功！可以开始运行仿真。")
            
        except ValueError as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText("数据导入失败")
            self.show_error("数据导入失败", str(e))
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText("导入出错")
            self.show_error("导入错误", f"未知错误：{str(e)}")
    
    def get_data_summary(self):
        """Generate a summary of imported data."""
        if not self.data:
            return "无数据"
        
        summary = """数据导入摘要
==============

"""
        
        summary += f"断面关系表: {len(self.data['section_mapping'])} 条记录\n"
        summary += f"电站信息表: {len(self.data['stations'])} 条记录\n"
        summary += f"负荷信息表: {len(self.data['loads'])} 条记录\n"
        summary += f"点位状态: {len(self.data['point_status'])} 个点位\n"
        summary += f"可选规则文件: {len(self.data['optional_rules'])} 个文件\n\n"
        
        # Show active/inactive points
        active_points = sum(1 for status in self.data['point_status'].values() if status == 1)
        inactive_points = len(self.data['point_status']) - active_points
        
        summary += f"激活点位: {active_points} 个\n"
        summary += f"非激活点位: {inactive_points} 个\n\n"
        
        # Show loaded rule files
        if self.data['optional_rules']:
            summary += "已加载的可选规则文件:\n"
            for filename in self.data['optional_rules'].keys():
                summary += f"- {filename}\n"
        
        summary += "\n准备运行仿真..."
        
        return summary
    
    def run_simulation(self):
        """Run the allocation simulation."""
        if not self.data:
            self.show_error("无数据", "请先导入数据文件。")
            return
        
        self.status_label.setText("正在运行仿真...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.simulate_btn.setEnabled(False)
        
        # Start simulation in worker thread
        self.simulation_worker = SimulationWorker(".")
        self.simulation_worker.finished.connect(self.on_simulation_finished)
        self.simulation_worker.error.connect(self.on_simulation_error)
        self.simulation_worker.progress.connect(self.on_simulation_progress)
        self.simulation_worker.start()
    
    def on_simulation_progress(self, message):
        """Handle simulation progress updates."""
        self.status_label.setText(message)
    
    def on_simulation_finished(self, results):
        """Handle simulation completion."""
        self.progress_bar.setVisible(False)
        self.simulate_btn.setEnabled(True)
        self.status_label.setText("仿真完成")
        self.results = results
        
        # Display results
        self.results_text.setText(results['report'])
        
        self.show_info("仿真完成", "电力分配仿真已成功完成！")
    
    def on_simulation_error(self, error_message):
        """Handle simulation errors."""
        self.progress_bar.setVisible(False)
        self.simulate_btn.setEnabled(True)
        self.status_label.setText("仿真失败")
        
        self.show_error("仿真失败", error_message)
    
    def clear_results(self):
        """Clear the results display."""
        self.results_text.clear()
        self.status_label.setText("结果已清空")
    
    def show_info(self, title, message):
        """Show information message box."""
        QMessageBox.information(self, title, message)
    
    def show_error(self, title, message):
        """Show error message box."""
        QMessageBox.critical(self, title, message)
    
    def closeEvent(self, event):
        """Handle application close event."""
        if hasattr(self, 'simulation_worker') and self.simulation_worker.isRunning():
            self.simulation_worker.terminate()
            self.simulation_worker.wait()
        event.accept()


def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("电力分配仿真系统")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("PowerGrid Simulation")
    
    try:
        # Create and show main window
        window = AllocationApp()
        window.show()
        
        # Start event loop
        sys.exit(app.exec_())
        
    except Exception as e:
        # Show critical error if GUI fails to start
        error_msg = f"应用程序启动失败：{str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        
        # Try to show error dialog
        try:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("启动错误")
            msg_box.setText("应用程序无法启动")
            msg_box.setDetailedText(error_msg)
            msg_box.exec_()
        except:
            pass
        
        sys.exit(1)


if __name__ == "__main__":
    main()