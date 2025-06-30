import sys
import os
import json
import pandas as pd
import numpy as np
import xlrd


from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QFileDialog, QListWidget, QLabel, 
                           QMessageBox, QCheckBox, QHBoxLayout, QDialog,
                           QLineEdit, QListWidgetItem, QTableWidget, 
                           QTableWidgetItem, QHeaderView, QComboBox,
                           QSpinBox, QDialogButtonBox, QInputDialog, QMenu,
                           QGroupBox, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt

class SaveConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('保存配置')
        self.setGeometry(300, 300, 400, 150)
        
        layout = QVBoxLayout()
        
        # 添加名称输入
        layout.addWidget(QLabel('配置名称:'))
        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)
        
        # 添加描述输入
        layout.addWidget(QLabel('配置描述 (可选):'))
        self.desc_input = QLineEdit()
        layout.addWidget(self.desc_input)
        
        # 添加按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def get_config_info(self):
        return {
            'name': self.name_input.text(),
            'description': self.desc_input.text()
        }

class LoadConfigDialog(QDialog):
    def __init__(self, configs, parent=None):
        super().__init__(parent)
        self.configs = configs
        self.selected_config = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('加载配置')
        self.setGeometry(300, 300, 600, 400)
        
        layout = QVBoxLayout()
        
        # 添加说明标签
        layout.addWidget(QLabel('选择要加载的配置:'))
        
        # 创建配置列表
        self.config_table = QTableWidget()
        self.config_table.setColumnCount(3)
        self.config_table.setHorizontalHeaderLabels(['配置名称', '创建日期', '描述'])
        
        # 设置列宽
        header = self.config_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        
        # 填充配置列表
        self.config_table.setRowCount(len(self.configs))
        for i, (config_id, config) in enumerate(self.configs.items()):
            self.config_table.setItem(i, 0, QTableWidgetItem(config['name']))
            self.config_table.setItem(i, 1, QTableWidgetItem(config.get('date', '')))
            self.config_table.setItem(i, 2, QTableWidgetItem(config.get('description', '')))
            # 存储配置ID为用户数据
            self.config_table.item(i, 0).setData(Qt.UserRole, config_id)
        
        self.config_table.itemDoubleClicked.connect(self.accept)
        layout.addWidget(self.config_table)
        
        # 添加按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.on_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def on_accept(self):
        """当用户点击确定按钮时调用"""
        selected_items = self.config_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, '警告', '请选择一个配置!')
            return
        
        # 获取所选行的第一列
        row = selected_items[0].row()
        config_item = self.config_table.item(row, 0)
        self.selected_config = config_item.data(Qt.UserRole)
        self.accept()
    
    def get_selected_config(self):
        return self.selected_config

class SortRowsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('行排序设置')
        self.setGeometry(300, 300, 400, 180)
        self.sort_settings = {
            'method': 'valid_count',  # 默认按有效元素数量排序
            'order': 'descending',    # 默认降序
        }
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        # 排序方法
        method_group = QGroupBox('排序方法')
        method_layout = QVBoxLayout()
        
        self.valid_count_radio = QRadioButton('按有效元素数量排序')
        self.valid_count_radio.setChecked(True)
        self.valid_count_radio.toggled.connect(self.update_settings)
        method_layout.addWidget(self.valid_count_radio)
        
        method_group.setLayout(method_layout)
        layout.addWidget(method_group)
        
        # 排序顺序
        order_group = QGroupBox('排序顺序')
        order_layout = QVBoxLayout()
        
        self.ascending_radio = QRadioButton('升序 (少 → 多)')
        self.descending_radio = QRadioButton('降序 (多 → 少)')
        self.descending_radio.setChecked(True)
        
        self.ascending_radio.toggled.connect(self.update_settings)
        self.descending_radio.toggled.connect(self.update_settings)
        
        order_layout.addWidget(self.ascending_radio)
        order_layout.addWidget(self.descending_radio)
        
        order_group.setLayout(order_layout)
        layout.addWidget(order_group)
        
        # 添加按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def update_settings(self):
        # 更新排序方法
        if self.valid_count_radio.isChecked():
            self.sort_settings['method'] = 'valid_count'
        
        # 更新排序顺序
        if self.ascending_radio.isChecked():
            self.sort_settings['order'] = 'ascending'
        else:
            self.sort_settings['order'] = 'descending'
    
    def get_sort_settings(self):
        return self.sort_settings

class PreviewDataDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.data = data
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('数据预览')
        self.setGeometry(100, 100, 900, 600)
        
        layout = QVBoxLayout()
        
        # 添加控制面板
        control_layout = QHBoxLayout()
        
        # 预览行数选择
        control_layout.addWidget(QLabel('预览行数:'))
        self.row_spinbox = QSpinBox()
        self.row_spinbox.setRange(1, min(100, len(self.data)))
        self.row_spinbox.setValue(min(10, len(self.data)))
        self.row_spinbox.valueChanged.connect(self.update_preview)
        control_layout.addWidget(self.row_spinbox)
        
        # 预览位置选择
        control_layout.addWidget(QLabel('预览位置:'))
        self.position_combo = QComboBox()
        self.position_combo.addItems(['开头', '中间', '结尾'])
        self.position_combo.currentIndexChanged.connect(self.update_preview)
        control_layout.addWidget(self.position_combo)
        
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # 设置为只读
        
        # 设置表格
        layout.addWidget(self.table)
        
        # 添加按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
        # 初始加载预览数据
        self.update_preview()
    
    def update_preview(self):
        """更新预览表格"""
        rows_to_show = self.row_spinbox.value()
        position = self.position_combo.currentText()
        
        # 根据选择的位置获取数据
        if position == '开头':
            preview_data = self.data.head(rows_to_show)
        elif position == '结尾':
            preview_data = self.data.tail(rows_to_show)
        else:  # 中间
            middle_idx = len(self.data) // 2
            start_idx = max(0, middle_idx - rows_to_show // 2)
            preview_data = self.data.iloc[start_idx:start_idx + rows_to_show]
        
        # 更新表格
        self.table.setRowCount(len(preview_data))
        self.table.setColumnCount(len(preview_data.columns))
        self.table.setHorizontalHeaderLabels(preview_data.columns)
        
        # 填充数据
        for row in range(len(preview_data)):
            for col in range(len(preview_data.columns)):
                value = str(preview_data.iloc[row, col])
                item = QTableWidgetItem(value)
                self.table.setItem(row, col, item)
        
        # 调整列宽
        header = self.table.horizontalHeader()
        for col in range(len(preview_data.columns)):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)

class ExportValuesDialog(QDialog):
    def __init__(self, field_name, values, parent=None):
        super().__init__(parent)
        self.field_name = field_name
        self.values = values
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle('导出初始值')
        self.setGeometry(200, 200, 400, 150)
        
        layout = QVBoxLayout()
        
        # 添加说明标签
        layout.addWidget(QLabel(f'将字段 "{self.field_name}" 的所有唯一值导出为替换规则模板文件'))
        
        # 添加按钮
        button_layout = QHBoxLayout()
        
        export_excel_button = QPushButton('导出为Excel')
        export_excel_button.clicked.connect(lambda: self.export_values('excel'))
        
        export_csv_button = QPushButton('导出为CSV')
        export_csv_button.clicked.connect(lambda: self.export_values('csv'))
        
        cancel_button = QPushButton('取消')
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(export_excel_button)
        button_layout.addWidget(export_csv_button)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def export_values(self, format_type):
        # 创建数据框
        df = pd.DataFrame({
            '原始值': self.values,
            '新值': self.values  # 初始时新值与原始值相同
        })
        
        # 获取保存路径
        file_filters = 'Excel files (*.xlsx)' if format_type == 'excel' else 'CSV files (*.csv)'
        file_ext = '.xlsx' if format_type == 'excel' else '.csv'
        
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            '保存替换规则模板',
            f'{self.field_name}_替换规则模板{file_ext}',
            file_filters
        )
        
        if file_name:
            try:
                if format_type == 'excel':
                    df.to_excel(file_name, index=False)
                else:
                    df.to_csv(file_name, index=False, encoding='utf-8-sig')
                
                QMessageBox.information(self, '成功', '替换规则模板已导出！')
                self.accept()
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'导出文件时出错：{str(e)}')

class ImportRulesDialog(QDialog):
    def __init__(self, field_names, parent=None):
        super().__init__(parent)
        self.field_names = field_names
        self.rules_data = None
        self.selected_field = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('导入替换规则')
        self.setGeometry(200, 200, 400, 200)
        
        layout = QVBoxLayout()
        
        # 添加说明标签
        layout.addWidget(QLabel('请选择替换规则文件(Excel或CSV)：\n文件应包含"原始值"和"新值"两列'))
        
        # 添加文件选择按钮
        self.file_button = QPushButton('选择规则文件')
        self.file_button.clicked.connect(self.select_file)
        layout.addWidget(self.file_button)
        
        # 添加字段选择下拉框
        layout.addWidget(QLabel('选择要应用规则的字段：'))
        self.field_combo = QComboBox()
        self.field_combo.addItems(self.field_names)
        layout.addWidget(self.field_combo)
        
        # 添加按钮
        button_layout = QHBoxLayout()
        
        self.import_button = QPushButton('导入规则')
        self.import_button.clicked.connect(self.accept)
        self.import_button.setEnabled(False)
        
        cancel_button = QPushButton('取消')
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(self.import_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            '选择规则文件',
            '',
            'Excel files (*.xlsx *.xls);;CSV files (*.csv)'
        )
        
        if file_name:
            try:
                # 读取规则文件
                if file_name.endswith('.csv'):
                    self.rules_data = pd.read_csv(file_name)
                else:
                    self.rules_data = pd.read_excel(file_name, engine='openpyxl')
                
                # 验证文件格式
                required_columns = ['原始值', '新值']
                if not all(col in self.rules_data.columns for col in required_columns):
                    QMessageBox.warning(self, '警告', '文件格式错误！必须包含"原始值"和"新值"两列。')
                    self.rules_data = None
                    return
                
                self.import_button.setEnabled(True)
                QMessageBox.information(self, '成功', '规则文件加载成功！')
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'加载规则文件时出错：{str(e)}')
                self.rules_data = None
    
    def get_rules_and_field(self):
        if self.rules_data is not None:
            return {
                'field': self.field_combo.currentText(),
                'rules': dict(zip(
                    self.rules_data['原始值'].astype(str),
                    self.rules_data['新值'].astype(str)
                ))
            }
        return None

class EditValuesDialog(QDialog):
    def __init__(self, field_name, values, parent=None):
        super().__init__(parent)
        self.field_name = field_name
        self.values = values
        self.edited_values = {}  # 存储修改后的值映射
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle(f'修改字段值 - {self.field_name}')
        self.setGeometry(200, 200, 600, 400)
        
        layout = QVBoxLayout()
        
        # 添加说明标签
        layout.addWidget(QLabel(f'字段 "{self.field_name}" 的所有唯一值：'))
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['原始值', '新值'])
        
        # 设置表格列宽
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        
        # 填充表格
        self.populate_table()
        
        layout.addWidget(self.table)
        
        # 添加按钮布局
        button_layout = QHBoxLayout()
        
        # 添加导出按钮
        export_button = QPushButton('导出初始值')
        export_button.clicked.connect(self.export_initial_values)
        button_layout.addWidget(export_button)
        
        # 添加导入规则按钮
        import_rules_button = QPushButton('从文件导入替换规则')
        import_rules_button.clicked.connect(self.import_rules)
        button_layout.addWidget(import_rules_button)
        
        layout.addLayout(button_layout)
        
        # 添加确认和取消按钮
        confirm_layout = QHBoxLayout()
        
        apply_button = QPushButton('应用修改')
        apply_button.clicked.connect(self.apply_changes)
        
        cancel_button = QPushButton('取消')
        cancel_button.clicked.connect(self.reject)
        
        confirm_layout.addWidget(apply_button)
        confirm_layout.addWidget(cancel_button)
        layout.addLayout(confirm_layout)
        
        self.setLayout(layout)
    
    def export_initial_values(self):
        """导出初始值为替换规则模板"""
        dialog = ExportValuesDialog(self.field_name, self.values, self)
        dialog.exec_()
    
    def populate_table(self):
        """填充表格数据"""
        self.table.setRowCount(len(self.values))
        for i, value in enumerate(self.values):
            # 原始值（不可编辑）
            original_item = QTableWidgetItem(str(value))
            original_item.setFlags(original_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(i, 0, original_item)
            
            # 新值（可编辑）
            new_value_item = QTableWidgetItem(str(value))
            self.table.setItem(i, 1, new_value_item)
    
    def import_rules(self):
        """导入替换规则"""
        dialog = ImportRulesDialog([self.field_name], self)
        if dialog.exec_() == QDialog.Accepted:
            result = dialog.get_rules_and_field()
            if result and result['rules']:
                # 更新表格中的值
                for row in range(self.table.rowCount()):
                    original_value = self.table.item(row, 0).text()
                    if original_value in result['rules']:
                        self.table.item(row, 1).setText(result['rules'][original_value])
                
                QMessageBox.information(self, '成功', '替换规则已应用到表格')
    
    def apply_changes(self):
        # 收集修改后的值
        for row in range(self.table.rowCount()):
            original_value = self.table.item(row, 0).text()
            new_value = self.table.item(row, 1).text()
            if original_value != new_value:
                self.edited_values[original_value] = new_value
        
        if self.edited_values:
            self.accept()
        else:
            self.reject()
    
    def get_edited_values(self):
        return self.edited_values

class RenameDialog(QDialog):
    def __init__(self, old_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle('重命名字段')
        self.setModal(True)
        
        layout = QVBoxLayout()
        
        # 添加说明标签
        layout.addWidget(QLabel(f'当前字段名: {old_name}'))
        layout.addWidget(QLabel('请输入新的字段名:'))
        
        # 添加输入框
        self.name_input = QLineEdit()
        self.name_input.setText(old_name)
        layout.addWidget(self.name_input)
        
        # 添加确认和取消按钮
        button_layout = QHBoxLayout()
        confirm_button = QPushButton('确认')
        cancel_button = QPushButton('取消')
        
        confirm_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(confirm_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def get_new_name(self):
        return self.name_input.text()

class FieldListItem(QWidget):
    def __init__(self, field_name, parent=None):
        super().__init__(parent)
        self.field_name = field_name
        self.display_name = field_name
        self.setup_ui()
        
    def setup_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 2, 5, 2)
        
        # 复选框
        self.checkbox = QCheckBox(self.display_name)
        layout.addWidget(self.checkbox)
        
        # 向上移动按钮
        self.up_button = QPushButton('↑')
        self.up_button.setFixedWidth(30)
        layout.addWidget(self.up_button)
        
        # 向下移动按钮
        self.down_button = QPushButton('↓')
        self.down_button.setFixedWidth(30)
        layout.addWidget(self.down_button)
        
        # 重命名按钮
        self.rename_button = QPushButton('重命名')
        layout.addWidget(self.rename_button)
        
        # 编辑值按钮
        self.edit_values_button = QPushButton('编辑值')
        layout.addWidget(self.edit_values_button)
        
        self.setLayout(layout)

class FileInfoSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.dataset = None
        self.available_columns = []
        self.field_mapping = {}  # 用于存储字段重命名映射
        self.value_mapping = {}  # 用于存储值映射
        self.configs = {}  # 用于存储配置
        self.config_file = 'file_info_system_configs.json'
        
        # 加载已保存的配置
        self.load_configs()
        
    def initUI(self):
        self.setWindowTitle('文件信息识别系统')
        self.setGeometry(100, 100, 1000, 600)
        
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        
        # 添加选择数据集按钮
        file_buttons_layout = QHBoxLayout()
        
        self.load_button = QPushButton('选择数据集文件', self)
        self.load_button.clicked.connect(self.load_dataset)
        file_buttons_layout.addWidget(self.load_button)
        
        # 添加配置按钮
        self.config_button = QPushButton('配置', self)
        self.config_button.setContextMenuPolicy(Qt.CustomContextMenu)
        self.config_button.customContextMenuRequested.connect(self.show_config_menu)
        file_buttons_layout.addWidget(self.config_button)
        
        layout.addLayout(file_buttons_layout)
        
        # 添加可用字段列表标签
        self.fields_label = QLabel('可选信息种类：', self)
        layout.addWidget(self.fields_label)
        
        # 添加可选字段列表
        self.fields_list = QListWidget(self)
        layout.addWidget(self.fields_list)
        
        # 添加按钮布局
        button_layout = QHBoxLayout()
        
        # 添加预览按钮
        self.preview_button = QPushButton('预览数据', self)
        self.preview_button.clicked.connect(self.preview_data)
        self.preview_button.setEnabled(False)
        button_layout.addWidget(self.preview_button)
        
        # 添加排序按钮
        self.sort_button = QPushButton('行排序', self)
        self.sort_button.clicked.connect(self.sort_rows)
        self.sort_button.setEnabled(False)
        button_layout.addWidget(self.sort_button)
        
        # 添加导出按钮
        self.export_button = QPushButton('导出选中字段到Excel', self)
        self.export_button.clicked.connect(self.export_to_excel)
        self.export_button.setEnabled(False)
        button_layout.addWidget(self.export_button)
        
        layout.addLayout(button_layout)
        
        central_widget.setLayout(layout)
    
    def show_config_menu(self, pos):
        """显示配置菜单"""
        menu = QMenu(self)
        
        save_action = menu.addAction('保存当前配置')
        load_action = menu.addAction('加载配置')
        manage_action = menu.addAction('管理配置')
        
        # 在按钮位置显示菜单
        action = menu.exec_(self.config_button.mapToGlobal(pos))
        
        if action == save_action:
            self.save_config()
        elif action == load_action:
            self.load_config()
        elif action == manage_action:
            self.manage_configs()
    
    def load_configs(self):
        """从文件加载所有配置"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.configs = json.load(f)
            except Exception as e:
                QMessageBox.warning(self, '警告', f'加载配置文件时出错: {str(e)}')
                self.configs = {}
    
    def save_configs(self):
        """保存所有配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.configs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, '错误', f'保存配置文件时出错: {str(e)}')
    
    def save_config(self):
        """保存当前配置"""
        if self.dataset is None:
            QMessageBox.warning(self, '警告', '请先加载数据集!')
            return
        
        # 获取配置名称和描述
        dialog = SaveConfigDialog(self)
        if dialog.exec_() != QDialog.Accepted:
            return
        
        config_info = dialog.get_config_info()
        if not config_info['name']:
            QMessageBox.warning(self, '警告', '配置名称不能为空!')
            return
        
        # 获取当前配置状态
        selected_fields = []
        field_states = {}
        
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            field_widget = self.fields_list.itemWidget(item)
            
            if field_widget:
                field_name = field_widget.field_name
                display_name = field_widget.display_name
                is_checked = field_widget.checkbox.isChecked()
                
                field_states[field_name] = {
                    'display_name': display_name,
                    'is_checked': is_checked,
                    'order': i
                }
                
                if is_checked:
                    selected_fields.append(field_name)
        
        # 创建配置对象
        from datetime import datetime
        config_id = f"config_{len(self.configs) + 1}_{int(datetime.now().timestamp())}"
        config = {
            'name': config_info['name'],
            'description': config_info['description'],
            'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'field_states': field_states,
            'value_mapping': self.value_mapping
        }
        
        # 添加到配置列表
        self.configs[config_id] = config
        
        # 保存配置
        self.save_configs()
        
        QMessageBox.information(self, '成功', f'配置 "{config_info["name"]}" 已保存!')
    
    def load_config(self):
        """加载配置"""
        if not self.configs:
            QMessageBox.information(self, '提示', '没有保存的配置!')
            return
        
        if self.dataset is None:
            QMessageBox.warning(self, '警告', '请先加载数据集!')
            return
        
        # 显示配置选择对话框
        dialog = LoadConfigDialog(self.configs, self)
        if dialog.exec_() != QDialog.Accepted:
            return
        
        config_id = dialog.get_selected_config()
        if not config_id:
            return
        
        config = self.configs[config_id]
        
        # 应用配置
        self.apply_config(config)
        
        QMessageBox.information(self, '成功', f'配置 "{config["name"]}" 已加载!')
    
    def apply_config(self, config):
        """应用配置到当前状态"""
        # 设置值映射
        self.value_mapping = config.get('value_mapping', {})
        
        # 应用值映射到数据集
        for field_name, mapping in self.value_mapping.items():
            if field_name in self.dataset.columns:
                new_data = self.dataset[field_name].copy()
                for old_val, new_val in mapping.items():
                    new_data = new_data.replace(old_val, new_val)
                self.dataset[field_name] = new_data
        
        # 重新创建字段列表
        self.fields_list.clear()
        
        # 获取字段状态
        field_states = config.get('field_states', {})
        
        # 创建排序后的字段列表
        sorted_fields = []
        for field in self.available_columns:
            if field in field_states:
                sorted_fields.append((field, field_states[field]['order']))
            else:
                # 对于不在配置中的字段，放在末尾
                sorted_fields.append((field, 999))
        
        # 按顺序排序
        sorted_fields.sort(key=lambda x: x[1])
        
        # 按顺序添加字段
        for field_name, _ in sorted_fields:
            # 获取字段状态
            state = field_states.get(field_name, {})
            display_name = state.get('display_name', field_name)
            is_checked = state.get('is_checked', False)
            
            # 创建并添加字段项
            item, widget = self.add_field_item(field_name, display_name, is_checked)
            self.fields_list.addItem(item)
            self.fields_list.setItemWidget(item, widget)
    
    def manage_configs(self):
        """管理配置"""
        if not self.configs:
            QMessageBox.information(self, '提示', '没有保存的配置!')
            return
        
        # 创建配置列表
        config_names = [config['name'] for config in self.configs.values()]
        selected, ok = QInputDialog.getItem(
            self, '管理配置', '选择要删除的配置:', 
            config_names, 0, False
        )
        
        if ok and selected:
            # 找到配置ID
            config_id = None
            for cid, config in self.configs.items():
                if config['name'] == selected:
                    config_id = cid
                    break
            
            if config_id:
                # 确认删除
                reply = QMessageBox.question(
                    self, '确认删除', 
                    f'确定要删除配置 "{selected}" 吗?',
                    QMessageBox.Yes | QMessageBox.No, 
                    QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    # 删除配置
                    del self.configs[config_id]
                    self.save_configs()
                    QMessageBox.information(self, '成功', f'配置 "{selected}" 已删除!')

    def preview_data(self):
        """预览处理后的数据"""
        selected_fields, display_names = self.get_selected_fields()
        if not selected_fields:
            QMessageBox.warning(self, '警告', '请至少选择一个字段！')
            return
            
        # 创建预览数据
        preview_data = self.dataset[selected_fields].copy()
        # 重命名列
        preview_data.columns = display_names
        
        # 显示预览对话框
        dialog = PreviewDataDialog(preview_data, self)
        dialog.exec_()

    def create_field_widget(self, field_name, display_name=None, is_checked=False):
        """创建字段小部件"""
        widget = FieldListItem(field_name)
        if display_name:
            widget.display_name = display_name
            widget.checkbox.setText(display_name)
        widget.checkbox.setChecked(is_checked)
        return widget

    def add_field_item(self, field_name, display_name=None, is_checked=False):
        """添加字段到列表"""
        item = QListWidgetItem()
        widget = self.create_field_widget(field_name, display_name, is_checked)
        
        # 连接按钮信号
        widget.up_button.clicked.connect(lambda: self.move_item_up(self.fields_list.row(item)))
        widget.down_button.clicked.connect(lambda: self.move_item_down(self.fields_list.row(item)))
        widget.rename_button.clicked.connect(lambda: self.rename_field(self.fields_list.row(item)))
        widget.edit_values_button.clicked.connect(lambda: self.edit_field_values(self.fields_list.row(item)))
        
        item.setSizeHint(widget.sizeHint())
        return item, widget

    def edit_field_values(self, row):
        """编辑字段值"""
        item = self.fields_list.item(row)
        field_widget = self.fields_list.itemWidget(item)
        
        if field_widget and self.dataset is not None:
            field_name = field_widget.field_name
            # 获取字段的所有唯一值
            unique_values = self.dataset[field_name].unique()
            
            dialog = EditValuesDialog(field_name, unique_values, self)
            if dialog.exec_() == QDialog.Accepted:
                # 获取修改后的值映射
                value_mapping = dialog.get_edited_values()
                if value_mapping:
                    # 更新数据集中的值
                    new_data = self.dataset[field_name].copy()
                    for old_val, new_val in value_mapping.items():
                        new_data = new_data.replace(old_val, new_val)
                    self.dataset[field_name] = new_data
                    
                    # 保存值映射
                    if field_name not in self.value_mapping:
                        self.value_mapping[field_name] = {}
                    self.value_mapping[field_name].update(value_mapping)
                    
                    QMessageBox.information(self, '成功', f'已更新字段 "{field_name}" 的值')

    def load_dataset(self):
        file_name, _ = QFileDialog.getOpenFileName(self, '选择数据集文件', '', 
                                                 'CSV files (*.csv);;Excel files (*.xlsx *.xls)')
        if file_name:
            try:
                if file_name.endswith('.csv'):
                    self.dataset = pd.read_csv(file_name)
                else:
                    self.dataset = pd.read_excel(file_name)
                
                # 清空并更新可用字段列表
                self.fields_list.clear()
                self.available_columns = list(self.dataset.columns)
                self.field_mapping = {col: col for col in self.available_columns}
                
                # 为每个字段创建可选项
                for column in self.available_columns:
                    item, widget = self.add_field_item(column)
                    self.fields_list.addItem(item)
                    self.fields_list.setItemWidget(item, widget)
                
                self.export_button.setEnabled(True)
                self.preview_button.setEnabled(True)
                self.sort_button.setEnabled(True)
                QMessageBox.information(self, '成功', '数据集加载成功！')
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'加载数据集时出错：{str(e)}')
    
    def move_item_up(self, row):
        if row > 0:
            # 获取当前项和上一项
            current_item = self.fields_list.item(row)
            current_widget = self.fields_list.itemWidget(current_item)
            
            if current_widget:
                # 保存当前项的状态
                field_name = current_widget.field_name
                display_name = current_widget.display_name
                is_checked = current_widget.checkbox.isChecked()
                
                # 移除当前项
                self.fields_list.takeItem(row)
                
                # 在新位置创建并插入项
                new_item, new_widget = self.add_field_item(field_name, display_name, is_checked)
                self.fields_list.insertItem(row - 1, new_item)
                self.fields_list.setItemWidget(new_item, new_widget)
                self.fields_list.setCurrentRow(row - 1)
    
    def move_item_down(self, row):
        if row < self.fields_list.count() - 1:
            # 获取当前项
            current_item = self.fields_list.item(row)
            current_widget = self.fields_list.itemWidget(current_item)
            
            if current_widget:
                # 保存当前项的状态
                field_name = current_widget.field_name
                display_name = current_widget.display_name
                is_checked = current_widget.checkbox.isChecked()
                
                # 移除当前项
                self.fields_list.takeItem(row)
                
                # 在新位置创建并插入项
                new_item, new_widget = self.add_field_item(field_name, display_name, is_checked)
                self.fields_list.insertItem(row + 1, new_item)
                self.fields_list.setItemWidget(new_item, new_widget)
                self.fields_list.setCurrentRow(row + 1)
    
    def rename_field(self, row):
        item = self.fields_list.item(row)
        field_widget = self.fields_list.itemWidget(item)
        
        if field_widget:
            dialog = RenameDialog(field_widget.display_name, self)
            
            if dialog.exec_() == QDialog.Accepted:
                new_name = dialog.get_new_name()
                if new_name and new_name.strip():
                    # 更新显示名称和映射
                    field_widget.display_name = new_name
                    field_widget.checkbox.setText(new_name)
                    self.field_mapping[field_widget.field_name] = new_name
    
    def get_selected_fields(self):
        selected_fields = []
        field_names = []
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            field_widget = self.fields_list.itemWidget(item)
            if field_widget and field_widget.checkbox.isChecked():
                selected_fields.append(field_widget.field_name)
                field_names.append(field_widget.display_name)
        return selected_fields, field_names
    
    def export_to_excel(self):
        selected_fields, display_names = self.get_selected_fields()
        if not selected_fields:
            QMessageBox.warning(self, '警告', '请至少选择一个字段！')
            return
            
        file_name, _ = QFileDialog.getSaveFileName(self, '保存Excel文件', '', 
                                                 'Excel files (*.xlsx)')
        if file_name:
            try:
                if not file_name.endswith('.xlsx'):
                    file_name += '.xlsx'
                    
                # 导出选中的字段
                selected_data = self.dataset[selected_fields].copy()
                # 重命名列
                selected_data.columns = display_names
                selected_data.to_excel(file_name, index=False)
                QMessageBox.information(self, '成功', '数据导出成功！')
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'导出数据时出错：{str(e)}')

    def sort_rows(self):
        """根据行中有效元素数量排序（无效元素包括NaN和'not performed'）"""
        if self.dataset is None:
            QMessageBox.warning(self, '警告', '请先加载数据集!')
            return
            
        selected_fields, display_names = self.get_selected_fields()
        if not selected_fields:
            QMessageBox.warning(self, '警告', '请至少选择一个字段!')
            return
        
        # 显示排序设置对话框
        dialog = SortRowsDialog(self)
        if dialog.exec_() != QDialog.Accepted:
            return
        
        sort_settings = dialog.get_sort_settings()
        
        # 创建用于排序的数据（仅包含选中的字段）
        sort_data = self.dataset[selected_fields].copy()
        
        # 计算每行有效元素数量（非NaN且不等于'not performed'）
        valid_mask = sort_data.notna() & (sort_data != 'Not performed')
        valid_counts = valid_mask.sum(axis=1)
        
        # 根据设置确定排序顺序
        ascending = sort_settings['order'] == 'ascending'
        
        # 创建临时列存储有效元素数量
        self.dataset['__valid_count__'] = valid_counts
        
        # 使用临时列排序
        self.dataset = self.dataset.sort_values(
            by='__valid_count__',
            ascending=ascending
        )
        
        # 删除临时列
        self.dataset = self.dataset.drop(columns=['__valid_count__'])
        
        QMessageBox.information(
            self, 
            '成功', 
            f'数据已按行有效元素{"升序" if ascending else "降序"}排序！（无效元素包括NaN和\'not performed\'，仅考虑选中的{len(selected_fields)}个字段）'
        )

def main():
    app = QApplication(sys.argv)
    ex = FileInfoSystem()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main() 