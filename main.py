import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLineEdit, QFileDialog, QLabel, QComboBox,
                             QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
                             QStackedWidget, QMessageBox, QProgressBar, QCheckBox)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6 import QtGui
from excel_handler import ExcelHandler
from compare_logic import CompareLogic

class CompareThread(QThread):
    finished = Signal(object, object)
    error = Signal(str)
    
    def __init__(self, left_file, left_sheet, right_file, right_sheet, pk_cols, sub_pk_cols, compare_cols, compare_by_row):
        super().__init__()
        self.left_file = left_file
        self.left_sheet = left_sheet
        self.right_file = right_file
        self.right_sheet = right_sheet
        self.pk_cols = pk_cols
        self.sub_pk_cols = sub_pk_cols
        self.compare_cols = compare_cols
        self.compare_by_row = compare_by_row
    
    def run(self):
        try:
            left_df = ExcelHandler.read_sheet(self.left_file, self.left_sheet)
            right_df = ExcelHandler.read_sheet(self.right_file, self.right_sheet)
            result_df, diff_cols = CompareLogic.compare_dfs(left_df, right_df, self.pk_cols, self.sub_pk_cols, self.compare_cols, compare_by_row=self.compare_by_row)
            self.finished.emit(result_df, diff_cols)
        except Exception as e:
            self.error.emit(str(e))

class ExcelCompareApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel比对合并工具")
        self.setGeometry(100, 100, 1000, 700)
        
        self.left_file = ""
        self.right_file = ""
        self.left_sheets = []
        self.right_sheets = []
        self.left_df = None
        self.right_df = None
        self.result_df = None
        self.full_result_df = None
        self.diff_cols = []
        self.compare_thread = None
        
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        
        self.page1 = QWidget()
        self.init_page1()
        self.page2 = QWidget()
        self.init_page2()
        self.page3 = QWidget()
        self.init_page3()
        
        self.stack.addWidget(self.page1)
        self.stack.addWidget(self.page2)
        self.stack.addWidget(self.page3)
    
    def init_page1(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(50, 50, 50, 50)
        
        title = QLabel("第一步：选择Excel文件")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)
        
        left_layout = QHBoxLayout()
        self.left_input = QLineEdit()
        self.left_input.setPlaceholderText("选择左表Excel文件（可拖拽）")
        self.left_input.setAcceptDrops(True)
        self.left_input.dragEnterEvent = lambda e: e.acceptProposedAction() if e.mimeData().hasUrls() else None
        self.left_input.dropEvent = lambda e: self.drop_file(e, "left")
        left_btn = QPushButton("浏览")
        left_btn.clicked.connect(lambda: self.select_file("left"))
        left_layout.addWidget(self.left_input)
        left_layout.addWidget(left_btn)
        layout.addLayout(left_layout)
        
        self.left_sheet_combo = QComboBox()
        self.left_sheet_combo.setPlaceholderText("选择左表Sheet")
        self.left_sheet_combo.currentTextChanged.connect(self.load_left_columns)
        layout.addWidget(self.left_sheet_combo)
        
        right_layout = QHBoxLayout()
        self.right_input = QLineEdit()
        self.right_input.setPlaceholderText("选择右表Excel文件（可拖拽）")
        self.right_input.setAcceptDrops(True)
        self.right_input.dragEnterEvent = lambda e: e.acceptProposedAction() if e.mimeData().hasUrls() else None
        self.right_input.dropEvent = lambda e: self.drop_file(e, "right")
        right_btn = QPushButton("浏览")
        right_btn.clicked.connect(lambda: self.select_file("right"))
        right_layout.addWidget(self.right_input)
        right_layout.addWidget(right_btn)
        layout.addLayout(right_layout)
        
        self.right_sheet_combo = QComboBox()
        self.right_sheet_combo.setPlaceholderText("选择右表Sheet")
        self.right_sheet_combo.currentTextChanged.connect(self.load_right_columns)
        layout.addWidget(self.right_sheet_combo)
        
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("下一步")
        next_btn.clicked.connect(self.go_page2)
        next_btn.setMinimumHeight(40)
        btn_layout.addStretch()
        btn_layout.addWidget(next_btn)
        layout.addLayout(btn_layout)
        
        self.page1.setLayout(layout)
    
    def drop_file(self, event, side):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                try:
                    sheets = ExcelHandler.get_sheet_names(file_path)
                    if side == "left":
                        self.left_file = file_path
                        self.left_input.setText(file_path)
                        self.left_sheet_combo.clear()
                        self.left_sheet_combo.addItems(sheets)
                        self.left_sheets = sheets
                    else:
                        self.right_file = file_path
                        self.right_input.setText(file_path)
                        self.right_sheet_combo.clear()
                        self.right_sheet_combo.addItems(sheets)
                        self.right_sheets = sheets
                except Exception as e:
                    QMessageBox.critical(self, "错误", str(e))
            else:
                QMessageBox.warning(self, "提示", "请拖拽Excel格式文件")
    
    def select_file(self, side):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if not file_path:
            return
        
        try:
            sheets = ExcelHandler.get_sheet_names(file_path)
            if side == "left":
                self.left_file = file_path
                self.left_input.setText(file_path)
                self.left_sheet_combo.clear()
                self.left_sheet_combo.addItems(sheets)
                self.left_sheets = sheets
            else:
                self.right_file = file_path
                self.right_input.setText(file_path)
                self.right_sheet_combo.clear()
                self.right_sheet_combo.addItems(sheets)
                self.right_sheets = sheets
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
    
    def load_left_columns(self, sheet_name):
        if not sheet_name or not self.left_file:
            return
        try:
            self.left_df = ExcelHandler.read_sheet(self.left_file, sheet_name)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
    
    def load_right_columns(self, sheet_name):
        if not sheet_name or not self.right_file:
            return
        try:
            self.right_df = ExcelHandler.read_sheet(self.right_file, sheet_name)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
    
    def init_page2(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(50, 50, 50, 50)
        
        title = QLabel("第二步：配置比对规则")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)
        
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("比对模式:"))
        self.mode_pk = QCheckBox("按匹配主键对齐")
        self.mode_pk.setChecked(True)
        self.mode_row = QCheckBox("按行号顺序比对")
        self.mode_pk.toggled.connect(self.toggle_mode)
        self.mode_row.toggled.connect(self.toggle_mode)
        mode_layout.addWidget(self.mode_pk)
        mode_layout.addWidget(self.mode_row)
        layout.addLayout(mode_layout)
        
        pk_container = QHBoxLayout()
        self.pk_layout = QVBoxLayout()
        self.pk_layout.addWidget(QLabel("主匹配主键（可多选）:"))
        self.pk_list = QListWidget()
        self.pk_list.setSelectionMode(QListWidget.MultiSelection)
        self.pk_list.setMaximumHeight(150)
        self.pk_layout.addWidget(self.pk_list)
        pk_container.addLayout(self.pk_layout)
        
        self.sub_pk_layout = QVBoxLayout()
        self.sub_pk_layout.addWidget(QLabel("辅助匹配主键（可选）:"))
        self.sub_pk_list = QListWidget()
        self.sub_pk_list.setSelectionMode(QListWidget.MultiSelection)
        self.sub_pk_list.setMaximumHeight(150)
        self.sub_pk_layout.addWidget(self.sub_pk_list)
        pk_container.addLayout(self.sub_pk_layout)
        layout.addLayout(pk_container)
        
        cols_layout = QHBoxLayout()
        compare_col_layout = QVBoxLayout()
        compare_col_layout.addWidget(QLabel("选择要比对的字段:"))
        self.compare_list = QListWidget()
        self.compare_list.setSelectionMode(QListWidget.MultiSelection)
        compare_col_layout.addWidget(self.compare_list)
        cols_layout.addLayout(compare_col_layout)
        layout.addLayout(cols_layout)
        
        btn_layout = QHBoxLayout()
        prev_btn = QPushButton("上一步")
        prev_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        compare_btn = QPushButton("开始比对")
        compare_btn.clicked.connect(self.start_compare)
        compare_btn.setMinimumHeight(40)
        btn_layout.addWidget(prev_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(compare_btn)
        layout.addLayout(btn_layout)
        
        self.page2.setLayout(layout)
    
    def toggle_mode(self):
        sender = self.sender()
        if sender == self.mode_pk and sender.isChecked():
            self.mode_row.setChecked(False)
            self.pk_layout.setEnabled(True)
            self.sub_pk_layout.setEnabled(True)
        elif sender == self.mode_row and sender.isChecked():
            self.mode_pk.setChecked(False)
            self.pk_layout.setEnabled(False)
            self.sub_pk_layout.setEnabled(False)
    
    def go_page2(self):
        if not self.left_file or not self.right_file:
            QMessageBox.warning(self, "提示", "请先选择两个Excel文件")
            return
        left_sheet = self.left_sheet_combo.currentText()
        right_sheet = self.right_sheet_combo.currentText()
        if not left_sheet or not right_sheet:
            QMessageBox.warning(self, "提示", "请选择两个文件对应的Sheet")
            return
        if self.left_df is None or self.left_df.empty:
            QMessageBox.warning(self, "提示", "左表Sheet为空，请重新选择")
            return
        if self.right_df is None:
            reply = QMessageBox.question(self, "右表为空", "右表Sheet为空，是否继续比对？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
        
        common_cols = list(set(self.left_df.columns) & set(self.right_df.columns)) if not self.right_df.empty else []
        if not common_cols and not self.right_df.empty:
            QMessageBox.warning(self, "提示", "两个Sheet没有共同的列，无法比对")
            return
        
        self.pk_list.clear()
        self.sub_pk_list.clear()
        self.compare_list.clear()
        output_cols = self.left_df.columns.tolist()
        for col in common_cols:
            pk_item = QListWidgetItem(col)
            self.pk_list.addItem(pk_item)
            sub_pk_item = QListWidgetItem(col)
            self.sub_pk_list.addItem(sub_pk_item)
        for col in output_cols:
            compare_item = QListWidgetItem(col)
            self.compare_list.addItem(compare_item)
        
        self.stack.setCurrentIndex(1)
    
    def start_compare(self):
        if self.compare_thread and self.compare_thread.isRunning():
            QMessageBox.warning(self, "提示", "正在比对中，请等待当前任务完成")
            return
        
        compare_by_row = self.mode_row.isChecked()
        pk_cols = []
        sub_pk_cols = []
        
        if not compare_by_row:
            pk_cols = [item.text() for item in self.pk_list.selectedItems()]
            sub_pk_cols = [item.text() for item in self.sub_pk_list.selectedItems()]
            if not pk_cols:
                QMessageBox.warning(self, "提示", "请至少选择一个主匹配主键")
                return
            
            # 主主键+辅助主键共同查重
            all_pk_cols = list(dict.fromkeys(pk_cols + sub_pk_cols))
            left_duplicates = CompareLogic.check_duplicate_pk(self.left_df, all_pk_cols)
            right_duplicates = CompareLogic.check_duplicate_pk(self.right_df, all_pk_cols)
            
            warn_msg = ""
            if left_duplicates:
                warn_msg += f"左表存在{len(left_duplicates)}组重复主键值\n"
            if right_duplicates:
                warn_msg += f"右表存在{len(right_duplicates)}组重复主键值\n"
            
            if warn_msg:
                warn_msg += "\n重复主键的行将按顺序配对比对，是否继续？"
                reply = QMessageBox.question(self, "主键重复提示", warn_msg, QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    return
        
        compare_cols = [item.text() for item in self.compare_list.selectedItems()]
        if not compare_cols:
            QMessageBox.warning(self, "提示", "请至少选择一个要比对的字段")
            return
        
        # 清空所有旧的结果缓存
        self.result_df = None
        self.full_result_df = None
        self.diff_cols = []
        # 清空表格和搜索框
        self.result_table.setRowCount(0)
        self.result_table.setColumnCount(0)
        self.search_input.clear()
        
        self.compare_thread = CompareThread(
            self.left_file, self.left_sheet_combo.currentText(),
            self.right_file, self.right_sheet_combo.currentText(),
            pk_cols, sub_pk_cols, compare_cols, compare_by_row
        )
        self.compare_thread.finished.connect(self.compare_finished)
        self.compare_thread.error.connect(self.compare_error)
        self.compare_thread.finished.connect(self.clear_thread)
        self.compare_thread.error.connect(self.clear_thread)
        self.compare_thread.start()
        
        QMessageBox.information(self, "提示", "正在比对，请稍候...")
    
    def clear_thread(self):
        self.compare_thread = None
    
    def compare_finished(self, result_df, diff_cols):
        self.result_df = result_df
        self.full_result_df = result_df.copy()
        self.diff_cols = diff_cols
        
        # 更新统计数据
        total = len(result_df)
        added = len(result_df[result_df['__status'] == 'added'])
        modified = len(result_df[result_df['__status'] == 'modified'])
        deleted = len(result_df[result_df['__status'] == 'deleted'])
        
        self.stat_total.setText(f"总记录数：{total}")
        self.stat_added.setText(f"新增：{added}")
        self.stat_modified.setText(f"修改：{modified}")
        self.stat_deleted.setText(f"删除：{deleted}")
        
        self.load_result_table()
        self.stack.setCurrentIndex(2)
    
    def compare_error(self, err_msg):
        QMessageBox.critical(self, "比对失败", err_msg)
    
    def init_page3(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(50, 50, 50, 50)
        
        title = QLabel("第三步：查看结果并导出")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)
        
        # 统计面板
        stat_layout = QHBoxLayout()
        self.stat_total = QLabel("总记录数：0")
        self.stat_added = QLabel("新增：0")
        self.stat_added.setStyleSheet("color: green; font-weight: bold;")
        self.stat_modified = QLabel("修改：0")
        self.stat_modified.setStyleSheet("color: #cc8800; font-weight: bold;")
        self.stat_deleted = QLabel("删除：0")
        self.stat_deleted.setStyleSheet("color: red; font-weight: bold;")
        
        stat_layout.addWidget(self.stat_total)
        stat_layout.addWidget(self.stat_added)
        stat_layout.addWidget(self.stat_modified)
        stat_layout.addWidget(self.stat_deleted)
        stat_layout.addStretch()
        layout.addLayout(stat_layout)
        
        # 筛选栏
        filter_layout = QHBoxLayout()
        self.filter_all = QPushButton("全部")
        self.filter_added = QPushButton("仅新增")
        self.filter_modified = QPushButton("仅修改")
        self.filter_deleted = QPushButton("仅删除")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索内容...")
        self.search_btn = QPushButton("搜索")
        
        filter_layout.addWidget(self.filter_all)
        filter_layout.addWidget(self.filter_added)
        filter_layout.addWidget(self.filter_modified)
        filter_layout.addWidget(self.filter_deleted)
        filter_layout.addStretch()
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.search_btn)
        layout.addLayout(filter_layout)
        
        self.result_table = QTableWidget()
        layout.addWidget(self.result_table)
        
        self.filter_all.clicked.connect(lambda: self.filter_result('all'))
        self.filter_added.clicked.connect(lambda: self.filter_result('added'))
        self.filter_modified.clicked.connect(lambda: self.filter_result('modified'))
        self.filter_deleted.clicked.connect(lambda: self.filter_result('deleted'))
        self.search_btn.clicked.connect(self.search_result)
        
        btn_layout = QHBoxLayout()
        prev_btn = QPushButton("上一步")
        prev_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        export_btn = QPushButton("导出合并Excel")
        export_btn.clicked.connect(self.export_result)
        export_btn.setMinimumHeight(40)
        btn_layout.addWidget(prev_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(export_btn)
        layout.addLayout(btn_layout)
        
        self.page3.setLayout(layout)
    
    def update_stat(self):
        if self.result_df is None or self.result_df.empty:
            self.stat_total.setText("总记录数：0")
            self.stat_added.setText("新增：0")
            self.stat_modified.setText("修改：0")
            self.stat_deleted.setText("删除：0")
            return
        total = len(self.result_df)
        added = len(self.result_df[self.result_df['__status'] == 'added'])
        modified = len(self.result_df[self.result_df['__status'] == 'modified'])
        deleted = len(self.result_df[self.result_df['__status'] == 'deleted'])
        
        self.stat_total.setText(f"总记录数：{total}")
        self.stat_added.setText(f"新增：{added}")
        self.stat_modified.setText(f"修改：{modified}")
        self.stat_deleted.setText(f"删除：{deleted}")
    
    def filter_result(self, filter_type):
        if self.full_result_df is None:
            return
        if filter_type == 'all':
            self.result_df = self.full_result_df.copy()
        else:
            self.result_df = self.full_result_df[self.full_result_df['__status'] == filter_type].copy()
        self.update_stat()
        self.load_result_table()
    
    def search_result(self):
        if self.full_result_df is None:
            return
        keyword = self.search_input.text().strip()
        if not keyword:
            self.result_df = self.full_result_df.copy()
        else:
            mask = self.full_result_df.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)
            self.result_df = self.full_result_df[mask].copy()
        self.update_stat()
        self.load_result_table()
    
    def load_result_table(self):
        self.result_table.setSortingEnabled(False)
        # 彻底清空表格所有内容
        self.result_table.clear()
        self.result_table.setRowCount(0)
        self.result_table.setColumnCount(0)
        
        if self.result_df is None or self.result_df.empty:
            self.update_stat()
            QMessageBox.information(self, "提示", "没有匹配的结果")
            return
        
        display_cols = [col for col in self.result_df.columns if col not in ['__status', '__diff_info']]
        self.result_table.setRowCount(len(self.result_df))
        self.result_table.setColumnCount(len(display_cols))
        self.result_table.setHorizontalHeaderLabels(display_cols)
        self.result_table.setSortingEnabled(True)
        
        green = QtGui.QColor(0, 255, 0)
        yellow = QtGui.QColor(255, 255, 0)
        red = QtGui.QColor(255, 0, 0)
        diff_red = QtGui.QColor(255, 0, 0)
        
        for i in range(len(self.result_df)):
            row_status = self.result_df.iloc[i]['__status']
            diff_info = self.result_df.iloc[i]['__diff_info']
            bg_color = None
            if row_status == 'added':
                bg_color = green
            elif row_status == 'modified':
                bg_color = yellow
            elif row_status == 'deleted':
                bg_color = red
            
            for j, col in enumerate(display_cols):
                val = str(self.result_df.iloc[i][col])
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if bg_color:
                    item.setBackground(bg_color)
                if col in diff_info:
                    item.setBackground(diff_red)
                self.result_table.setItem(i, j, item)
        
        self.result_table.resizeColumnsToContents()
    
    def export_result(self):
        if self.result_df is None or self.result_df.empty:
            QMessageBox.warning(self, "提示", "没有比对结果可以导出")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(self, "保存导出文件", "", "Excel文件 (*.xlsx)")
        if not output_path:
            return
        
        if os.path.exists(output_path):
            reply = QMessageBox.question(self, "文件已存在", "文件已存在，是否覆盖？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
        
        try:
            ExcelHandler.export_compare_result(self.result_df, output_path, self.diff_cols)
            QMessageBox.information(self, "成功", f"导出成功！文件已保存到:\n{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelCompareApp()
    window.show()
    sys.exit(app.exec())