# excel_pdf_gui.py
# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QMessageBox, QHBoxLayout, QFileDialog, QLineEdit, QFormLayout
)
from PyQt5.QtCore import Qt
import os
from excel_pdf_matcher import compare_excel_pdf

class CompareApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel & PDF PartNo比对工具")
        self.setAcceptDrops(True)
        self.resize(520, 320)

        self.excel_path = None
        self.pdf_path = None

        layout = QVBoxLayout()

        self.label = QLabel("请拖入 Excel 和 PDF 文件，或点击按钮选择文件")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.status_label = QLabel("Excel: [未选择]   |   PDF: [未选择]")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # ✅ 添加字段名称输入框
        # 添加一个水平布局：说明文字 + 输入框
        field_layout = QHBoxLayout()

        field_label = QLabel("对比列标题：")
        field_label.setFixedWidth(80)
        field_input = QLineEdit()
        field_input.setPlaceholderText("如：Part No")
        field_input.setText("Part No,NW(KG)")  # 默认值

        field_layout.addWidget(field_label)
        field_layout.addWidget(field_input)
        layout.addLayout(field_layout)

        self.field_input = field_input  # 保留引用，供后续调用使用

        # 文件选择按钮组
        btn_layout = QHBoxLayout()
        self.btn_select_excel = QPushButton("选择 Excel 文件")
        self.btn_select_excel.clicked.connect(self.select_excel_file)
        btn_layout.addWidget(self.btn_select_excel)

        self.btn_select_pdf = QPushButton("选择 PDF 文件")
        self.btn_select_pdf.clicked.connect(self.select_pdf_file)
        btn_layout.addWidget(self.btn_select_pdf)
        layout.addLayout(btn_layout)

        # 执行和清除按钮组
        action_layout = QHBoxLayout()
        self.btn_compare = QPushButton("开始比对")
        self.btn_compare.clicked.connect(self.handle_compare)
        self.btn_compare.setEnabled(False)
        action_layout.addWidget(self.btn_compare)

        self.btn_clear = QPushButton("清除文件")
        self.btn_clear.clicked.connect(self.reset_files)
        action_layout.addWidget(self.btn_clear)
        layout.addLayout(action_layout)

        self.setLayout(layout)

    def update_status(self):
        excel_name = os.path.basename(self.excel_path) if self.excel_path else "[未选择]"
        pdf_name = os.path.basename(self.pdf_path) if self.pdf_path else "[未选择]"
        self.status_label.setText(f"Excel: {excel_name}   |   PDF: {pdf_name}")
        self.btn_compare.setEnabled(bool(self.excel_path and self.pdf_path))

    def reset_files(self):
        self.excel_path = None
        self.pdf_path = None
        self.label.setText("请拖入 Excel 和 PDF 文件，或点击按钮选择文件")
        self.update_status()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        for f in files:
            if f.lower().endswith(('.xls', '.xlsx')):
                self.excel_path = f
            elif f.lower().endswith('.pdf'):
                self.pdf_path = f
        self.label.setText("✔ 文件已拖入")
        self.update_status()

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.excel_path = file_path
            self.label.setText("✔ 已选择 Excel 文件")
            self.update_status()

    def select_pdf_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.label.setText("✔ 已选择 PDF 文件")
            self.update_status()

    def handle_compare(self):
        field_name = self.field_input.text().strip()
        if not field_name:
            QMessageBox.warning(self, "字段名缺失", "请输入要匹配的字段名称")
            return
        try:
            output_path = compare_excel_pdf(self.excel_path, self.pdf_path, field_name)
            QMessageBox.information(self, "比对完成", f"结果文件已生成：\n{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生错误：\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CompareApp()
    window.show()
    sys.exit(app.exec_())
