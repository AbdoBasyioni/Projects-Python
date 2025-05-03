
import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QSplashScreen, QStackedWidget, QDateEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QFormLayout
)
from PyQt5.QtCore import QDate, Qt, QTimer
from PyQt5.QtGui import QPixmap, QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

def create_excel_if_not_exists():
    if not os.path.exists('sales_data.xlsx'):
        columns = ['Date', 'Product', 'Customer', 'Qty', 'Total']
        pd.DataFrame(columns=columns).to_excel('sales_data.xlsx', index=False)

def create_inventory_if_not_exists():
    if not os.path.exists('inventory.xlsx'):
        columns = ['Product', 'Quantity', 'Price', 'Date']
        pd.DataFrame(columns=columns).to_excel('inventory.xlsx', index=False)

def save_data(date, product, customer, qty, total):
    df = pd.read_excel('sales_data.xlsx')
    new_row = {'Date': date, 'Product': product, 'Customer': customer, 'Qty': qty, 'Total': total}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel('sales_data.xlsx', index=False)
#------> صفحة المخزون
class AddProductDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("إضافة منتج جديد")
        layout = QFormLayout()
        self.setLayout(layout)

        self.name_input = QLineEdit()
        self.qty_input = QLineEdit()
        self.price_input = QLineEdit()
        self.date_input = QDateEdit()
        self.date_input.setDate(QDate.currentDate())

        layout.addRow("اسم المنتج:", self.name_input)
        layout.addRow("الكمية:", self.qty_input)
        layout.addRow("السعر:", self.price_input)
        layout.addRow("التاريخ:", self.date_input)

        save_btn = QPushButton("حفظ")
        save_btn.clicked.connect(self.save_product)
        layout.addRow(save_btn)

    def save_product(self):
        try:
            df = pd.read_excel('inventory.xlsx')
            new_row = {
                'Product': self.name_input.text(),
                'Quantity': int(self.qty_input.text()),
                'Price': float(self.price_input.text()),
                'Date': self.date_input.date().toString("yyyy-MM-dd")
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel('inventory.xlsx', index=False)
            QMessageBox.information(self, "تم", "تمت إضافة المنتج.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", str(e))

class InventoryPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.table = QTableWidget()
        layout.addWidget(self.table)

        btn_show = QPushButton("عرض المخزون")
        btn_add = QPushButton("إضافة منتج")
        btn_back = QPushButton("رجوع")

        btn_show.clicked.connect(self.load_data)
        btn_add.clicked.connect(self.add_product)
        btn_back.clicked.connect(lambda: self.main_window.setCurrentIndex(0))

        layout.addWidget(btn_show)
        layout.addWidget(btn_add)
        layout.addWidget(btn_back)

    def load_data(self):
        df = pd.read_excel('inventory.xlsx')
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def add_product(self):
        dialog = AddProductDialog(self)
        dialog.exec_()
        self.load_data()

class SalesForm(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.date_input = QDateEdit()
        self.date_input.setDate(QDate.currentDate())
        self.product_input = QLineEdit()
        self.customer_input = QLineEdit()
        self.qty_input = QLineEdit()
        self.total_input = QLineEdit()

        save_btn = QPushButton("حفظ")
        new_btn = QPushButton("تسجيل جديد")
        back_btn = QPushButton("رجوع")

        save_btn.clicked.connect(self.save)
        new_btn.clicked.connect(self.clear_fields)
        back_btn.clicked.connect(lambda: self.main_window.setCurrentIndex(0))

        layout.addLayout(self.create_row("التاريخ:", self.date_input))
        layout.addLayout(self.create_row("المنتج:", self.product_input))
        layout.addLayout(self.create_row("العميل:", self.customer_input))
        layout.addLayout(self.create_row("الكمية:", self.qty_input))
        layout.addLayout(self.create_row("الإجمالي:", self.total_input))
        layout.addWidget(save_btn)
        layout.addWidget(new_btn)
        layout.addWidget(back_btn)

    def create_row(self, label_text, widget):
        row = QHBoxLayout()
        row.addWidget(QLabel(label_text))
        row.addWidget(widget)
        return row

    def save(self):
        try:
            save_data(
                self.date_input.date().toString("yyyy-MM-dd"),
                self.product_input.text(),
                self.customer_input.text(),
                int(self.qty_input.text()),
                float(self.total_input.text())
            )
            QMessageBox.information(self, "تم", "تم حفظ البيانات.")
            self.clear_fields()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", str(e))

    def clear_fields(self):
        self.product_input.clear()
        self.customer_input.clear()
        self.qty_input.clear()
        self.total_input.clear()

class DataView(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.table = QTableWidget()
        layout.addWidget(self.table)

        refresh_btn = QPushButton("تحديث البيانات")
        back_btn = QPushButton("رجوع")
        refresh_btn.clicked.connect(self.load_data)
        back_btn.clicked.connect(lambda: self.main_window.setCurrentIndex(0))

        layout.addWidget(refresh_btn)
        layout.addWidget(back_btn)

        self.load_data()

    def load_data(self):
        df = pd.read_excel('sales_data.xlsx')
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        for i, row in df.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

class AnalysisView(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)

        analyze_btn = QPushButton("تحليل المبيعات")
        back_btn = QPushButton("رجوع")
        analyze_btn.clicked.connect(self.plot_sales)
        back_btn.clicked.connect(lambda: self.main_window.setCurrentIndex(0))

        layout.addWidget(analyze_btn)
        layout.addWidget(back_btn)

    def plot_sales(self):
        df = pd.read_excel('sales_data.xlsx')
        df['Date'] = pd.to_datetime(df['Date'])
        df['Month'] = df['Date'].dt.to_period('M')
        summary = df.groupby('Month')['Total'].sum()

        self.ax.clear()
        summary.plot(kind='bar', ax=self.ax, color='skyblue')
        self.ax.set_title("إجمالي المبيعات لكل شهر")
        self.ax.set_xlabel("الشهر")
        self.ax.set_ylabel("الإجمالي")
        self.canvas.draw()

class HomePage(QWidget):
    def __init__(self, stacked_widget):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("مرحبًا بك في نظام تتبع المبيعات"))
        btn_input = QPushButton("إدخال بيانات")
        btn_view = QPushButton("عرض البيانات")
        btn_analyze = QPushButton("تحليل البيانات")
        btn_inventory = QPushButton("إدارة المخزون")
        btn_exit = QPushButton("خروج")

        btn_input.clicked.connect(lambda: stacked_widget.setCurrentIndex(1))
        btn_view.clicked.connect(lambda: stacked_widget.setCurrentIndex(2))
        btn_analyze.clicked.connect(lambda: stacked_widget.setCurrentIndex(3))
        btn_inventory.clicked.connect(lambda: stacked_widget.setCurrentIndex(4))
        btn_exit.clicked.connect(sys.exit)

        layout.addWidget(btn_input)
        layout.addWidget(btn_view)
        layout.addWidget(btn_analyze)
        layout.addWidget(btn_inventory)
        layout.addWidget(btn_exit)

def show_splash(app):
    splash = QSplashScreen()
    splash.setPixmap(QPixmap())
    splash.setFont(QFont("Arial", 16))
    splash.showMessage("مرحبًا بك في نظام المبيعات\nجارٍ التحميل...", Qt.AlignCenter | Qt.AlignBottom, Qt.white)
    splash.setStyleSheet("background-color: #34495e; color: white;")
    splash.show()
    return splash

if __name__ == "__main__":
    create_excel_if_not_exists()
    create_inventory_if_not_exists()

    app = QApplication(sys.argv)
    splash = show_splash(app)

    stacked_widget = QStackedWidget()
    home = HomePage(stacked_widget)
    input_page = SalesForm(stacked_widget)
    view_page = DataView(stacked_widget)
    analyze_page = AnalysisView(stacked_widget)
    inventory_page = InventoryPage(stacked_widget)

    stacked_widget.addWidget(home)
    stacked_widget.addWidget(input_page)
    stacked_widget.addWidget(view_page)
    stacked_widget.addWidget(analyze_page)
    stacked_widget.addWidget(inventory_page)

    QTimer.singleShot(5000, splash.close)
    QTimer.singleShot(5000, stacked_widget.show)

    sys.exit(app.exec_())
