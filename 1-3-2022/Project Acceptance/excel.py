from PyQt6 import QtGui
from PyQt6 import QtCore
from PyQt6 import QtWidgets

import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, QTimer, Qt, pyqtSlot
from PyQt6.QtWidgets import QComboBox, QItemDelegate, QListWidget, QListWidgetItem, QMessageBox, QWidget
import configparser
import json


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
                value = self._data.iloc[index.row(), index.column()]
                return str(value)

        if index.isValid():
            if role == Qt.ItemDataRole.BackgroundRole:
                value = self._data.iloc[index.row(), index.column()]
                if (
                    (isinstance(value, str))
                    and value == "Accepted"
                ):
                    return QtGui.QColor(207,225,167)

                if (
                    (isinstance(value, str))
                    and value == "Rejected"
                ):
                    return QtGui.QColor(187,28,42)
                


    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable


class ComboBoxDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super(ComboBoxDelegate, self).__init__(parent)
        self.items = []

    def setItems(self, items):
        self.items = items

    def createEditor(self, widget, option, index):
        editor = QComboBox(widget)
        editor.addItems(self.items)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole)
        
        if value:
            editor.setCurrentText(str(value))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole )

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

    # def paint(self, painter, option, index):
    #     text = self.items[index.row()]
    #     option.text = text
    #     QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)



def init_table(self, table, concat, pump_checkB,gravity_checkB,pressure_checkB,cip_checkB,development_checkB):

    print("in excel.py")
    self.structure_delegate = ComboBoxDelegate()
    self.category_delegate = ComboBoxDelegate()
    self.approved_cctv = ComboBoxDelegate()
    self.vendor = ComboBoxDelegate()
    self.reviewer = ComboBoxDelegate()

    config = configparser.RawConfigParser()
    config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\delegates.properties")

    if pump_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Pump.csv"

    if pump_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Pump.csv"

    if gravity_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Gravity.csv"
        
    if gravity_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Gravity.csv"

    if pressure_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Pressurized.csv"

    if pressure_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Pressurized.csv"


    excel_filename = concat + "/Excel/" + filename
    print(excel_filename)
    df = pd.read_csv(excel_filename)
    if df.size == 0:
        return

    filtered_df = df.loc[:, ~df.columns.isin(["Notes", "Notes.1"])]

    model = PandasModel(df)
    table.resizeRowsToContents()
    table.resizeColumnToContents(0)
    table.resizeColumnToContents(1)
    table.resizeColumnToContents(2)
    table.resizeColumnToContents(3)
    table.resizeColumnToContents(4)
    table.setColumnWidth(5, 300)
    table.resizeColumnToContents(6)
    table.setColumnWidth(7, 300)
    table.setModel(model)
    self.delegate = ComboBoxDelegate()
    self.delegate.setItems(["Accepted", "Rejected"])

    if pump_checkB.isChecked() and cip_checkB.isChecked():
        pump_values = config.get("PUMP", "category")
        pump_value_list = json.loads(pump_values)
        self.category_delegate.setItems(pump_value_list)
        table.setItemDelegateForColumn(0, self.category_delegate)

    if pump_checkB.isChecked() and development_checkB.isChecked():
        pump_values = config.get("PUMP", "category")
        pump_value_list = json.loads(pump_values)
        self.category_delegate.setItems(pump_value_list)
        table.setItemDelegateForColumn(0, self.category_delegate)

    if gravity_checkB.isChecked() and cip_checkB.isChecked():
        gravity_structure = config.get("GRAVITY", "structure")
        gravity_structure_list = json.loads(gravity_structure)

        
        self.structure_delegate.setItems(gravity_structure_list)
        table.setItemDelegateForColumn(3, self.structure_delegate)
        
    if gravity_checkB.isChecked() and development_checkB.isChecked():
        approved_cctv = config.get("GRAVITY", "approved_cctv")
        approved_cctv_list = json.loads(approved_cctv)
        self.approved_cctv.setItems(approved_cctv_list)

        vendor_surveyor = config.get("GRAVITY", "oc_surveyor")
        vendor_surveyor_list = json.loads(vendor_surveyor)
        self.vendor.setItems(vendor_surveyor_list)

        reviewer = config.get("GRAVITY", "reviewer")
        reviewer_list = json.loads(reviewer)
        self.reviewer.setItems(reviewer_list)

        table.setItemDelegateForColumn(2, self.approved_cctv)
        table.setItemDelegateForColumn(3, self.delegate)
        table.setItemDelegateForColumn(5, self.reviewer)
        table.setItemDelegateForColumn(7, self.delegate)
        table.setItemDelegateForColumn(8, self.vendor)
        table.setItemDelegateForColumn(9, self.reviewer)

    if pressure_checkB.isChecked() and cip_checkB.isChecked():
        pressure_structure_vals = config.get("PRESSURE", "structure")
        pressure_structure_vals_list = json.loads(pressure_structure_vals)

        pressure_category_val = config.get("PRESSURE", "category")
        category_list = json.loads(pressure_category_val)

        self.structure_delegate.setItems(pressure_structure_vals_list)
        self.category_delegate.setItems(category_list)
        table.setItemDelegateForColumn(0, self.category_delegate)
        table.setItemDelegateForColumn(3, self.structure_delegate)

    if pressure_checkB.isChecked() and development_checkB.isChecked():
        pressure_structure_vals = config.get("PRESSURE", "structure")
        pressure_structure_vals_list = json.loads(pressure_structure_vals)

        pressure_category_val = config.get("PRESSURE", "category")
        category_list = json.loads(pressure_category_val)

        self.structure_delegate.setItems(pressure_structure_vals_list)
        self.category_delegate.setItems(category_list)
        table.setItemDelegateForColumn(0, self.category_delegate)
        table.setItemDelegateForColumn(3, self.structure_delegate)
    
    table.setWordWrap(True)
    table.show()

def showError():
    msgBox = QMessageBox()
    msgBox.setText("File Not Found")
    msgBox.setWindowTitle("Error")
    msgBox.exec()


