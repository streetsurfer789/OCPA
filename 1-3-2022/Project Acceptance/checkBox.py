from PyQt6.QtWidgets import QMessageBox
 
def check_work_area(self,pump_checkB,pressure_checkB,gravity_checkB, planfile_Btn,planfile_entry, work_entry):

    pump_string = {'Pump', 'pump', 'pumpstation', 'PumpStation','Pumpstation'}
    pressure_string = {"Pressure", "Pressurized", "Pressurized Pipe", "Pipe", "pressure","pressurized"}
    wastewater_string = {"Gravity" , "gravity" , "wastewater" , "WasteWater"}

    if work_entry in pump_string:
        print("1")
        pump_checkB.setChecked(True)
        pressure_checkB.setChecked(False)
        gravity_checkB.setChecked(False)
        planfile_Btn.setEnabled(True)
        planfile_entry.setEnabled(True)
    elif work_entry in pressure_string:
        print("2")
        pump_checkB.setChecked(False)
        pressure_checkB.setChecked(True)
        gravity_checkB.setChecked(False)
        planfile_Btn.setEnabled(True)
        planfile_entry.setEnabled(True)
    elif work_entry in wastewater_string:
        print("3")
        pump_checkB.setChecked(False)
        gravity_checkB.setChecked(True)
        pressure_checkB.setChecked(False)
        planfile_Btn.setEnabled(True)
        planfile_entry.setEnabled(True)
    else:
        showError()
        planfile_Btn.setEnabled(False)
        planfile_entry.setEnabled(False)


def showError():
    msgBox = QMessageBox()
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()