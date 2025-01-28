import os
from checkpath import *
from PyQt6.QtWidgets import QMessageBox


def search_clicked(self,workOrder,development, cip, pump, pressure, gravity):
    print("made it to function")
    path = "O:\Field Services Division\Field Support Center\Project Acceptance"
    if workOrder != "" and len(workOrder) >= 5:
        for root, subdir, files in os.walk('O:\Field Services Division\Field Support Center\Project Acceptance'):
            for d in subdir:
                if d.find(workOrder) != -1:
                        print(d)
                        print("Im HERE")
                        walk = d
                        self.concat = path + "/" + walk
                        isdir = os.path.isdir(self.concat)
                        if isdir:
                            pump_Found, pressure_Found, gravity_Found = check_path(self.concat)
                            print(str(pump_Found) + " " + str(pressure_Found) + " " + str(gravity_Found))
                            development.setEnabled(True)
                            cip.setEnabled(True)
                            if pump_Found == 1:
                                pump.setChecked(True)
                            if pressure_Found == 1:
                                pressure.setChecked(True)
                            if gravity_Found == 1:
                                gravity.setChecked(True)
                            
                            #folders_found(pump_Found, pressure_Found, gravity_Found, concat, left_frame, right_frame, work_entry, workOrder.get(), top_right_frame)
                             
                        else:
                            pass
            break
    else:
        showError()


def showError():
    msgBox = QMessageBox()
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()
