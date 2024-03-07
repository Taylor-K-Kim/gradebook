# Programmer: Taylor Kim (taylor97@csu.fullerton.edu)
# Date: 09/25/2023
#
# Program Description: 
# Simple python Gradebook GUI using pyqt5 and pyqtgraph which lets
# the user to import/export csv file, add/delete student, edit student grades,
# sort list by ascending/descending order, search/filter student id, 
# compute final letter grade, view averages, and view average graph.
#
# Install Requirements: pyqt5 and pyqtgraph using pip install

import sys
import csv
import pyqtgraph as pg
from pyqtgraph import plot
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QMainWindow,
    QToolBar
)


class GraphWindow(QtWidgets.QWidget):
    def __init__(self, avg_list):
        super().__init__()
        graph_layout = QtWidgets.QVBoxLayout()
        window = pg.plot()
        self.setWindowTitle("Average Scores Graph")

        # set horizontal tick labels
        xLabel = ['HW1', 'HW2', 'HW3', 'Quiz1', 'Quiz2', 'Quiz3', 'Quiz4', 'Midterm', 'Final']
        xVal = list(range(1, len(xLabel)+1))
        ticks = []
        for i, item in enumerate(xLabel):
            ticks.append((xVal[i], item))
        ticks=[ticks]
        
        # Adding bar graph to the plot window
        self.bar_graph = pg.BarGraphItem(x = xVal, height = avg_list, width = 0.6, brush = 'orange')
        window.addItem(self.bar_graph)
        xTicks = window.getAxis('bottom')
        xTicks.setTicks(ticks)
        graph_layout.addWidget(window)

        self.setLayout(graph_layout)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Welcome Display
        self.splash_screen = QtWidgets.QLabel('<font color=Grey size=20><b> Welcome to Gradebook </b></font>')
        self.splash_screen.setAlignment(QtCore.Qt.AlignCenter)
        self.splash_screen.setWindowFlags(QtCore.Qt.SplashScreen | QtCore.Qt.FramelessWindowHint)
        self.splash_screen.resize(700, 600)
        self.splash_screen.show()
        QtCore.QTimer.singleShot(3000, self.splash_screen.destroy)

        # Main Window Display
        self.setWindowTitle("Gradebook")
        self.resize(700, 600)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)

        # Table Display
        self.gradebook = QtGui.QStandardItemModel(self)
        
        self.tableView = QtWidgets.QTableView(self)    
        self.tableView.setModel(self.gradebook)
        self.tableView.horizontalHeader().setStretchLastSection(True)
        self.tableView.setSortingEnabled(True)
        self.tableView.resizeColumnsToContents()

        # Search Bar Display
        self.search_bar = QtWidgets.QLineEdit()
        self.search_bar.setPlaceholderText("Search SID here...")
        self.search_bar.textChanged.connect(self.filter_search)

        self.proxyModel = QtCore.QSortFilterProxyModel(self.tableView)
        self.proxyModel.setFilterCaseSensitivity(QtCore.Qt.CaseSensitivity.CaseInsensitive)
        self.proxyModel.setSourceModel(self.gradebook)
        self.tableView.setModel(self.proxyModel)

        # Add search display and table display on main display
        MainLayout = QtWidgets.QVBoxLayout()
        MainLayout.addWidget(self.search_bar)
        MainLayout.addWidget(self.tableView)

        # Main Tool Bars
        toolbar = QToolBar("Main toolbar")
        self.addToolBar(toolbar)

        import_button = QAction("Import", self)
        import_button.triggered.connect(self.ImportButtonClick)
        toolbar.addAction(import_button)

        toolbar.addSeparator()

        export_button = QAction("Export", self)
        export_button.triggered.connect(self.ExportButtonClick)
        toolbar.addAction(export_button)

        toolbar.addSeparator()

        add_button = QAction("Add Student", self)
        add_button.triggered.connect(self.AddButtonClick)
        toolbar.addAction(add_button)

        toolbar.addSeparator()

        delete_button = QAction("Delete Student", self)
        delete_button.triggered.connect(self.DeleteButtonClick)
        toolbar.addAction(delete_button)

        toolbar.addSeparator()

        final_letter_button = QAction("Compute Final Grade", self)
        final_letter_button.triggered.connect(self.ComputeButtonClick)
        toolbar.addAction(final_letter_button)

        toolbar.addSeparator()

        view_avg_button = QAction("View Average", self)
        view_avg_button.triggered.connect(self.ViewAvgButtonClick)
        toolbar.addAction(view_avg_button)

        toolbar.addSeparator()

        view_graph_button = QAction("View Graph", self)
        view_graph_button.triggered.connect(self.ViewGraphButtonClick)
        toolbar.addAction(view_graph_button)

        toolbar.addSeparator()

        # Status Bar Display initalization
        self.status_bar = self.statusBar()

        # Main Menu Bar
        menu = self.menuBar()

        file_menu = menu.addMenu("&File")
        file_menu.addAction(import_button)
        file_menu.addAction(export_button)
        file_menu.addSeparator()

        # No actions added for edit/help menu
        edit_menu = menu.addMenu("&Edit")
        edit_menu.addSeparator()
        help_menu = menu.addMenu("&Help")
        help_menu.addSeparator()

        # Submenu
        #file_submenu = file_menu.addMenu("Save")
        #file_submenu.addAction(export_button)
        
        # Display all widgets
        central_widget.setLayout(MainLayout)
        

    def ImportButtonClick(self):
        # Importing new file will clear old data
        self.gradebook.clear()

        # Get file path
        filename = QtWidgets.QFileDialog.getOpenFileName()
        path = filename[0]

        # Display path on status bar
        self.status_bar.showMessage(path)
        header = True

        # Write file data to table
        if path != '':
            with open(path, "r") as fi:  
                for row in csv.reader(fi):
                    if header is True:
                        self.gradebook.setHorizontalHeaderLabels(row)
                        header = False
                    else:
                        items = [
                            QtGui.QStandardItem(field)
                            for field in row
                        ]
                        self.gradebook.appendRow(items)


    def ExportButtonClick(self):
        # Get file name to save as...
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Export/Save CSV', '', 'CSV(*.csv)')
        path = filename[0]

        if path != '':
            with open(path, 'w') as fx:
                cols = range(self.gradebook.columnCount())
                header = [self.gradebook.horizontalHeaderItem(col).text()
                        for col in cols]
                writer = csv.writer(fx, dialect='excel', lineterminator='\n')
                writer.writerow(header)
                for row in range(self.gradebook.rowCount()):
                    row_data = []
                    for column in range(self.gradebook.columnCount()):
                        item = self.gradebook.item(row, column)
                        if item is not None:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    writer.writerow(row_data)


    def AddButtonClick(self):
        self.gradebook.insertRows(self.gradebook.rowCount(), 1)


    def DeleteButtonClick(self):
        selected = sorted(set(index.row() for index in self.tableView.selectedIndexes()))
        #first_name = self.gradebook.item(selected[0],1).text()
        #last_name = self.gradebook.item(selected[0],2).text()
        #delete_msg = "Do you really want to delete "+first_name+" "+last_name+"?"
        delete_msg = "Do you really want to delete?"
        # Double check for deleting student info
        del_conf_msg = QtWidgets.QMessageBox.warning(self, "Delete Row",delete_msg , QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes)
        if del_conf_msg == QtWidgets.QMessageBox.Yes:
            self.gradebook.removeRow(selected[0])


    def ComputeButtonClick(self):
        col = self.gradebook.columnCount()
        final_col = self.gradebook.horizontalHeaderItem(col-1).text()

        # Add new column for final grade
        if final_col != "Final Grade":
            self.gradebook.insertColumns(col, 1)
            self.gradebook.setHeaderData(col, QtCore.Qt.Orientation.Horizontal ,"Final Grade")
            self.calc_letter_grade(col=col)
        else:
            # Update Final grade
            self.calc_letter_grade(col=col-1)
        

    def ViewAvgButtonClick(self):
        avg=[]
        for i in range(4, 13):
            avg.append(self.calculate_average(col=i))

        hw_msg = "HW1: {:0.2f}\nHW2: {:0.2f}\nHW3:  {:0.2f}".format(avg[0], avg[1], avg[2])
        qz_msg = "\nQuiz1:  {:0.2f}\nQuiz2:  {:0.2f}\nQuiz3:  {:0.2f}\nQuiz4: {:0.2f}".format(avg[3], avg[4], avg[5], avg[6])
        ex_msg = "\nMidterm:  {:0.2f}\nFinal Exam: {:0.2f}".format(avg[7], avg[8])
        avg_msg = "Average Scores\n\n"+hw_msg+qz_msg+ex_msg
        # pop-up window open for average scores for all assignments and exams
        QtWidgets.QMessageBox.information(self, "Average Scores", avg_msg,QtWidgets.QMessageBox.Ok)
    
        
    def get_avg_list(self):
        # getting current average list for GraphWindow class
        avg_list = []
        for i in range(4, 13):
            avg_list.append(self.calculate_average(col=i))
        return avg_list


    def ViewGraphButtonClick(self):
        avg_list = self.get_avg_list()
        self.graph_window = GraphWindow(avg_list)
        self.ToggleGraphWindow(self.graph_window)


    def ToggleGraphWindow(self, gwindow):
        # toggles graph window
        if gwindow.isVisible():
            gwindow.hide()
        else:
            gwindow.show()


    def filter_search(self, text):
        self.proxyModel.setFilterFixedString(text)


    def calc_letter_grade(self, col):
        hw_data, quiz_data, midterm_data, final_exam_data = ([] for _ in range(4))

        for row in range(self.gradebook.rowCount()):
            # HW grades
            for c in range(4, 7):
                if self.gradebook.item(row,c) != None:
                    hw_data.append(int(self.gradebook.item(row,c).text()))
                else:
                     msg = "Error: Empty cell in row-"+str(row+1)+" HW"+str(c-3)
                     msg_click = QtWidgets.QMessageBox.information(self, "Empty cell", msg ,QtWidgets.QMessageBox.Ok)
                     if msg_click == QtWidgets.QMessageBox.Ok:
                         return
            # Quiz grades
            for c in range(7, 11):
                if self.gradebook.item(row,c) != None:
                    quiz_data.append(int(self.gradebook.item(row,c).text()))
                else:
                    msg = "Error: Empty cell in row-"+str(row+1)+" Quiz"+str(c-6)
                    msg_click = QtWidgets.QMessageBox.information(self, "Empty cell", msg ,QtWidgets.QMessageBox.Ok)
                    if msg_click == QtWidgets.QMessageBox.Ok:
                        return
            # Midterm Exam grade
            if self.gradebook.item(row,11) != None:
                midterm_data.append(int(self.gradebook.item(row,11).text()))
            else:
                msg = "Error: Empty cell in row-"+str(row+1)+" MidtermExam"
                msg_click = QtWidgets.QMessageBox.information(self, "Empty cell", msg ,QtWidgets.QMessageBox.Ok)
                if msg_click == QtWidgets.QMessageBox.Ok:
                    return
            # Final Exam grade
            if self.gradebook.item(row,12) != None:
                final_exam_data.append(int(self.gradebook.item(row, 12).text()))
            else:
                msg = "Error: Empty cell in row-"+str(row+1)+" FinalExam"
                msg_click = QtWidgets.QMessageBox.information(self, "Empty cell", msg ,QtWidgets.QMessageBox.Ok)
                if msg_click == QtWidgets.QMessageBox.Ok:
                    return
            
            # Calculate weigted grade
            hw_weighted = (sum(hw_data)/3)*0.2
            quiz_weighted = (sum(quiz_data)/4)*0.2
            midterm_weighted  = (midterm_data[0])*0.3
            final_exam_weighted = (final_exam_data[0])*0.3

            final_percentage = hw_weighted + quiz_weighted + midterm_weighted + final_exam_weighted
            
            # Calculate letter grade
            if final_percentage <= 100 and final_percentage > 90:
                self.gradebook.setData(self.gradebook.index(row, col), "A ({:0.2f}%)".format(final_percentage))
            elif final_percentage <= 90 and final_percentage > 80:
                self.gradebook.setData(self.gradebook.index(row, col), "B ({:0.2f}%)".format(final_percentage))
            elif final_percentage <= 80 and final_percentage > 70:
                self.gradebook.setData(self.gradebook.index(row, col), "C ({:0.2f}%)".format(final_percentage))
            elif final_percentage <= 70 and final_percentage > 60:
                self.gradebook.setData(self.gradebook.index(row, col), "D ({:0.2f}%)".format(final_percentage))
            elif final_percentage <= 60:
                self.gradebook.setData(self.gradebook.index(row, col), "F ({:0.2f}%)".format(final_percentage))
            
            hw_data.clear()
            quiz_data.clear()
            midterm_data.clear()
            final_exam_data.clear()


    def calculate_average(self, col):
        avg_list = []
        avg = 0
        total_rows = self.gradebook.rowCount()
        
        for row in range(total_rows):
            if self.gradebook.item(row,col) != None:
                avg_list.append(int(self.gradebook.item(row,col).text()))
            else:
                msg = "Error: Empty cell in row"+str(row+1)+" col"+str(col+1)+"\n\nScore set to 0"
                msg_click = QtWidgets.QMessageBox.information(self, "Empty cell", msg ,QtWidgets.QMessageBox.Ok)
                if msg_click == QtWidgets.QMessageBox.Ok:
                    avg_list.append(0)
                    
        sum_list = sum(avg_list)
        if total_rows != 0: # seg fault error handling when first window opened with empty data
            avg = sum_list / total_rows
    
        return avg
    

    def closeEvent(self, event):
        close_msg = QtWidgets.QMessageBox.question(self, "Quit", "Do you want to Quit?", QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes)
        if close_msg == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())