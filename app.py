import sys

from PyQt6 import QtWidgets as qtw
from PyQt6 import QtCore as qtc
from PyQt6 import QtGui as qtq
from PyQt6 import uic
from PyQt6.QtWidgets import QVBoxLayout, QListWidgetItem, QTableWidgetItem, QWidget

# For statistics
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np

# Files
import xlcontrol as xl

# Import from ui file
gui, baseClass = uic.loadUiType('gui.ui')


# Argument can be baseClass
class MainWindow(baseClass):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Set up gui
        self.ui = gui()
        self.ui.setupUi(self)

        # Manage mistakes
        self.load_mistakes_table()
        self.load_mistake_headers()
        self.ui.add_mistake.clicked.connect(self.add_mistake)
        self.ui.rm_mistake.clicked.connect(self.rm_mistake)

        # Manage answer data
        self.total_student_number = xl.total_student_number()
        self.load_student_answers(0, 1)

        self.load_task_list()
        self.ui.pre_sudent.clicked.connect(lambda: self.load_student_answers(self.student - 1, self.task))
        self.ui.next_student.clicked.connect(lambda: self.load_student_answers(self.student + 1, self.task))

        # Statistics ----------------------------------------
        bar_layout = qtw.QVBoxLayout(self.ui.bar_widget)
        self.ui.bar_frame = QWidget()
        bar_layout.addWidget(self.ui.bar_frame)
        # Dummy data
        x = np.arange(8)
        y = np.random.randint(1, 10, size=8)
        # Figure for bar chart
        fig, ax = plt.subplots()
        ax.bar(x, y)
        # Matplotlib-canvas-widget
        self.ui.canvas = FigureCanvas(fig)
        bar_layout.addWidget(self.ui.canvas)
        # ----------------------------------------------------

        # Code stops here

    # Statistics
    def load_bar_diagram(self):
        pass

    def load_task_list(self):
        number = len(xl.find_task_numbers())
        for i in range(number):
            item = QListWidgetItem(f'Task {str(i + 1)}')
            self.ui.task_list.addItem(item)

        self.ui.task_list.itemClicked.connect(self.handle_item_clicked)

    def handle_item_clicked(self, item):
        task_trigger = item.text().split()[-1]  # get the task number
        self.load_student_answers(self.student, int(task_trigger))

    def update_progress_bar(self):
        ratio = int(100 * (self.student / self.total_student_number))
        self.ui.progressBar.setValue(ratio)

    def load_student_answers(self, student, task_number):
        if student >= 0 and student < self.total_student_number:
            self.student = student
        self.task = task_number

        # Load right amount of columns
        self.load_subtasks(self.task)

        # Display candidateNr
        print(xl.candidate_nbr(self.student))
        self.ui.candidate_edit.setText(str(xl.candidate_nbr(self.student)))

        all_answers = self.load_student_data(self.student)
        number_of_answers = xl.organize_subtasks(xl.list_subtasks())[self.task - 1]
        print('all_answers: ', all_answers)
        # print('To be answered here: ', number_of_answers)
        # print('number of answers: ', len(number_of_answers))

        counter = len(number_of_answers)

        # Load answers to column
        for i in range(counter):
            item = QTableWidgetItem(str(all_answers[i]))
            self.ui.answer_table.setItem(i, 1, item)
            print(all_answers[i])

        self.update_progress_bar()

    def load_subtasks(self, task_number):

        all_tasks = xl.organize_subtasks(xl.list_subtasks())
        index = task_number - 1
        selected_tasks = all_tasks[index]
        number_of_tasks = len(selected_tasks)

        selected_headers = ['Max Points:', 'Points:', 'Mistakes:']

        # Make rows of tasks
        self.ui.answer_table.setColumnCount(len(selected_headers))
        # |character|points|comment|
        self.ui.answer_table.setRowCount(number_of_tasks)

        self.ui.answer_table.setHorizontalHeaderLabels(selected_headers)
        self.ui.answer_table.setVerticalHeaderLabels(selected_tasks)

    # Mistake manager
    def load_mistake_headers(self):
        headers = ['Apply', 'Points lost', 'Explanation']
        self.ui.mistake_table.setHorizontalHeaderLabels(headers)

    def add_mistake(self):
        row_count = self.ui.mistake_table.rowCount()
        self.ui.mistake_table.insertRow(row_count)

        item = QTableWidgetItem()
        item.setFlags(item.flags() | qtc.Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(qtc.Qt.CheckState.Unchecked)
        self.ui.mistake_table.setItem(row_count, 0, item)

    def rm_mistake(self):
        selected_mistake = self.ui.mistake_table.currentRow()
        # NEEDS CONFIRMATION!
        if selected_mistake >= 0:
            self.ui.mistake_table.removeRow(selected_mistake)

    def load_mistakes_table(self):
        self.ui.mistake_table.setRowCount(1)
        item = QTableWidgetItem()
        item.setFlags(item.flags() | qtc.Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(qtc.Qt.CheckState.Unchecked)
        self.ui.mistake_table.setItem(0, 0, item)

    ### LIST ###
    # def load_mistake_list(self):
    #     mistake_list = ['Mistake 1:    - 2', 'Mistake 2:    - 4']
    #     for mistake in mistake_list:
    #         item = QListWidgetItem(mistake)
    #         item.setFlags(item.flags() | qtc.Qt.ItemFlag.ItemIsUserCheckable)
    #         item.setCheckState(qtc.Qt.CheckState.Unchecked)
    #         self.ui.lista.addItem(item)

    def load_student_data(self, student_nbr):
        row_nr = student_nbr + 4  # first 4 rows does'nt count
        candidate_values = []
        for cell in xl.ws[f'B{row_nr}':f'U{row_nr}'][0]:
            candidate_values.append(cell.value)
        # print(candidate_values)

        return candidate_values

    def add_row(self):
        num_rows = self.ui.answer_table.rowCount()
        self.ui.answer_table.setRowCount(num_rows + 1)

    def rm_row(self):
        num_rows = self.ui.answer_table.rowCount()
        self.ui.answer_table.setRowCount(num_rows - 1)

    def add_col(self):
        num_cols = self.ui.answer_table.columnCount()
        self.ui.answer_table.setColumnCount(num_cols + 1)

    def rm_col(self):
        num_cols = self.ui.answer_table.columnCount()
        self.ui.answer_table.setColumnCount(num_cols - 1)

    # Dummy Chart


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
