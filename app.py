import sys

from PyQt6 import QtWidgets as qtw
from PyQt6 import QtCore as qtc
from PyQt6 import QtGui as qtq
from PyQt6 import uic
from PyQt6.QtWidgets import QVBoxLayout, QListWidgetItem, QTableWidgetItem, QWidget, QMessageBox

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
        self.student=1
        self.isCharged=False

        # Manage mistakes
        self.load_mistake_headers()
        self.mistakes=xl.get_mistakes_list()
        self.ui.add_mistake.clicked.connect(self.add_mistake)
        self.ui.rm_mistake.clicked.connect(self.rm_mistake)

        # Manage answer data
        self.total_student_number = xl.total_student_number()
        self.load_student_answers(0, 1)

        self.load_task_list()
        self.ui.pre_sudent.clicked.connect(lambda: self.load_student_answers(self.student - 1, self.task))
        self.ui.next_student.clicked.connect(lambda: self.load_student_answers(self.student + 1, self.task))

        self.update_mistakes_tables()

        # Statistics ----------------------------------------
        bar_layout = qtw.QVBoxLayout(self.ui.bar_widget)
        self.ui.bar_frame = QWidget()
        bar_layout.addWidget(self.ui.bar_frame)
        # Loads bar chart
        x, y = self.load_chart_mistakes()
        # Figure for bar chart
        fig, ax = plt.subplots()
        ax.bar(x, y)
        # Labels
        plt.xlabel('Mistakes')
        plt.ylabel('Frequency')

        # Matplotlib-canvas-widget
        self.ui.canvas = FigureCanvas(fig)
        bar_layout.addWidget(self.ui.canvas)
        # ----------------------------------------------------

        # Code stops here

    # Statistics
    def load_chart_mistakes(self):

        # Dummy data, needs read data from excel
        x = ['mistake 1', 'mistake 2', 'mistake 3']
        y = [1, 2, 3]
        return x, y

    def load_task_list(self):
        number = len(xl.find_task_numbers())
        for i in range(number):
            item = QListWidgetItem(f'Task {str(i + 1)}')
            self.ui.task_list.addItem(item)
            if xl.mistakesnumber() < number:
                self.mistakes.append({ "mistakeID" : xl.add_mistakes(i+1), "task" : i+1, "index" : 0, "malus" : 0, "description" : "" }) # to add to the excel files
        self.ui.task_list.itemClicked.connect(self.handle_item_clicked)

    def handle_item_clicked(self, item):
        #print(self.ui.task_list.row(item))
        self.task = self.ui.task_list.row(item)+1
        task_trigger = item.text().split()[-1]  # get the task number
        self.load_student_answers(self.student, int(task_trigger))

    def update_progress_bar(self):
        ratio = int(100 * (self.student / self.total_student_number))
        self.ui.progressBar.setValue(ratio)

    def extract_subtask(self, subtask_names, all_answer, task_index):
        subtask_sizes = [len(subtask) for subtask in subtask_names]
        
        start_index = sum(subtask_sizes[:task_index])
        end_index = start_index + subtask_sizes[task_index]

        subtask_answers = all_answer[start_index:end_index]
        
        return subtask_answers

    def load_student_answers(self, student, task_number):
        if student >= 0 and student < self.total_student_number:
            self.student = student
        self.task = task_number

        # Load right amount of columns
        self.load_subtasks(self.task)

        # Display candidateNr
        self.ui.candidate_edit.setText(str(xl.candidate_nbr(self.student)))

        all_answers = self.load_student_data(self.student)
        subtask_names = xl.organize_subtasks(xl.list_subtasks())
        
        subtask_answers = self.extract_subtask(subtask_names, all_answers, self.task-1)

        # Load answers to column
        for i in range(len(subtask_answers)):
            item = QTableWidgetItem(str(subtask_answers[i]))
            self.ui.answer_table.setItem(i, 1, item)

        self.update_progress_bar()
        self.update_mistakes_tables()

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
        self.ui.mistake_table.cellChanged.connect(self.up_mistake)

    def up_mistake(self, row, column):
        if self.isCharged:
            if column in (0, 1, 2) :
                base_value = self.ui.mistake_table.item(row, column)
                modified_value = base_value.text()
                if modified_value != "" or base_value.checkState() != None:
                    for mistake in self.mistakes:
                        if mistake["index"] == row and mistake["task"] == self.task:
                            if column == 1:
                                mistake["malus"] = int(modified_value)
                                xl.up_mistakes(mistake["mistakeID"], mistake["task"], mistake["malus"], mistake["description"])
                            elif column == 2:
                                mistake["description"] = modified_value
                                xl.up_mistakes(mistake["mistakeID"], mistake["task"], mistake["malus"], mistake["description"])
                            if base_value.checkState() == qtc.Qt.CheckState.Unchecked:
                                xl.student_rem_mistakes(xl.candidate_nbr(self.student), mistake["mistakeID"])
                            if base_value.checkState() == qtc.Qt.CheckState.Checked:
                                xl.student_add_mistakes(xl.candidate_nbr(self.student), mistake["mistakeID"])
                print("Cell ({}, {}) modified w/ : {}".format(row, column, modified_value))

    def add_mistake(self):
        row_count = self.ui.mistake_table.rowCount()
        self.ui.mistake_table.insertRow(row_count)


        self.mistakes.append({ "mistakeID" : xl.add_mistakes(self.task), "task" : self.task, "index" : self.ui.mistake_table.rowCount()-1, "malus" : 0, "description" : "" }) # to add to the excel files

        item = QTableWidgetItem()
        item.setFlags(item.flags() | qtc.Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(qtc.Qt.CheckState.Unchecked)
        self.ui.mistake_table.setItem(row_count, 0, item)

    def rm_mistake(self):
        selected_mistake = self.ui.mistake_table.currentRow()

        # Should read mistake description and pass to message box. Not working for now

        # info_cell = self.mistake_table.item(selected_mistake, 1)
        # info_text = str(info_cell.text())
        #index = selected_mistake et task = self.task

        # Warning
        if selected_mistake >= 0:
            status = self.rm_mistake_warning()

            mistake_id=None

            for mistake in self.mistakes:
                if mistake["index"] == selected_mistake and mistake["task"] == self.task:
                    mistake_id=mistake["mistakeID"]
                    self.mistakes.remove(mistake)

            if mistake_id != None:
                xl.del_mistakes(mistake_id)

            if status == 0:
                pass
            elif status == 1:
                # if selected_mistake >= 0:
                self.ui.mistake_table.removeRow(selected_mistake)

    def rm_mistake_warning(self):
        button = QMessageBox.warning(
            self,
            'Warning!',
            'Are you sure you want to delete {description}? This action cannot be undone',
            buttons=QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            defaultButton=QMessageBox.StandardButton.No
        )
        if button == QMessageBox.StandardButton.No:
            return 0
        elif button == QMessageBox.StandardButton.Yes:
            return 1

    def update_mistakes_tables(self):
        self.isCharged=False
        self.ui.mistake_table.clearContents()
        self.ui.mistake_table.setRowCount(0)
        count=0
        mistake_of_this_student = xl.student_get_mistakes(xl.candidate_nbr(self.student))
        for i, mistake in enumerate(self.mistakes):
            if mistake["task"] == self.task:
                row_count = self.ui.mistake_table.rowCount()
                self.ui.mistake_table.insertRow(row_count)

                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(checkbox_item.flags() | qtc.Qt.ItemFlag.ItemIsUserCheckable)
                #checkbox_item.stateChanged.connect(lambda state, m_id=mistake["mistakeID"]: handle_checkbox_state_changed(state, m_id))
                if mistake["mistakeID"] in mistake_of_this_student:
                    checkbox_item.setCheckState(qtc.Qt.CheckState.Checked)
                else:
                    checkbox_item.setCheckState(qtc.Qt.CheckState.Unchecked)
                if type(mistake["malus"]) is not int :
                    mistake["malus"] = 0
                self.ui.mistake_table.setItem(row_count, 0, checkbox_item)
                txt_item = QTableWidgetItem()
                txt_item.setText(str(mistake["malus"]))
                self.ui.mistake_table.setItem(row_count, 1, txt_item)
                desc_item = QTableWidgetItem()
                desc_item.setText(mistake["description"])
                self.ui.mistake_table.setItem(row_count, 2, desc_item)
                mistake["index"] = row_count
                count+=1

        self.ui.mistake_table.setRowCount(count)
        self.isCharged=True

    def load_student_data(self, student_nbr):
        row_nr = student_nbr + 4  # first 4 rows does'nt count
        candidate_values = []
        for cell in xl.ws[f'B{row_nr}':f'U{row_nr}'][0]:
            candidate_values.append(cell.value)

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
