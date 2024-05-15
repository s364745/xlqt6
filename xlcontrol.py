# OpenPyXl
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

wb = load_workbook('Retting.xlsx')
ws = wb.active


def subtasks_per_tasks(task_number):
    list_of_subtasks = organize_subtasks(list_subtasks())[task_number - 1]
    print('subtasks_per_task: ', list_of_subtasks)


# UNFINISHED
def generate_mistake_lists():
    list_of_subtasks = organize_subtasks(list_subtasks())
    new_task_number = 0
    for task in list_of_subtasks:
        for subtask in task:
            if subtask == 'a':
                new_task_number += 1
            full_task_name = f'{new_task_number}{subtask}'
            print(full_task_name)


# Returns tasks numbers in a list
def find_task_numbers():
    rng = ws.iter_cols(min_row=2, max_row=2, min_col=2, max_col=ws.max_column)
    tasks = []
    for i in rng:
        for j in i:
            selected_cell = j
            if j.value is not None:
                tasks.append(j.value)

    return tasks

def total_student_number():
    rng = ws.iter_cols(min_row=4, max_row=ws.max_row, min_col=1, max_col=1)
    student_nbr = 0
    for i in rng:
        for j in i:
            selected_cell = j
            if j.value is not None:
                student_nbr+=1
    return student_nbr

def candidate_nbr(student_index):
    row = student_index+4
    return ws["A"+str(row)].value


def list_subtasks():
    rng = ws.iter_cols(min_row=3, max_row=3, min_col=2, max_col=ws.max_column)
    tasks = []
    for i in rng:
        for j in i:
            tasks.append(j.value)

    tasks = list(filter(lambda x: x is not None, tasks))

    return tasks


def organize_subtasks(subtasks):
    result = []
    sublist = []
    for element in subtasks:
        if element == 'a':
            if sublist:
                result.append(sublist)
            sublist = []
        sublist.append(element)
    result.append(sublist)
    return result


# Return nearby cell (left, right, up, down)
def nearby_cell(selected_cell, direction):
    if direction == 'left':
        next_cell = ws.cell(row=selected_cell.row, column=selected_cell.column - 1)
        return next_cell
    if direction == 'right':
        next_cell = ws.cell(row=selected_cell.row, column=selected_cell.column + 1)
        return next_cell

    if direction == 'up':
        next_cell = ws.cell(row=selected_cell.row - 1, column=selected_cell.column)
        return next_cell

    if direction == 'down':
        next_cell = ws.cell(row=selected_cell.row + 1, column=selected_cell.column)
        return next_cell


# Check
#print(organize_subtasks(list_subtasks())[0])
