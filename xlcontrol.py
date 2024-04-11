# OpenPyXl
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

wb = load_workbook('Retting.xlsx')
ws = wb.active


# Returns tasks numbers in a list
def find_task_numbers():
    rng = ws['B2':'U2']
    tasks = []
    for i in rng:
        for j in i:
            selected_cell = j
            if j.value is not None:
                tasks.append(j.value)

    return tasks


def list_subtasks():
    rng = ws['B3':'U3']
    tasks = []
    for i in rng:
        for j in i:
            tasks.append(j.value)

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
deloppgaver = list_subtasks()
oppgaver = organize_subtasks(deloppgaver)
print(find_task_numbers())
print(oppgaver)
