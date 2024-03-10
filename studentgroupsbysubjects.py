from tkinter import Tk
from tkinter import filedialog
import numpy
import pandas
import random
import ctypes
import os

APP_TITLE = "Student Groups By Subjects"
EXCEL_FILES = "Excel datoteke"
CHOOSE_EXCEL_FILE = "Izberi Excel datoteko"
SAVE_RESULT_IN_EXCEL_FILE = "Shrani rezultat v Excel datoteko"
GROUPS = "Skupine"

ERR_NOTIFY_DEVELOPER = "Prosimo, obvestite razvijalca o napaki!"
ERR_OPENING_EXCEL_FILE = "Napaka pri odpiranju Excel datoteke:"
ERR_PARSING_EXCEL_FILE = "Napaka pri branju podatkov iz Excel datoteke:"
ERR_SAVING_EXCEL_FILE = "Napaka pri shranjevanju Excel datoteke:"

def open_excel_file():
    try:
        filename = filedialog.askopenfilename(title=CHOOSE_EXCEL_FILE, filetypes=[(EXCEL_FILES, "*.xlsx")])
        return filename
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, ERR_OPENING_EXCEL_FILE + f"\n{e}\n\n{ERR_NOTIFY_DEVELOPER}", APP_TITLE, 0)
        exit()

def get_excel_data(filename):
    try:
        students_sheet = pandas.read_excel(filename, sheet_name=0)
        exclusions_sheet = pandas.read_excel(filename, sheet_name=1)
        return students_sheet, exclusions_sheet
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, ERR_PARSING_EXCEL_FILE +  f"\n{e}\n\n{ERR_NOTIFY_DEVELOPER}", APP_TITLE, 0)
        exit()

def create_groups(students_sheet, exclusions_sheet):
    subjects = {subject: students_sheet[subject].tolist() for subject in students_sheet.columns}
    exlusions = {group: exclusions_sheet[group].tolist() for group in exclusions_sheet.columns}

    for subject in subjects:
        subjects[subject] = [student for student in subjects[subject] if str(student) != "nan"]
        subjects[subject] = random.sample(subjects[subject], len(subjects[subject]))
    for exclusionGroup in exlusions:
        exlusions[exclusionGroup] = [exclusion for exclusion in exlusions[exclusionGroup] if str(exclusion) != "nan"]
        exlusions[exclusionGroup] = random.sample(exlusions[exclusionGroup], len(exlusions[exclusionGroup]))
    
    groups = []

    for exclusionGroup in exlusions:
        for exclusion in exlusions[exclusionGroup]:                             # Za vsakega izključenega študenta v skupini izključitev
            exclusion_subject = [subject for subject in subjects if exclusion in subjects[subject]][0]
            #print(f"Gledamo: študent {exclusion} ima predmet {exclusion_subject}")
            added = False

            for group in groups:                                                # Za vsako že narejeno skupino
                can_add = True

                for student in group:                                           # ALi sploh lahko dodamo študenta v to skupino
                    if (student in exlusions[exclusionGroup]):
                        can_add = False
                        #print(f"V skupini {group} je že študent {student}, glejmo naslednjo skupino!")
                        break

                if (can_add):                                                   # Če lahko dodamo študenta v skupino preveri še, če ta skupina že ima študenta z istim predmetom
                    for student in group:
                        if (student in subjects[exclusion_subject]):
                            can_add = False
                            #print(f"V skupini {group} je že študent {student}, ki ima predmet {exclusion_subject}, glejmo naslednjo skupino!")
                            break
                    if (can_add):
                        group.append(exclusion)
                        added = True
                        #print(f"Študent {exclusion} je bil dodan v skupino {group}.\n")
                        break
                
            if (added == False):
                groups.append([exclusion])
                #print(f"Študent {exclusion} je bil dodan v novo skupino!\n")

    #print(f"\nTrenutne skupine: {groups}\nVsi študenti:\n")
    for subject in subjects:                                                    # Dodaj študente, ki niso v nobeni skupini  (niso v nobeni izključitvi)
        for student in subjects[subject]:
            student_subject = [subject for subject in subjects if student in subjects[subject]][0]
            #print(f"Gledamo: študent {student} ima predmet {student_subject}")
            added = False

            for group in groups:
                if (student in group):
                    added = True
                    #print(f"Študent {student} je že v skupini {group}.\n")
                    break

            if (added == False):                                                # Če študent ni bil dodan v nobeno skupino ga dodaj v prvo prosto skupino
                for group in groups:
                    can_add = True

                    for student_in_group in group:
                        if (student_in_group in subjects[student_subject]):     # Skupina že ima študenta z istim predmetom
                            can_add = False
                            break

                    if (can_add):
                        group.append(student)
                        #print(f"Študent {exclusion} je bil dodan v skupino {group}.\n")
                        added = True
                        break
            
            if (added == False):
                groups.append([student])
                #print(f"Študent {student} je bil dodan v novo skupino!\n")


    return organize_groups(groups, subjects)

def organize_groups(groups, students):
    organized_groups = []

    for group in groups:
        organized_group = []
        for class_name, student_list in students.items():
            added = False
            for student in group:
                if student in student_list:
                    organized_group.append(student)
                    added = True
                    break
            if (added == False):
                organized_group.append(numpy.nan)
        organized_groups.append(organized_group)

    return organized_groups

def export_to_excel(groups):
    try:
        data_frame = pandas.DataFrame(groups)
        filename = save_excel_file()
        if (filename.endswith(".xlsx") == False):
            filename += ".xlsx"
        data_frame.to_excel(filename, index=False, header=False, sheet_name=GROUPS)
        os.system(f"{filename}")
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, ERR_SAVING_EXCEL_FILE + f"\n{e}\n\n{ERR_NOTIFY_DEVELOPER}", APP_TITLE, 0)
        exit()

def save_excel_file():
    filename = filedialog.asksaveasfilename(title=SAVE_RESULT_IN_EXCEL_FILE, filetypes=[(EXCEL_FILES, "*.xlsx")])
    return filename

def main():
    tkinter = Tk()
    tkinter.withdraw()
    filename = open_excel_file()
    students_sheet, exclusion_sheet = get_excel_data(filename)
    groups = create_groups(students_sheet, exclusion_sheet)
    export_to_excel(groups)

if __name__ == "__main__":
    main()