import openpyxl
from utils.file_paths import get_paths
from models.staff import Staff
from models.program_type import ProgramType
from models.school import School
from models.program import Program

def load_data():
    primary_path, alternate_path, path = get_paths()
    wb_obj = openpyxl.load_workbook(path)

    staff_list = create_staff(wb_obj)
    program_type_list = create_program_types(wb_obj)
    school_list = create_schools(wb_obj)
    program_list = create_programs(wb_obj)

    return staff_list, program_type_list, school_list, program_list

def create_staff(wb_obj):
    staff_sheet = wb_obj['staff']
    staff_list = []
    for i in range(2, staff_sheet.max_row + 1):
        name = staff_sheet.cell(row=i, column=1).value
        job = staff_sheet.cell(row=i, column=2).value
        fte = staff_sheet.cell(row=i, column=3).value
        team = staff_sheet.cell(row=i, column=4).value
        sped_programs = staff_sheet.cell(row=i, column=5).value
        beh_programs = staff_sheet.cell(row=i, column=6).value

        sped_programs = sped_programs.split(', ') if sped_programs else []
        beh_programs = beh_programs.split(', ') if beh_programs else []

        staff_member = Staff(name, job, fte, team, sped_programs, beh_programs)
        staff_list.append(staff_member)
    return staff_list

def create_program_types(wb_obj):
    program_type_sheet = wb_obj['program_types']
    program_type_list = []
    for i in range(2, program_type_sheet.max_row + 1):
        name = program_type_sheet.cell(row=i, column=1).value
        team = program_type_sheet.cell(row=i, column=2).value
        adaptive_func = program_type_sheet.cell(row=i, column=3).value
        cognitive_func = program_type_sheet.cell(row=i, column=4).value
        soc_emo_beh_func = program_type_sheet.cell(row=i, column=5).value
        phys_med_need = program_type_sheet.cell(row=i, column=6).value
        weight = program_type_sheet.cell(row=i, column=7).value

        program = ProgramType(name, team, adaptive_func, cognitive_func, soc_emo_beh_func, phys_med_need, weight)
        program_type_list.append(program)
    return program_type_list

def create_schools(wb_obj):
    school_sheet = wb_obj['schools']
    program_sheet = wb_obj['programs']
    school_list = []
    for i in range(2, school_sheet.max_row + 1):
        name = school_sheet.cell(row=i, column=1).value
        area = school_sheet.cell(row=i, column=2).value
        school_psych = school_sheet.cell(row=i, column=3).value
        address = school_sheet.cell(row=i, column=4).value
        latitude_radian = school_sheet.cell(row=i, column=7).value
        longitude_radian = school_sheet.cell(row=i, column=8).value

        programs = []
        for j in range(2, program_sheet.max_row + 1):
            program_school_name = program_sheet.cell(row=j, column=1).value
            if program_school_name == name:
                program_name = program_sheet.cell(row=j, column=3).value
                programs.append(program_name)

        school = School(name, area, school_psych, address, latitude_radian, longitude_radian, programs)
        school_list.append(school)
    return school_list

def create_programs(wb_obj):
    program_sheet = wb_obj['programs']
    program_list = []
    for i in range(2, program_sheet.max_row + 1):
        school = program_sheet.cell(row=i, column=1).value
        program_type = program_sheet.cell(row=i, column=3).value
        psych = program_sheet.cell(row=i, column=4).value

        program = Program(school, program_type, psych)
        program_list.append(program)
    return program_list