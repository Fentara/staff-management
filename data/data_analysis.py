from data.data_writer import write_table_data, auto_adjust_column_width, title_merge_format
from collections import defaultdict

def analyze_data(staff_list, program_type_list, school_list, program_list):
    # Implement your analysis functions here
    pass

def program_fte(sheet): 
    """Write the number and weighted FTE required for each program type to the program analysis sheet."""
    total_schools = len(school_list)

    # Count occurrences of each program type
    program_counts = {program.get_program(): 0 for program in program_list}
    for program in program_list:
        program_counts[program.get_program()] += 1

    # Write headers
    headers = ["Program Type", "Quantity", "FTE Weight", "Total FTE Needed"]
    write_table_data(sheet, 2, headers, cell_color='C6EFCE', font = "bold", horizontal_alignment="center", vertical_alignment="center")

    # Initialize totals
    sum_total_fte = sum_total_quantity = 0
    row_index = 3

    # Write program data
    for program in program_type_list:
        name = program.get_name()
        quantity = (
            total_schools if name == "Schools" 
            else 1 if name == "COPE" 
            else program_counts.get(name, 0)
        )
        
        fte_weight = program.get_weight()
        total_fte = quantity * fte_weight
        sum_total_fte += total_fte
        sum_total_quantity += quantity

        write_table_data(sheet, row_index, [name, quantity, fte_weight, total_fte], cell_color='C6EFCE')
        row_index += 1

    # Write total row
    write_table_data(sheet, row_index, ["Total", sum_total_quantity, "", sum_total_fte], cell_color='1AAD5A', font = "bold")
    
    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, 1, 1, len(headers), content = "Total District Program Summary")

    sheet.parent.save(output_path)

def program_area(sheet):    
    """Write the number of programs in each area to the program analysis sheet."""

    # Write headers
    headers = ["Program Type", "NW", "NE", "Central", "SW", "SE", "Total"]
    start_col = find_first_empty_column(sheet) + 1
    write_table_data(sheet, 2, headers, start_column=start_col, cell_color="C6EFCE", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Initialize area totals
    area_totals = {area: 0 for area in ["NW", "NE", "CENTRAL", "SW", "SE"]}
    program_type_counts = {p.get_name(): area_totals.copy() for p in program_type_list}

    # Count programs per area
    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_type_counts.setdefault(program.get_program(), area_totals.copy())[school.get_area()] += 1

    # Write program data to the sheet
    row_index = 3
    grand_total = 0
    for program_type, area_counts in program_type_counts.items():
        total_count = sum(area_counts.values())
        grand_total += total_count
        for area in area_totals:
            area_totals[area] += area_counts[area]
        write_table_data(sheet, row_index, [program_type, *area_counts.values(), total_count], start_column=start_col, cell_color="C6EFCE")
        row_index += 1

    # Count and write school totals
    schools_area_counts = {area: sum(1 for s in school_list if s.get_area() == area) for area in area_totals}
    total_schools_count = sum(schools_area_counts.values())
    grand_total += total_schools_count

    write_table_data(sheet, row_index, ["Schools", *schools_area_counts.values(), total_schools_count], start_column=start_col, cell_color="C6EFCE")
    row_index += 1

    # Write overall total row
    write_table_data(sheet, row_index, ["Total", *area_totals.values(), grand_total], start_column=start_col, cell_color="1AAD5A", font="bold")

    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, start_col, 1, start_col+len(headers)-1, content = "District Program Summary by Area")

    sheet.parent.save(output_path)

def program_mismatches(sheet):
    """Calculates and identifies how many and which programs are supported by psychologists who are not the school psych"""
    # Write headers
    headers = ["Program", "School", "Program Psychologist", "School Psychologist"]
    start_col = find_last_filled_column(sheet) + 1
    write_table_data(sheet, 2, headers, start_column=start_col, cell_color="FFC7CE", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Initialize totals
    row_index = 3
    mismatches = 0

    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_psych = next((s for s in staff_list if s.get_name() == program.get_psych()), None)
            if program_psych and program_psych.get_team() != "Behaviour":
                if program.get_psych() != school.get_school_psych():
                    write_table_data(sheet, row_index, [program.get_program(), program.get_school(), program.get_psych(), school.get_school_psych()], start_column=start_col, cell_color="FFC7CE")
                    row_index += 1
                    mismatches += 1

    # Write total row
    write_table_data(sheet, row_index, ["Total", "", "", mismatches], start_column=start_col, cell_color="CD5C5C", font="bold")

    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, start_col, 1, start_col+len(headers), content="Programs Psych/School Psych Mismatches")
    
    # Save the workbook
    sheet.parent.save(output_path)

def program_matches(sheet):
    """Calculates and identifies how many and which programs are supported by psychologists who are the school psych"""
    # Write headers
    headers = ["Program", "School", "Program Psychologist", "School Psychologist"]
    start_col = find_last_filled_column(sheet)
    write_table_data(sheet, 2, headers, start_column=start_col, cell_color="C6EFCE", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Initialize totals
    row_index = 3
    mismatches = 0

    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_psych = next((s for s in staff_list if s.get_name() == program.get_psych()), None)
            if program_psych and program_psych.get_team() != "Behaviour":
                if program.get_psych() == school.get_school_psych():
                    write_table_data(sheet, row_index, [program.get_program(), program.get_school(), program.get_psych(), school.get_school_psych()], start_column=start_col, cell_color="C6EFCE")
                    row_index += 1
                    mismatches += 1

    # Write total row
    write_table_data(sheet, row_index, ["Total", "", "", mismatches], start_column=start_col, cell_color="1AAD5A", font="bold")

    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, start_col, 1, start_col+len(headers), content="Programs Psych/School Psych Matches")

# Functions to Summarize Psychologists #
def program_totals_by_psych(sheet):
    """Writes the number of programs each psychologist supports in each area and their total program FTE to the psych analysis sheet."""

    # Write headers
    headers = ["Psychologist", "NW", "NE", "CENTRAL", "SW", "SE", "Total Programs", "Total Program FTE"]
    write_table_data(sheet, 2, headers, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Create lookup dictionaries for efficiency
    school_dict = {s.get_name(): s for s in school_list}
    program_dict = {pt.get_name(): pt for pt in program_type_list}

    row_index = 3
    psychologists = [p for p in staff_list if p.get_job().casefold() == "psychologist"]

    # Initialize totals
    total_area_counts = defaultdict(int)
    grand_total_programs = grand_total_fte = 0

    for psych in psychologists:
        psych_name = psych.get_name()
        area_counts = defaultdict(int)
        total_programs = total_fte = 0

        for program in program_list:
            if program.get_psych() == psych_name:
                school = school_dict.get(program.get_school())
                if school:
                    area_counts[school.get_area()] += 1
                    total_programs += 1
                    program_type = program_dict.get(program.get_program())
                    if program_type:
                        total_fte += program_type.get_weight()

        data = [
            psych_name,
            area_counts["NW"],
            area_counts["NE"],
            area_counts["CENTRAL"],
            area_counts["SW"],
            area_counts["SE"],
            total_programs,
            total_fte]
        write_table_data(sheet, row_index, data, cell_color="DBE0FF")
        row_index += 1

        # Update totals
        for area, count in area_counts.items():
            total_area_counts[area] += count
        grand_total_programs += total_programs
        grand_total_fte += total_fte

    # Write Total Row
    total_data = [
        "Total",
        total_area_counts["NW"],
        total_area_counts["NE"],
        total_area_counts["CENTRAL"],
        total_area_counts["SW"],
        total_area_counts["SE"],
        grand_total_programs,
        grand_total_fte]
    write_table_data(sheet, row_index, total_data, cell_color="848699", font="bold")
    
    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, 1, 1, len(headers), content = "Psych Program Support by Area")

    # Save the workbook
    sheet.parent.save(output_path)

def school_totals_by_psych(sheet):
    """Writes the number of schools each psychologist supports in each area to the psych analysis sheet."""

    # Determine the starting column
    start_col = find_first_empty_column(sheet) + 1

    # Write headers
    headers = ["Psychologist", "NW", "NE", "CENTRAL", "SW", "SE", "Total Schools", "Total School FTE"]
    write_table_data(sheet, 2, headers, start_column=start_col, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")

    row_index = 3
    psychologists = [p for p in staff_list if p.get_job().casefold() == "psychologist"]

    # Initialize totals
    total_area_counts = defaultdict(int)
    grand_total_schools = grand_total_weight = 0

    for psych in psychologists:
        psych_name = psych.get_name()
        area_counts = defaultdict(int)
        total_schools = total_weight = 0

        for program_type in program_type_list:
            if program_type.get_name() == "Schools":
                school_weight = program_type.get_weight()
                break

        for school in school_list:
            assigned_psych = school.get_school_psych()
            school_area = school.get_area()

            if assigned_psych == psych_name:
                area_counts[school_area] += 1
                total_schools += 1

        # Ensure at least one school is counted for the psychologist
        if total_schools > 0:
            total_weight = total_schools * school_weight
            data = [
                psych_name,
                area_counts["NW"],
                area_counts["NE"],
                area_counts["CENTRAL"],
                area_counts["SW"],
                area_counts["SE"],
                total_schools,
                total_weight]
            write_table_data(sheet, row_index, data, start_column=start_col, cell_color="DBE0FF")
            row_index += 1

            # Update totals
            for area, count in area_counts.items():
                total_area_counts[area] += count
            grand_total_schools += total_schools
            grand_total_weight += total_weight

    # Write Total Row
    total_data = [
        "Total",
        total_area_counts["NW"],
        total_area_counts["NE"],
        total_area_counts["CENTRAL"],
        total_area_counts["SW"],
        total_area_counts["SE"],
        grand_total_schools,
        grand_total_weight]
    write_table_data(sheet, row_index, total_data, start_column=start_col, cell_color="848699", font="bold")
    
    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, start_col, 1, start_col + len(headers) - 1, content="Psych School Support by Area")


    # Save the workbook
    sheet.parent.save(output_path)

def total_psych_fte(sheet):
    """Writes the total FTE of all psychologists to the psych analysis sheet."""

    # Determine the starting row and column
    row_index = 1
    start_col = find_last_filled_column(sheet) + 1

    # Write headers
    headers = ["Psychologist", "Total FTE"]
    write_table_data(sheet, 2, headers, start_column=start_col, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Create lookup dictionaries for efficiency
    program_dict = {pt.get_name(): pt for pt in program_type_list}

    # Calculate the total FTE of each psychologist
    row_index = 3
    psychologists = [p for p in staff_list if p.get_job().casefold() == "psychologist"]
    grand_total = 0

    for psych in psychologists:
        psych_name = psych.get_name()
        school_weight = next((pt.get_weight() for pt in program_type_list if pt.get_name() == "Schools"), 0)
        total_fte = 0

        for program in program_list:
            program_type = program_dict.get(program.get_program())
            if program.get_psych() == psych_name:
                if program_type:
                    total_fte += program_type.get_weight()
                    grand_total += program_type.get_weight()
        
        for school in school_list:
            if school.get_school_psych() == psych_name:
                total_fte += school_weight
                grand_total += school_weight

    # Write data to the sheet
        data = [psych_name, total_fte]
        write_table_data(sheet, row_index, data, start_column=start_col, cell_color="DBE0FF")
        row_index += 1
    
    # Write total row
    write_table_data(sheet, row_index, ["TOTAL", grand_total], start_column=start_col, cell_color="848699", font="bold")

    # Format the table
    auto_adjust_column_width(sheet)
    title_merge_format(sheet, 1, start_col, 1, start_col + len(headers)-1, "Total Assigned FTE per Psych")

    # Save the workbook
    sheet.parent.save(output_path)

# Functions for Individual Worksheets
def psychs_for_program_types(sheet):
    """Creates a sheet that lists all program types and writes all of the psychologists who support that program type."""
    
    # Write the program names across the first row
    program_names = [program.get_name() for program in program_type_list]
    write_table_data(sheet, 1, program_names, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")

    # Iterate over each program name and write the names of psychologists
    for col_index, program_name in enumerate(program_names, start=1):
        row_index = 2 
        for staff in staff_list:
            if staff.get_job().lower() == "psychologist":
                sped_programs, beh_programs = staff.get_programs()
                if program_name in sped_programs or program_name in beh_programs:
                    write_table_data(sheet, row_index, [staff.get_name()], start_column=col_index)
                    row_index += 1

    # Adjust column widths
    auto_adjust_column_width(sheet)

    # Save the workbook
    sheet.parent.save(output_path)

def psych_portfolios(sheet):
    """Creates a worksheet that lists all programs and all schools supported by each psychologist."""
    
    psych_names = [psych.get_name() for psych in staff_list if psych.get_job().lower() == "psychologist"]
    col_index = 1

    # Write "Schools" and "Programs" in the second row, alternating across columns
    col_index = 1
    while col_index <= len(psych_names) * 3:
        write_table_data(sheet, 2, ["Schools", "Programs"], start_column=col_index, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")
        col_index += 3  # Increment by 3 to leave a blank column between each psychologist

    # Write the list of schools in the "Schools" column for each psychologist
    for col_index, psych_name in enumerate(psych_names, start=1):
        row_index = 3  # Start writing schools from the third row
        for school in school_list:
            if school.get_school_psych() == psych_name:
                write_table_data(sheet, row_index, [school.get_name()], start_column=col_index * 3 - 2)
                row_index += 1

    # Write the list of programs in the "Programs" column for each psychologist
    for col_index, psych_name in enumerate(psych_names, start=1):
        row_index = 3  # Start writing programs from the third row
        for program in program_list:
            if program.get_psych() == psych_name:
                program_name = program.get_program()
                school_name = program.get_school()
                write_table_data(sheet, row_index, [f"{program_name} - {school_name}"], start_column=col_index * 3 - 1)
                row_index += 1

    # Adjust column widths
    auto_adjust_column_width(sheet)

    # Write the psychologist names across the first row
    col_index = 1
    for psych_name in psych_names:
        title_merge_format(sheet, 1, col_index, 1, col_index + 1, content=psych_name)
        col_index += 3  # Increment by 3 to leave a blank column between each psychologist

    # Save the workbook
    sheet.parent.save(output_path)
    print("Data written to CCSD Output.xlsx")

def school_distances(sheet):
    """Creates a worksheet calculating distances between schools in a psychologist's portfolio."""
    all_distances = []
    col_index = 1
    psych_distance_data = {}
    
    # Calculate the distance between each pair of schools in a psychologist's portfolio and write to the sheet (and median and max distance)
    for psych in [p for p in staff_list if p.get_job().lower() == "psychologist"]:
        psych_name = psych.get_name()
        psych_schools = [s for s in school_list if s.get_school_psych() == psych_name]
        row_index = 2
        distances = []

        for i, school1 in enumerate(psych_schools):
            for j in range(i + 1, len(psych_schools)):
                school2 = psych_schools[j]
                distance = haversine_distance(school1, school2)
                all_distances.append(distance)
                distances.append(distance)

                write_table_data(sheet, row_index, [f"{school1.get_name()} - {school2.get_name()}", distance], start_column=col_index)
                row_index += 1
        
        if not distances:
            continue

        # Calculate the median and maximum distance between schools in the portfolio and assign to psych_distance_data
        median_distance = statistics.median(distances)
        max_distance = max(distances)
        psych_distance_data[psych_name] = (median_distance, max_distance)

        write_table_data(sheet, row_index, ["Median Distance", median_distance], start_column=col_index)
        row_index += 1

        write_table_data(sheet, row_index, ["Max Distance", max_distance], start_column=col_index)

        # Write the psychologist's name in the first row
        title_merge_format(sheet, 1, col_index, 1, col_index + 1, content=psych_name)

        col_index += 3  # Increment by 3 to leave a blank column between each psychologist

    # Calculate the median and maximum distance between schools in all portfolios
    if all_distances:
        median_distance_all = statistics.median(all_distances)
        max_distance_all = max(all_distances)

    # Create Summary Table 
    summary_col_index = sheet.max_column + 2  # Ensures it appears after existing data
    write_table_data(sheet, 1, ["Psychologist", "Median Distance", "Max Distance"], start_column=summary_col_index, cell_color="DBE0FF", font="bold", horizontal_alignment="center", vertical_alignment="center")
    for row_index, (psych_name, (median_distance, max_distance)) in enumerate(psych_distance_data.items(), start=2):
        write_table_data(sheet, row_index, [psych_name, median_distance, max_distance], start_column=summary_col_index)
    write_table_data(sheet, row_index + 1, ["ALL Psychs", round(median_distance_all, 2), round(max_distance_all, 2)], start_column=summary_col_index, cell_color="848699", font="bold")

    # Adjust column widths
    auto_adjust_column_width(sheet)

    # Save the workbook
    sheet.parent.save(output_path)
    print("Data written to CCSD Output.xlsx")