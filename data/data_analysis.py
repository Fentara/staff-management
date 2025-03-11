import statistics
from data.helpers import (
    write_table_data, auto_adjust_column_width, title_merge_format, 
    find_first_empty_column, find_last_filled_column, haversine_distance,
    AREAS, TABLE_COLORS, get_psychologists, create_area_dict,
    write_table_header, write_total_row, format_table
)
from collections import defaultdict

# Functions to Summarize Programs #
def program_fte(sheet, staff_list, program_type_list, school_list, program_list, output_path): 
    """Write the number and weighted FTE required for each program type to the program analysis sheet."""
    total_schools = len(school_list)

    # Count occurrences of each program type
    program_counts = {program.get_program(): 0 for program in program_list}
    for program in program_list:
        program_counts[program.get_program()] += 1

    # Write headers
    headers = ["Program Type", "Quantity", "FTE Weight", "Total FTE Needed"]
    write_table_header(sheet, 2, headers, color_key="Light green")

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

        write_table_data(sheet, row_index, [name, quantity, fte_weight, total_fte], 
                        cell_color=TABLE_COLORS["Light green"])
        row_index += 1

    # Write total row
    write_total_row(sheet, row_index, ["Total", sum_total_quantity, "", sum_total_fte], 
                   color_key="Dark green")
    
    # Format the table
    format_table(sheet, 1, 1, len(headers), "Total District Program Summary")

def program_area(sheet, staff_list, program_type_list, school_list, program_list, output_path):    
    """Write the number of programs in each area to the program analysis sheet."""

    # Write headers
    headers = ["Program Type", *AREAS, "Total"]
    start_col = find_first_empty_column(sheet) + 1
    write_table_header(sheet, 2, headers, start_col, color_key="Light green")

    # Initialize area totals
    area_totals = create_area_dict()
    program_type_counts = {p.get_name(): create_area_dict() for p in program_type_list}

    # Count programs per area
    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_type_counts.setdefault(program.get_program(), create_area_dict())[school.get_area()] += 1

    # Write program data to the sheet
    row_index = 3
    grand_total = 0
    for program_type, area_counts in program_type_counts.items():
        total_count = sum(area_counts.values())
        grand_total += total_count
        for area in AREAS:
            area_totals[area] += area_counts[area]
        
        # Write row for this program type
        row_data = [program_type]
        for area in AREAS:
            row_data.append(area_counts[area])
        row_data.append(total_count)
        
        write_table_data(sheet, row_index, row_data, start_column=start_col, 
                        cell_color=TABLE_COLORS["Light green"])
        row_index += 1

    # Count and write school totals
    schools_area_counts = {area: sum(1 for s in school_list if s.get_area() == area) for area in AREAS}
    total_schools_count = sum(schools_area_counts.values())
    grand_total += total_schools_count

    school_row = ["Schools"]
    for area in AREAS:
        school_row.append(schools_area_counts[area])
    school_row.append(total_schools_count)
    
    write_table_data(sheet, row_index, school_row, start_column=start_col, 
                    cell_color=TABLE_COLORS["Light green"])
    row_index += 1

    # Write overall total row
    total_row = ["Total"]
    for area in AREAS:
        total_row.append(area_totals[area])
    total_row.append(grand_total)
    
    write_total_row(sheet, row_index, total_row, start_col, color_key="Dark green")
    
    # Format the table
    format_table(sheet, 1, start_col, start_col + len(headers) - 1, 
                "District Program Summary by Area")

def program_mismatches(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Calculates and identifies how many and which programs are supported by psychologists who are not the school psych"""
    
    # Write headers
    headers = ["Program", "School", "Program Psychologist", "School Psychologist"]
    start_col = find_last_filled_column(sheet) + 1
    write_table_header(sheet, 2, headers, start_col, color_key="Light red")

    # Initialize totals
    row_index = 3
    mismatches = 0

    # Find mismatched program psychologists
    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_psych = next((s for s in staff_list if s.get_name() == program.get_psych()), None)
            if program_psych and program_psych.get_team() != "Behaviour":
                if program.get_psych() != school.get_school_psych():
                    # Add mismatch to table
                    write_table_data(
                        sheet, row_index, 
                        [program.get_program(), program.get_school(), program.get_psych(), school.get_school_psych()], 
                        start_column=start_col, 
                        cell_color=TABLE_COLORS["Light red"]
                    )
                    row_index += 1
                    mismatches += 1

    # Write total row
    write_total_row(
        sheet, row_index, 
        ["Total", "", "", mismatches], 
        start_col, 
        color_key="Dark red"
    )

    # Format the table
    format_table(
        sheet, 1, start_col, start_col + len(headers) - 1, 
        "Programs Psych/School Psych Mismatches"
    )

def program_matches(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Calculates and identifies how many and which programs are supported by psychologists who are the school psych"""
    
    # Write headers
    headers = ["Program", "School", "Program Psychologist", "School Psychologist"]
    start_col = find_last_filled_column(sheet) + 1
    write_table_header(sheet, 2, headers, start_col, color_key="Light green")

    # Initialize totals
    row_index = 3
    matches = 0

    # Find matched program psychologists
    for program in program_list:
        school = next((s for s in school_list if s.get_name() == program.get_school()), None)
        if school:
            program_psych = next((s for s in staff_list if s.get_name() == program.get_psych()), None)
            if program_psych and program_psych.get_team() != "Behaviour":
                if program.get_psych() == school.get_school_psych():
                    # Add match to table
                    write_table_data(
                        sheet, row_index, 
                        [program.get_program(), program.get_school(), program.get_psych(), school.get_school_psych()], 
                        start_column=start_col, 
                        cell_color=TABLE_COLORS["Light green"]
                    )
                    row_index += 1
                    matches += 1

    # Write total row
    write_total_row(
        sheet, row_index, 
        ["Total", "", "", matches], 
        start_col, 
        color_key="Dark green"
    )

    # Format the table
    format_table(
        sheet, 1, start_col, start_col + len(headers) - 1, 
        "Programs Psych/School Psych Matches"
    )

# Functions to Summarize Psychologists #
def program_totals_by_psych(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Writes the number of programs each psychologist supports in each area and their total program FTE."""

    # Write headers
    headers = ["Psychologist", *AREAS, "Total Programs", "Total Program FTE"]
    write_table_header(sheet, 2, headers, color_key="Light blue")

    # Create lookup dictionaries for efficiency
    school_dict = {s.get_name(): s for s in school_list}
    program_dict = {pt.get_name(): pt for pt in program_type_list}

    # Get all psychologists
    psychologists = get_psychologists(staff_list)
    row_index = 3

    # Initialize totals
    total_area_counts = create_area_dict()
    grand_total_programs = grand_total_fte = 0

    for psych in psychologists:
        psych_name = psych.get_name()
        area_counts = create_area_dict()
        total_programs = total_fte = 0

        # Count programs by area and calculate FTE
        for program in program_list:
            if program.get_psych() == psych_name:
                school = school_dict.get(program.get_school())
                if school:
                    area_counts[school.get_area()] += 1
                    total_programs += 1
                    program_type = program_dict.get(program.get_program())
                    if program_type:
                        total_fte += program_type.get_weight()

        # Create data row with area counts
        data = [
            psych_name,
            *[area_counts[area] for area in AREAS],  # Unpack area counts in order
            total_programs,
            total_fte
        ]
        
        write_table_data(sheet, row_index, data, cell_color=TABLE_COLORS["Light blue"])
        row_index += 1

        # Update totals
        for area in AREAS:
            total_area_counts[area] += area_counts[area]
        grand_total_programs += total_programs
        grand_total_fte += total_fte

    # Write Total Row
    total_data = [
        "Total",
        *[total_area_counts[area] for area in AREAS],  # Unpack area totals in order
        grand_total_programs,
        grand_total_fte
    ]
    
    write_total_row(sheet, row_index, total_data, color_key="Dark blue")
    
    # Format the table
    format_table(sheet, 1, 1, len(headers), "Psych Program Support by Area")

def school_totals_by_psych(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Writes the number of schools each psychologist supports in each area to the psych analysis sheet."""

    # Determine the starting column
    start_col = find_first_empty_column(sheet) + 1

    # Write headers
    headers = ["Psychologist", *AREAS, "Total Schools", "Total School FTE"]
    write_table_header(sheet, 2, headers, start_col, color_key="Light blue")

    row_index = 3
    psychologists = get_psychologists(staff_list)

    # Initialize totals
    total_area_counts = create_area_dict()
    grand_total_schools = grand_total_weight = 0

    # Get school weight once to avoid repeated lookups
    school_weight = next((pt.get_weight() for pt in program_type_list 
                         if pt.get_name() == "Schools"), 0)

    for psych in psychologists:
        psych_name = psych.get_name()
        area_counts = create_area_dict()
        total_schools = 0

        # Count schools by area for this psychologist
        for school in school_list:
            if school.get_school_psych() == psych_name:
                area_counts[school.get_area()] += 1
                total_schools += 1

        # Ensure at least one school is counted for the psychologist
        if total_schools > 0:
            total_weight = total_schools * school_weight
            
            # Prepare row data using area constants to ensure order
            data = [
                psych_name,
                *[area_counts[area] for area in AREAS],  # Unpack area counts in order
                total_schools,
                total_weight
            ]
            
            write_table_data(sheet, row_index, data, start_column=start_col, 
                            cell_color=TABLE_COLORS["Light blue"])
            row_index += 1

            # Update totals
            for area in AREAS:
                total_area_counts[area] += area_counts[area]
            grand_total_schools += total_schools
            grand_total_weight += total_weight

    # Write Total Row
    total_data = [
        "Total",
        *[total_area_counts[area] for area in AREAS],  # Unpack area totals in order
        grand_total_schools,
        grand_total_weight
    ]
    
    write_total_row(sheet, row_index, total_data, start_col, color_key="Dark blue")
    
    # Format the table
    format_table(sheet, 1, start_col, start_col + len(headers) - 1, 
                "Psych School Support by Area")

def total_psych_fte(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Writes the total FTE of all psychologists to the psych analysis sheet."""

    # Determine the starting column
    start_col = find_last_filled_column(sheet) + 1

    # Write headers
    headers = ["Psychologist", "Total FTE"]
    write_table_header(sheet, 2, headers, start_col, color_key="Light blue")

    # Create lookup dictionaries for efficiency
    program_dict = {pt.get_name(): pt for pt in program_type_list}
    
    # Get school weight once to avoid repeated lookups
    school_weight = next((pt.get_weight() for pt in program_type_list 
                         if pt.get_name() == "Schools"), 0)

    # Calculate the total FTE of each psychologist
    row_index = 3
    psychologists = get_psychologists(staff_list)
    grand_total = 0

    for psych in psychologists:
        psych_name = psych.get_name()
        total_fte = 0

        # Calculate program FTE
        for program in program_list:
            if program.get_psych() == psych_name:
                program_type = program_dict.get(program.get_program())
                if program_type:
                    total_fte += program_type.get_weight()
        
        # Calculate school FTE
        for school in school_list:
            if school.get_school_psych() == psych_name:
                total_fte += school_weight

        # Write data to the sheet
        write_table_data(sheet, row_index, [psych_name, total_fte], 
                         start_column=start_col, cell_color=TABLE_COLORS["Light blue"])
        row_index += 1
        grand_total += total_fte
    
    # Write total row
    write_total_row(sheet, row_index, ["Total", grand_total], start_col, color_key="Dark blue")
    
    # Format the table
    format_table(sheet, 1, start_col, start_col + len(headers) - 1, 
                "Total Assigned FTE per Psych")

# Functions for Individual Worksheets
def psychs_for_program_types(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Creates a sheet that lists all program types and writes all of the psychologists who support that program type."""
    
    # Write the program names across the first row
    program_names = [program.get_name() for program in program_type_list]
    write_table_data(sheet, 1, program_names, 
                    cell_color=TABLE_COLORS["Light blue"], 
                    font="bold", 
                    horizontal_alignment="center", 
                    vertical_alignment="center")

    # Iterate over each program name and write the names of psychologists
    for col_index, program_name in enumerate(program_names, start=1):
        row_index = 2
        psychologists = get_psychologists(staff_list)
        
        for psych in psychologists:
            sped_programs, beh_programs = psych.get_programs()
            if program_name in sped_programs or program_name in beh_programs:
                write_table_data(sheet, row_index, [psych.get_name()], start_column=col_index)
                row_index += 1

    # Adjust column widths
    auto_adjust_column_width(sheet)

def psych_portfolios(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Creates a worksheet that lists all programs and all schools supported by each psychologist."""
    
    # Get all psychologists
    psychologists = get_psychologists(staff_list)
    psych_names = [psych.get_name() for psych in psychologists]
    
    # Write "Schools" and "Programs" in the second row, alternating across columns
    col_index = 1
    while col_index <= len(psych_names) * 3:
        write_table_data(sheet, 2, ["Schools", "Programs"], 
                        start_column=col_index, 
                        cell_color=TABLE_COLORS["Light blue"], 
                        font="bold", 
                        horizontal_alignment="center", 
                        vertical_alignment="center")
        col_index += 3  # Increment by 3 to leave a blank column between each psychologist

    # Write the list of schools in the "Schools" column for each psychologist
    for col_index, psych_name in enumerate(psych_names, start=1):
        row_index = 3  # Start writing schools from the third row
        for school in school_list:
            if school.get_school_psych() == psych_name:
                write_table_data(sheet, row_index, [school.get_name()], 
                                start_column=col_index * 3 - 2)
                row_index += 1

    # Write the list of programs in the "Programs" column for each psychologist
    for col_index, psych_name in enumerate(psych_names, start=1):
        row_index = 3  # Start writing programs from the third row
        for program in program_list:
            if program.get_psych() == psych_name:
                program_name = program.get_program()
                school_name = program.get_school()
                write_table_data(sheet, row_index, [f"{program_name} - {school_name}"], 
                                start_column=col_index * 3 - 1)
                row_index += 1

    # Adjust column widths
    auto_adjust_column_width(sheet)

    # Write the psychologist names across the first row
    col_index = 1
    for psych_name in psych_names:
        title_merge_format(sheet, 1, col_index, 1, col_index + 1, content=psych_name)
        col_index += 3  # Increment by 3 to leave a blank column between each psychologist

def school_distances(sheet, staff_list, program_type_list, school_list, program_list, output_path):
    """Creates a worksheet calculating distances between schools in a psychologist's portfolio."""
    all_distances = []
    col_index = 1
    psych_distance_data = {}
    
    # Calculate distances between schools for each psychologist
    psychologists = get_psychologists(staff_list)
    for psych in psychologists:
        psych_name = psych.get_name()
        psych_schools = [s for s in school_list if s.get_school_psych() == psych_name]
        row_index = 2
        distances = []

        for i, school1 in enumerate(psych_schools):
            for j in range(i + 1, len(psych_schools)):
                school2 = psych_schools[j]
                distance = haversine_distance(school1, school2)
                if distance is not None:  # Check for valid distance
                    all_distances.append(distance)
                    distances.append(distance)

                    write_table_data(
                        sheet, row_index, 
                        [f"{school1.get_name()} - {school2.get_name()}", distance], 
                        start_column=col_index
                    )
                    row_index += 1
        
        if not distances:
            continue

        # Calculate statistics for this psychologist
        median_distance = statistics.median(distances)
        max_distance = max(distances)
        psych_distance_data[psych_name] = (median_distance, max_distance)

        # Write summary for this psychologist
        write_table_data(sheet, row_index, ["Median Distance", median_distance], start_column=col_index)
        row_index += 1
        write_table_data(sheet, row_index, ["Max Distance", max_distance], start_column=col_index)

        # Write the psychologist's name in the first row
        title_merge_format(sheet, 1, col_index, 1, col_index + 1, content=psych_name)

        col_index += 3  # Leave a blank column between psychologists

    # Calculate overall statistics
    if all_distances:
        median_distance_all = statistics.median(all_distances)
        max_distance_all = max(all_distances)
    else:
        median_distance_all = max_distance_all = 0

    # Create Summary Table 
    summary_col_index = sheet.max_column + 2  # Appears after existing data
    headers = ["Psychologist", "Median Distance", "Max Distance"]
    write_table_header(sheet, 1, headers, summary_col_index, color_key="Light blue")
    
    row_index = 2
    for psych_name, (median_distance, max_distance) in psych_distance_data.items():
        write_table_data(
            sheet, row_index, 
            [psych_name, median_distance, max_distance], 
            start_column=summary_col_index
        )
        row_index += 1
    
    # Write all psychs row (total row)
    write_total_row(
        sheet, row_index, 
        ["ALL Psychs", round(median_distance_all, 2), round(max_distance_all, 2)], 
        summary_col_index, 
        color_key="Dark blue"
    )
    
    # Format the table
    format_table(
        sheet, 0, summary_col_index, summary_col_index + len(headers) - 1, 
        "School Distance Summary"
    )
    
    # Adjust column widths
    auto_adjust_column_width(sheet)

    # # Add a title to the individual psychologist section
    # title_merge_format(sheet, 1, 1, 1, col_index-2, content="Distance Between Schools by Psychologist")
    