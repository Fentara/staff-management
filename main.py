from data.data_loader import load_data
from data.data_analysis import analyze_data
from data.data_writer import write_data

def main():
    staff_list, program_type_list, school_list, program_list = load_data()
    analyze_data(staff_list, program_type_list, school_list, program_list)
    write_data()

if __name__ == "__main__":
    main()