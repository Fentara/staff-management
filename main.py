"""
Main module for CCSD staff management system.
Orchestrates data loading, analysis, and output generation.
"""

from data.data_loader import load_data
from data.data_writer import write_data

def main():
    """
    Main function that executes the data pipeline:
    1. Loads staff and program data
    2. Writes and analyzes results to output files
    
    Returns:
        int: 0 for success, non-zero for errors
    """
    try:
        # Load staff data
        print("Loading data...")
        staff_list, program_type_list, school_list, program_list = load_data()
        
        # Write data and perform analysis
        print("Analyzing data and writing results...")
        write_data(staff_list, program_type_list, school_list, program_list)
        
        print("Process completed successfully!")
        return 0
    except Exception as e:
        print(f"Error: {e}")
        return 1

if __name__ == "__main__":
    exit(main())