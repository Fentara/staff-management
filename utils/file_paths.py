import os
from pathlib import Path

def get_paths():
    # Try multiple possible locations
    possible_paths = [
        Path("C:/Users/david.williamson/OneDrive - Calgary Catholic School District/Dave/Python/CCSD Management/ccsd data.xlsx"),
        Path("C:/Users/Dave/OneDrive - Calgary Catholic School District/Dave/Python/CCSD Management/ccsd data.xlsx"),
        Path.home() / "OneDrive - Calgary Catholic School District/Dave/Python/CCSD Management/ccsd data.xlsx"
    ]
    
    # Return the first path that exists
    for path in possible_paths:
        if path.exists():
            return str(path.parent), str(path.parent), str(path)
    
    # If no path exists, raise a meaningful error
    raise FileNotFoundError("Could not find CCSD data file in any expected location")

def get_output_path():
    _, _, path = get_paths()
    return os.path.join(os.path.dirname(path), "CCSD Output.xlsx")