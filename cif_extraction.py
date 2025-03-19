"""
----------------------------------------------------------
ID15B-Postdocs presents
----------------------------------------------------------
"""

import os
import pandas as pd
import re

# Directory containing the .cif files
cif_folder = "C:/Users/mondal/Desktop/test"  # Change this to your folder path

# List to store data
cif_data = []
column_headers = ["File_index", "Filename", "Pressure", "a (Å)", "b (Å)", "c (Å)", "α (°)", "β (°)", "γ (°)", "vol (Å³)"]

# Function to extract float values from numbers with uncertainties (e.g., "0.21660(5)")
def extract_float(value):
    return float(re.sub(r"\([^)]*\)", "", value))  # Removes anything inside parentheses

# Loop through all .cif files in the folder
file_index = 1  # Counter for File_index
for file in sorted(os.listdir(cif_folder)):  # Sorted for consistency
    if file.endswith(".cif"):
        file_path = os.path.join(cif_folder, file)

        with open(file_path, "r") as f:
            lines = f.readlines()

        # Initialize variables
        a = b = c = alpha = beta = gamma = vol = None
        atoms = {}

        # Read the CIF file line by line
        read_atoms = False
        for line in lines:
            line = line.strip()

            # Extract lattice parameters
            if line.startswith("_cell_length_a"):
                a = extract_float(line.split()[-1])
            elif line.startswith("_cell_length_b"):
                b = extract_float(line.split()[-1])
            elif line.startswith("_cell_length_c"):
                c = extract_float(line.split()[-1])
            elif line.startswith("_cell_angle_alpha"):
                alpha = extract_float(line.split()[-1])
            elif line.startswith("_cell_angle_beta"):
                beta = extract_float(line.split()[-1])
            elif line.startswith("_cell_angle_gamma"):
                gamma = extract_float(line.split()[-1])
            elif line.startswith("_cell_volume"):
                vol = extract_float(line.split()[-1])

            # Detect the beginning of atomic site positions section
            if line.startswith("_atom_site_label"):
                read_atoms = True
                continue

            # Stop reading when anisotropic section starts
            if line.startswith("_atom_site_aniso_label"):
                read_atoms = False  # Ignore anisotropic parameters
                continue

            if read_atoms:
                parts = line.split()
                if len(parts) >= 5:
                    element = parts[0]  # Atom name
                    try:
                        x, y, z = extract_float(parts[2]), extract_float(parts[3]), extract_float(parts[4])
                        atoms[element] = (x, y, z)  # Store as dictionary { "Mo01": (x, y, z) }
                    except ValueError:
                        continue  # Skip if there is a format issue

        # Ensure all atom types are included in column headers dynamically
        for atom in atoms.keys():
            if f"{atom}_x" not in column_headers:
                column_headers.extend([f"{atom}_x", f"{atom}_y", f"{atom}_z"])

        # Store the extracted data in a dictionary format
        cif_entry = {
            "File_index": file_index,  # Assign unique index
            "Filename": file, 
            "Pressure": "",  # Blank column for manual entry
            "a (Å)": a, "b (Å)": b, "c (Å)": c, 
            "α (°)": alpha, "β (°)": beta, "γ (°)": gamma, "vol (Å³)": vol
        }
        for atom, (x, y, z) in atoms.items():
            cif_entry[f"{atom}_x"] = x
            cif_entry[f"{atom}_y"] = y
            cif_entry[f"{atom}_z"] = z

        cif_data.append(cif_entry)
        file_index += 1  # Increment file index for next file

# Convert data to DataFrame and ensure all columns exist
df = pd.DataFrame(cif_data, columns=column_headers)

# Save to Excel if data exists
if not df.empty:
    output_file = "cif_exported_data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
    print("All thanks to ID15B-Postdocs")
else:
    print("No valid data found in the CIF files. Check file formatting.")

