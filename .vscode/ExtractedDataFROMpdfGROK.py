import pandas as pd

# Define the common values
manufacturer = "PARS KAVIR ARVAND"
subject = "AC UPS-120/00-UPS/01 Up to 140/00-UPS/01"
equipment_reg = "UPS/01"
reference_serial = "SIEC-DGNRAC-ENEL-DWI-0006"
unit = "PCS."
material = "Multimaterial"
drawing_ref = ""  # Blank as not specified

# List of parts data: [total_installed, description, recommended]
parts_data = [
    (42, "Fuse, 10x38 mm, 2A", 33),
    (2, "Power Fuse, 350A", 3),
    (2, "MCB - 6A, 1P, for heater & lighting", 3),
    (20, "MCB - 10A, 2P, for distribution", 9),
    (3, "MCB - 16A, 2P, for distribution", 3),
    (5, "MCB - 20A, 2P, for distribution", 3),
    (12, "Rec. Thyristor (SCR), SKKT106", 12),
    (4, "STS Thyristor (SCR), SKKT106", 3),
    (2, "Block Diode Module", 3),
    (2, "IGBT (Transistor), CM400", 3),
    (4, "Rectifier Thyristor Driver PCB (control cards)", 3),
    (2, "IGBT Driver PCB (control cards)", 3),
    (2, "STS Thyristor Driver PCB (control cards)", 3),
    (1, "Isolator switch (make before break), 100A", 3),
    (352, "Batteries (SBM200, 176 Cells) (2 Banks)", 54)
]

# Create a list of dictionaries for the DataFrame
data_list = []
for total_installed, description, recommended in parts_data:
    data_list.append({
        "Manufacturer": manufacturer,
        "Subject": subject,
        "Equipment Reg No. or Tag No": equipment_reg,
        "Reference/Serial No.": reference_serial,
        "UNIT": unit,
        "Total number of identical parts installed": total_installed,
        "DESCRIPTION OF PARTS": description,
        "Drawing/Ref No": drawing_ref,
        "Material": material,
        "Recommended by manufacturer": recommended
    })

# Create DataFrame
df = pd.DataFrame(data_list)

# Save to Excel file
df.to_excel("extracted_parts.xlsx", index=False)

print("Excel file 'extracted_parts.xlsx' generated successfully.")