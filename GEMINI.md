# Project Overview

**Quantify** is a Python-based utility designed to parse civil engineering sewer network model files (such as `.mhc`, `.inv`, `.ngl`, and `.brn` files typically exported from civil design software) and automatically generate an Excel spreadsheet for quantity takeoffs. The generated spreadsheet (`Quantified_sewer.xlsx`) calculates key metrics like trench excavation volumes at varying depth bands, pipe lengths, and backfilling material requirements.

## Key Technologies
- **Python:** The primary programming language used.
- **openpyxl:** Used for generating and formatting the output Excel spreadsheet.

## Directory Overview

- `quantify.py`: The main script containing the logic for parsing the model files, extracting pipe/node data (chainages, invert levels, natural ground levels, pipe diameters), calculating volumes, and writing to Excel.
- `test.py`: A secondary script seemingly used for testing specific parsing logic (e.g., reading `.mhc` files).
- `Model Files/`: A directory containing the input civil model files (e.g., `Asbuilt - Sewer.Mhc`, `*.Inv`, `*.Ngl`).
- `Quantities Template/`: Contains existing Excel quantity templates.
- `Dwg/`: Contains PDF drawings related to the project.
- `_bmad/` & `.gemini/`: Configuration directories for AI assistant tools.

## Building and Running

### Prerequisites
You need Python installed along with the `openpyxl` library.
If it's not installed, you can install it via pip:
```bash
pip install openpyxl
```

### Execution
To run the main quantization script:
```bash
python quantify.py
```

Upon execution, the script is interactive and will prompt you via the terminal for:
1. The base filename of the MHC model (e.g., `Asbuilt - Sewer`).
2. The filepath to the folder containing the model files (e.g., the absolute path to the `Model Files` directory).
3. Network processing options (e.g., `1.Finish`, `2.Modelled Network`, `3.Customize Network`), allowing you to quantify the entire modeled network or specify specific upstream/downstream nodes.

After processing, it will output a file named `Quantified_sewer.xlsx` in the current working directory.

## Development Conventions

- The script relies heavily on reading specific text-based output files from civil design software, extracting node IDs, chainages, slopes, and invert levels. 
- Excel generation is heavily hardcoded with specific column widths, formulas, and formatting (borders and cell fill colors) applied directly via `openpyxl` to produce a ready-to-print bill of quantities format.
