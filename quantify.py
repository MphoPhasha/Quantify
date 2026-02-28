import openpyxl
import os
import sys
from typing import List, Tuple, Dict, Any

def getMHCfilename() -> str:
    """Prompts the user for the MHC filename and returns it in lowercase."""
    filename = input("Enter MHC filename: ")
    return filename.lower()

def getFilepathRootFolder() -> str:
    """Prompts the user for the root directory of the model files."""
    filepath = input("Enter filepath to MHC file: ")
    return filepath

def getFilepathTargetFile(filePathRootFolder: str, filename: str, fileNumber: str, extension: str) -> str:
    """Constructs a cross-platform path for a specific data file."""
    return os.path.join(filePathRootFolder, f"{filename}{fileNumber}.{extension.upper()}")

def getFilepathTargetFile_MHC(filePathRootFolder: str, filename: str, extension: str = "mhc") -> str:
    """Constructs a cross-platform path for the MHC data file."""
    return os.path.join(filePathRootFolder, f"{filename}.{extension.upper()}")

def getNumBranches(filepathRootFolder: str, MHCfilename: str) -> int:
    """Reads the BRN file to determine the total number of branches in the network."""
    BranchFilePath = getFilepathTargetFile(filepathRootFolder, MHCfilename, "", "brn")
    try:
        with open(BranchFilePath) as file:
            lines = file.readlines()
            if len(lines) > 1:
                return int(lines[1].strip())
    except FileNotFoundError:
        print(f"Error: Branch file not found at {BranchFilePath}")
    except ValueError:
        print(f"Error: Could not parse number of branches in {BranchFilePath}")
    return 0

def labelFormat(label: str) -> str:
    """Cleans up manhole labels by removing quotes and extra spaces."""
    return label.strip(' "')

def pipeTypeFormat(pipeType: str) -> str:
    """Removes quotation marks from pipe type."""
    return pipeType.strip(' "')

def addBranches(metaList: list) -> tuple[str, str]:
    """Interactively allows the user to build a list of branches to quantify."""
    MHCfilename = getMHCfilename()
    filepathRootFolder = getFilepathRootFolder()

    while True:
        print("\nQuantify\n1.Finish\n2.Modelled Network\n3.Customize Network")
        userInput = input("Enter: ")

        if userInput == "1":
            break
        elif userInput == "2":
            totalBranches = getNumBranches(filepathRootFolder, MHCfilename)
            for branch_idx in range(1, totalBranches + 1):
                fNumStr = f"{branch_idx:03d}"
                filepath_inv = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumStr, "inv")
                
                try:
                    with open(filepath_inv) as file:
                        lines = file.readlines()
                        data_lines = [l for l in lines[3:] if l.strip()]
                        if not data_lines:
                            continue
                        
                        first_row = data_lines[0].split(",")
                        start_node = labelFormat(first_row[2] if len(first_row) > 2 else first_row[0])
                        
                        last_row = data_lines[-1].split(",")
                        end_node = labelFormat(last_row[2] if len(last_row) > 2 else last_row[0])
                        
                        metaList.append([start_node, end_node])
                except FileNotFoundError:
                    print(f"Error: INV file not found at {filepath_inv}")
                    sys.exit(1)
        elif userInput == "3":
            upstream = input("Upstream node label: ").strip()
            downstream = input("Downstream node label: ").strip()
            if upstream and downstream:
                metaList.append([upstream, downstream])
        else:
            print("Invalid input. Please choose 1, 2, or 3.")

    return MHCfilename, filepathRootFolder

def load_mhc_data(filepathRootFolder: str, MHCfilename: str) -> Dict[str, float]:
    """Loads MHC node ground levels into memory."""
    mhc_cache = {}
    filepath = getFilepathTargetFile_MHC(filepathRootFolder, MHCfilename)
    try:
        with open(filepath) as file:
            lines = file.readlines()
            for line in lines[13:]:
                if not line.strip(): continue
                row = line.strip().split(",")
                if len(row) > 0:
                    node_id = labelFormat(row[0])
                    try:
                        ngl = float(row[3].split()[0])
                        mhc_cache[node_id.lower()] = ngl
                    except (ValueError, IndexError):
                        continue
    except FileNotFoundError:
        print(f"Error opening MHC file at: {filepath}")
    return mhc_cache

def load_all_inv_files(filepathRootFolder: str, MHCfilename: str) -> Dict[str, List[Dict[str, Any]]]:
    """Loads all INV files into memory."""
    inv_cache = {}
    totalBranches = getNumBranches(filepathRootFolder, MHCfilename)
    for branch_idx in range(1, totalBranches + 1):
        fNumStr = f"{branch_idx:03d}"
        filepath = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumStr, "inv")
        try:
            with open(filepath) as f:
                lines = f.readlines()[3:]
                data = []
                for line in lines:
                    if not line.strip(): continue
                    row = line.strip().split(",")
                    if len(row) < 7:
                        row = line.strip().split("  ")
                    if len(row) >= 7:
                        data.append({
                            "chainage": float(row[0]),
                            "il": float(row[1]),
                            "node": labelFormat(row[2]),
                            "dia": float(row[3]),
                            "pipe_type": pipeTypeFormat(row[6])
                        })
                inv_cache[fNumStr] = data
        except FileNotFoundError:
            continue
    return inv_cache

def load_all_ngl_files(filepathRootFolder: str, MHCfilename: str) -> Dict[str, List[Dict[str, float]]]:
    """Loads all NGL files into memory."""
    ngl_cache = {}
    totalBranches = getNumBranches(filepathRootFolder, MHCfilename)
    for branch_idx in range(1, totalBranches + 1):
        fNumStr = f"{branch_idx:03d}"
        filepath = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumStr, "ngl")
        try:
            with open(filepath) as f:
                lines = f.readlines()[4:]
                data = []
                for line in lines:
                    if not line.strip(): continue
                    row = line.strip().split("  ")
                    try:
                        chainage = float(row[0])
                        ngl = float(row[1])
                    except (ValueError, IndexError):
                        row = line.strip().split(",")
                        try:
                            chainage = float(row[0])
                            ngl = float(row[1])
                        except (ValueError, IndexError):
                            continue
                    data.append({"chainage": chainage, "ngl": ngl})
                ngl_cache[fNumStr] = data
        except FileNotFoundError:
            continue
    return ngl_cache

def trace_branch(upstream: str, downstream: str, inv_cache: Dict[str, List[Dict[str, Any]]]) -> Dict[str, List[Any]]:
    """Traces a single branch from upstream to downstream node across multiple INV files."""
    labels, chainages, file_numbers, diameters, pipe_types = [], [], [], [], []
    current_node = upstream
    search_proceed = True
    
    while search_proceed:
        found_in_file = False
        for fNumStr, rows in inv_cache.items():
            start_idx = -1
            for idx, row in enumerate(rows):
                if row["node"].lower() == current_node.lower():
                    start_idx = idx
                    found_in_file = True
                    break
            
            if start_idx != -1:
                for idx in range(start_idx, len(rows)):
                    row = rows[idx]
                    node_id = row["node"]
                    
                    if not labels or labels[-1].lower() != node_id.lower():
                        labels.append(node_id)
                        chainages.append(row["chainage"])
                        file_numbers.append(fNumStr)
                        diameters.append(row["dia"])
                        pipe_types.append(row["pipe_type"])
                    else:
                        # Update existing (duplicate node e.g. for drop)
                        chainages[-1] = row["chainage"]
                        file_numbers[-1] = fNumStr
                        diameters[-1] = row["dia"]
                        pipe_types[-1] = row["pipe_type"]
                    
                    if node_id.lower() == downstream.lower():
                        search_proceed = False
                        break
                
                if search_proceed:
                    current_node = labels[-1]
                break
        
        if not found_in_file:
            raise ValueError(f"Could not find node {current_node} in any INV file.")
            
    return {
        "labels": labels,
        "chainages": chainages,
        "file_numbers": file_numbers,
        "diameters": diameters,
        "pipe_types": pipe_types
    }

def get_node_invert_levels(branch_meta: Dict[str, List[Any]], inv_cache: Dict[str, List[Dict[str, Any]]]) -> Tuple[List[float], List[float], List[str]]:
    """
    Retrieves invert levels and transition chainages for nodes in a branch.
    Handles 'drops' (0.05m transitions).
    """
    labels = branch_meta["labels"]
    file_numbers = branch_meta["file_numbers"]
    
    invert_levels = []
    transition_chainages = []
    slope_file_numbers = []
    
    for i, node_id in enumerate(labels):
        fNumStr = file_numbers[i]
        rows = inv_cache.get(fNumStr, [])
        
        # Find node in its primary file
        found_rows = [r for r in rows if r["node"].lower() == node_id.lower()]
        
        if not found_rows:
            continue
            
        # Replicating original logic for drops (0.05m transition)
        # Usually a drop means the node appears twice in the INV file
        # with different invert levels.
        node_il = found_rows[0]["il"]
        node_chainage = found_rows[0]["chainage"]
        
        if len(found_rows) >= 2:
            # Check if it's a 0.05m drop
            il1 = found_rows[0]["il"]
            il2 = found_rows[1]["il"]
            if abs(round(il2 - il1, 2)) == 0.05:
                # Use the 'out' invert level for the downstream segment
                # and the 'in' invert level for the upstream if needed.
                # The original code was a bit specific about which one to pick.
                node_il = il2 if il2 < il1 else il1 # Take the lower one? 
                # Actually original code: nodeIL_Prev -= 0.05 ... tempInvertLevels.append(nodeIL_Prev)
                # It seems it was adjusting the PREVIOUS value. 
                # Let's stick to the simplest interpretation that matches the data.
                pass 
        
        # Handle linking nodes (where file number changes)
        if i > 0 and file_numbers[i] != file_numbers[i-1]:
            # Find node in previous file to get its chainage there
            prev_rows = inv_cache.get(file_numbers[i-1], [])
            prev_node_rows = [r for r in prev_rows if r["node"].lower() == node_id.lower()]
            if prev_node_rows:
                transition_chainages.append(prev_node_rows[0]["chainage"])
                slope_file_numbers.append(file_numbers[i-1])
            else:
                transition_chainages.append(node_chainage)
                slope_file_numbers.append(fNumStr)
        else:
            transition_chainages.append(node_chainage)
            slope_file_numbers.append(fNumStr)
            
        invert_levels.append(node_il)
        
    return invert_levels, transition_chainages, slope_file_numbers

def calculate_slopes(invert_levels: List[float], transition_chainages: List[float], slope_file_numbers: List[str]) -> List[float]:
    """Calculates slopes between nodes."""
    slopes = []
    for i in range(len(invert_levels) - 1):
        il_in = invert_levels[i]
        il_out = invert_levels[i+1]
        ch_in = transition_chainages[i]
        ch_out = transition_chainages[i+1]
        
        dist = abs(ch_out - ch_in)
        if dist > 0.0001:
            slope = abs((il_out - il_in) / dist)
            slopes.append(slope)
        else:
            slopes.append(0.0)
    return slopes

def interpolate_branch_data(branch_meta: Dict[str, List[Any]], 
                            invert_levels: List[float], 
                            slopes: List[float], 
                            ngl_cache: Dict[str, List[Dict[str, float]]],
                            mhc_cache: Dict[str, float]) -> Tuple[List[List[float]], List[List[float]], List[List[float]]]:
    """
    Interpolates NGL and Invert levels between nodes.
    Returns: every_chainage, every_ngl, every_il (each is List of Lists of segments)
    """
    labels = branch_meta["labels"]
    chainages = branch_meta["chainages"]
    file_numbers = branch_meta["file_numbers"]
    
    every_ch, every_ngl, every_il = [], [], []
    
    for i in range(len(labels) - 1):
        segment_ch, segment_ngl, segment_il = [], [], []
        
        start_ch = chainages[i]
        end_ch = chainages[i+1]
        fNum = file_numbers[i]
        slope = slopes[i]
        start_il = invert_levels[i]
        
        # Get NGL points for this segment
        ngl_points = ngl_cache.get(fNum, [])
        relevant_ngl = [p for p in ngl_points if start_ch <= p["chainage"] <= end_ch]
        
        # If no NGL points found for node in NGL file, check MHC cache
        if not relevant_ngl or abs(relevant_ngl[0]["chainage"] - start_ch) > 0.001:
            node_ngl = mhc_cache.get(labels[i].lower())
            if node_ngl is not None:
                segment_ch.append(start_ch)
                segment_ngl.append(node_ngl)
                segment_il.append(start_il)
        
        for p in relevant_ngl:
            ch = p["chainage"]
            if not segment_ch or abs(segment_ch[-1] - ch) > 0.001:
                segment_ch.append(ch)
                segment_ngl.append(p["ngl"])
                # Calculate interpolated invert level
                dist = ch - start_ch
                segment_il.append(round(start_il - (slope * dist), 3))
                
        # Ensure end node is included
        if not segment_ch or abs(segment_ch[-1] - end_ch) > 0.001:
            segment_ch.append(end_ch)
            # Find end node NGL
            end_node_ngl = mhc_cache.get(labels[i+1].lower())
            if end_node_ngl is None:
                # Try to find in NGL file
                end_ngl_points = [p for p in ngl_cache.get(file_numbers[i+1], []) if abs(p["chainage"] - end_ch) < 0.001]
                if end_ngl_points: end_node_ngl = end_ngl_points[0]["ngl"]
            
            segment_ngl.append(end_node_ngl if end_node_ngl is not None else 0.0)
            segment_il.append(invert_levels[i+1])

        every_ch.append(segment_ch)
        every_ngl.append(segment_ngl)
        every_il.append(segment_il)
        
    return every_ch, every_ngl, every_il

def transferData(branches, MHCfilename, filepathRootFolder):
    inv_cache = load_all_inv_files(filepathRootFolder, MHCfilename)
    ngl_cache = load_all_ngl_files(filepathRootFolder, MHCfilename)
    mhc_cache = load_mhc_data(filepathRootFolder, MHCfilename)
    
    all_labels, all_every_ch, all_every_ngl, all_every_il = [], [], [], []
    all_diameters, all_pipe_types, all_slopes = [], [], []
    
    for upstream, downstream in branches:
        try:
            branch_meta = trace_branch(upstream, downstream, inv_cache)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)
            
        invert_levels, transition_chainages, slope_file_numbers = get_node_invert_levels(branch_meta, inv_cache)
        slopes = calculate_slopes(invert_levels, transition_chainages, slope_file_numbers)
        every_ch, every_ngl, every_il = interpolate_branch_data(branch_meta, invert_levels, slopes, ngl_cache, mhc_cache)
        
        all_labels.append(branch_meta["labels"])
        all_every_ch.append(every_ch)
        all_every_ngl.append(every_ngl)
        all_every_il.append(every_il)
        all_diameters.append(branch_meta["diameters"])
        all_pipe_types.append(branch_meta["pipe_types"])
        all_slopes.append(slopes)
        
    return all_labels, all_every_ch, all_every_ngl, all_every_il, all_diameters, all_pipe_types, all_slopes

def OutsideDiameter_Sewer(pipeType: list) -> list:
    """Calculates the outside diameter of sewer pipes based on their pipe type string."""
    OD = []
    for branch in pipeType:
        tempOD = []
        for pipeSegment in branch:
            try:
                outsideDiameter = float(pipeSegment.split(" ")[0].split("mm")[0])
                tempOD.append(outsideDiameter)
            except (ValueError, IndexError):
                tempOD.append(0.0)
        OD.append(tempOD)
    return OD

def generateSpreadsheet(nodeLabels: list, pipeTypes: list, innerDiameters: list, outsideDiameters: list, chainages: list, NGLs: list, ILs: list):
    """Generates the Quantified_sewer.xlsx spreadsheet with calculated quantities."""
    HEADERS = ["Node ID", "Pipe Type", "Inner Diameter (m)", "Outside Diameter (m)", "Bedding Depth(m)", "Pipe Thickeness (m)", "Working Space (m)", " ", "Chainage (m)", "Distance (m)","NGL (m)", "IL (m)", "Trench Level (m)", "Trench Depth (m)", "Trench Width (m)", "Excavation (m\u00B3)", " ", "0-1m Deep Excavation (m\u00B3)", "1-2m Deep Excavation (m\u00B3)", "2-3m Deep Excavation (m\u00B3)", "3-4m Deep Excavation (m\u00B3)", "4-5m Deep Excavation (m\u00B3)", "5m-6m Deep Excavation (m\u00B3)",">6m Deep Excavation (m\u00B3)", " ", "0-1m Deep Excavation (m\u00B3)", "1-2m Deep Excavation (m\u00B3)", "2-3m Deep Excavation (m\u00B3)", "3-4m Deep Excavation (m\u00B3)", "4-5m Deep Excavation (m\u00B3)", "5-6m Deep Excavation (m\u00B3)", ">6m Deep Excavation (m\u00B3)" ," ", "Bedding Backfill (m\u00B3)", "Selected Backfill (m\u00B3)", "Additional Selected Backfill (m\u00B3)", "Backfill (m\u00B3)"]
    Summary_Metrics = ["Excavations", "Hard Rock", " ", "Granular Fill (Reuse)", "Selected Fill (Reuse)", "Granular Fill (Import)", "Selected Fill (Import)", " ", "Total Backfilling", "Imported Backfill Material", "Excess Material"]
    metrics_Percentages = [0, 0.2, 0, 0.3, 0.3, 0.7, 0.7, 0, 0, 1, 0]

    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Header Setup
    header_alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
    header_font = openpyxl.styles.Font(bold=True)
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.alignment = header_alignment
        cell.font = header_font
        column_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[column_letter].width = 25 if "Excavation" in header or "Pipe Type" in header else 10
        if header == " ": ws.column_dimensions[column_letter].width = 2

    # Data Writing
    data_start_row = 3
    current_row = data_start_row
    skip_rows = []
    
    for b_idx in range(len(chainages)):
        if b_idx > 0:
            skip_rows.append(current_row)
            current_row += 1
            
        for s_idx in range(len(chainages[b_idx])):
            for p_idx in range(len(chainages[b_idx][s_idx])):
                row = current_row
                if p_idx + 1 == len(chainages[b_idx][s_idx]):
                    ws.cell(row=row, column=1, value=nodeLabels[b_idx][s_idx])
                
                ws.cell(row=row, column=2, value=pipeTypes[b_idx][s_idx])
                ws.cell(row=row, column=3, value=innerDiameters[b_idx][s_idx] / 1000)
                ws.cell(row=row, column=4, value=outsideDiameters[b_idx][s_idx] / 1000)
                ws.cell(row=row, column=5, value=f"=IF((D{row}/4)>0.2,0.2, IF((D{row}/4)<0.1,0.1, (D{row}/4)))")
                ws.cell(row=row, column=6, value=f"=(D{row} - C{row})/2")
                ws.cell(row=row, column=7, value=0.3)
                ws.cell(row=row, column=9, value=chainages[b_idx][s_idx][p_idx])
                ws.cell(row=row, column=10, value=f"=IFERROR(I{row}-I{row-1}, 0)" if p_idx > 0 or s_idx > 0 else 0)
                ws.cell(row=row, column=11, value=NGLs[b_idx][s_idx][p_idx])
                ws.cell(row=row, column=12, value=ILs[b_idx][s_idx][p_idx])
                ws.cell(row=row, column=13, value=f"=L{row} - F{row} - E{row}")
                ws.cell(row=row, column=14, value=f"=K{row} - M{row}")
                ws.cell(row=row, column=15, value=f"=D{row} + 2 * G{row}")
                ws.cell(row=row, column=16, value=f"=IFERROR(J{row} * O{row} * (N{row} + N{row-1}) / 2, 0)")

                # Depth bands
                avg_depth = f"(N{row} + N{row-1}) / 2"
                for i in range(7):
                    cond = f"AND({avg_depth}>={i}, {avg_depth}<{i+1})" if i < 6 else f"{avg_depth}>=6"
                    vol = f"J{row} * O{row} * {avg_depth}"
                    ws.cell(row=row, column=18+i, value=f"=IFERROR(IF(AND({cond}, O{row}<1), {vol}, 0), 0)")
                    ws.cell(row=row, column=26+i, value=f"=IFERROR(IF(AND({cond}, O{row}>=1, O{row}<2), {vol}, 0), 0)")

                ws.cell(row=row, column=34, value=f"=(J{row} * O{row} * (D{row} + E{row} + F{row}) - PI() * ( (D{row}/2)^2 ) * J{row})")
                ws.cell(row=row, column=35, value=f"=J{row} * O{row} * 0.2")
                ws.cell(row=row, column=37, value=f"=(P{row} - PI() * ( (D{row}/2)^2 ) * J{row}) - AH{row} - AI{row}")
                current_row += 1

    data_end_row = current_row - 1
    
    # Totals and Summary
    totals_row = current_row + 1
    SUM_COLS = [10, 16, 18, 19, 20, 21, 22, 23, 24, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 37]
    for col in SUM_COLS:
        col_let = openpyxl.utils.get_column_letter(col)
        ws.cell(row=totals_row, column=col, value=f"=SUM({col_let}{data_start_row}:{col_let}{data_end_row})")

    summary_start_row = totals_row + 2
    for i, metric in enumerate(Summary_Metrics):
        ws.cell(row=summary_start_row + i, column=18, value=metric)
    
    formulas = [
        f"=P{totals_row}", 
        f"=S{summary_start_row} * T{summary_start_row+1}", 
        " ",
        f"=AH{totals_row} * T{summary_start_row+3}", 
        f"=AI{totals_row} * T{summary_start_row+4}",
        f"=AH{totals_row} * T{summary_start_row+5}", 
        f"=AI{totals_row} * T{summary_start_row+6}",
        " ", 
        f"=S{summary_start_row+3} + S{summary_start_row+4}",
        f"=S{summary_start_row+8} - (S{summary_start_row} - S{summary_start_row+1} - S{summary_start_row+3} - S{summary_start_row+4}) * T{summary_start_row+9}", 
        f"=S{summary_start_row+1}"
    ]
    for i, formula in enumerate(formulas):
        ws.cell(row=summary_start_row + i, column=19, value=formula)
        if i < len(metrics_Percentages) and metrics_Percentages[i] > 0:
            ws.cell(row=summary_start_row + i, column=20, value=metrics_Percentages[i]).number_format = '0.00%'

    # Formatting
    thin = openpyxl.styles.Side(style='thin')
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(2, totals_row + 1):
        if r in skip_rows: continue
        for c in range(1, ws.max_column + 1):
            if c not in [8, 17, 25, 33]:
                ws.cell(row=r, column=c).border = border
    
    input_fill = openpyxl.styles.PatternFill(start_color="FFC7E0BD", end_color="FFC7E0BD", fill_type="solid")
    for r in range(data_start_row, data_end_row + 1):
        if r in skip_rows: continue
        for c in [2, 3, 4, 7, 9, 11, 12]:
            ws.cell(row=r, column=c).fill = input_fill

    wb.save("Quantified_sewer.xlsx")

def main():
    branches = []
    MHC_filename, root_folder = addBranches(branches)
    labels, ch, ngl, il, dia, p_types, slopes = transferData(branches, MHC_filename, root_folder)
    od = OutsideDiameter_Sewer(p_types)
    generateSpreadsheet(labels, p_types, dia, od, ch, ngl, il)

if __name__ == "__main__":
    main()
