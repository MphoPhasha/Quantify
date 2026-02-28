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
                metaList.append(fNumStr)
            break
        elif userInput == "3":
            fNum = input("Enter branch file number (e.g. 001): ").strip()
            if fNum: metaList.append(fNum)
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
                    line = line.strip()
                    if not line: continue
                    row = line.split(",")
                    if len(row) < 7: row = line.split()
                    if len(row) >= 7:
                        try:
                            data.append({
                                "chainage": float(row[0]),
                                "il": float(row[1]),
                                "node": labelFormat(row[2]),
                                "dia": float(row[3]),
                                "pipe_type": pipeTypeFormat(row[6])
                            })
                        except (ValueError, IndexError): continue
                inv_cache[fNumStr] = data
        except FileNotFoundError: continue
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
                    line = line.strip()
                    if not line: continue
                    row = line.split()
                    if len(row) < 2: row = line.split(",")
                    if len(row) >= 2:
                        try:
                            data.append({"chainage": float(row[0]), "ngl": float(row[1])})
                        except (ValueError, IndexError): continue
                ngl_cache[fNumStr] = data
        except FileNotFoundError: continue
    return ngl_cache

def get_ngl_at_chainage(chainage: float, ngl_pts: List[Dict[str, float]]) -> float:
    """Linearly interpolates NGL level at a given chainage from a list of NGL points."""
    if not ngl_pts:
        return 0.0
    pts = sorted(ngl_pts, key=lambda x: x["chainage"])
    if chainage <= pts[0]["chainage"]: return pts[0]["ngl"]
    if chainage >= pts[-1]["chainage"]: return pts[-1]["ngl"]
    for i in range(len(pts) - 1):
        p1, p2 = pts[i], pts[i+1]
        if p1["chainage"] <= chainage <= p2["chainage"]:
            dx = p2["chainage"] - p1["chainage"]
            if dx < 0.0001: return p1["ngl"]
            return p1["ngl"] + (chainage - p1["chainage"]) * (p2["ngl"] - p1["ngl"]) / dx
    return pts[-1]["ngl"]

def transferData(branches, MHCfilename, filepathRootFolder):
    inv_cache = load_all_inv_files(filepathRootFolder, MHCfilename)
    ngl_cache = load_all_ngl_files(filepathRootFolder, MHCfilename)
    mhc_cache = load_mhc_data(filepathRootFolder, MHCfilename)
    
    all_labels, all_every_ch, all_every_ngl, all_every_il = [], [], [], []
    all_diameters, all_pipe_types, all_slopes = [], [], []
    
    for fNumStr in branches:
        inv_rows = inv_cache.get(fNumStr, [])
        ngl_pts = ngl_cache.get(fNumStr, [])
        if not inv_rows or not ngl_pts: continue
        
        # 1. Parse INV rows into physical points with IN and OUT levels
        # Duplicate chainages indicate a backdrop (vertical drop/step)
        points = []
        for row in inv_rows:
            if points and abs(row["chainage"] - points[-1]["chainage"]) < 0.0001:
                # This is a backdrop sentinel for the previous row
                # We update the OUT level of the current manhole position
                # If the value is small (like -0.05), it's a relative drop
                # If the value is large, it might be an absolute level (though rare for drops)
                # Designer software usually exports drops as relative values.
                if -500 < row["il"] < 500: # Relative drop range
                    points[-1]["il_out"] += row["il"]
                else: # Absolute level provided for the 'out' side
                    points[-1]["il_out"] = row["il"]
                continue
            
            points.append({
                "chainage": row["chainage"],
                "il_in": row["il"],
                "il_out": row["il"],
                "node": row["node"],
                "dia": row["dia"],
                "pt": row["pipe_type"]
            })

        every_ch, every_ngl, every_il = [], [], []
        slopes = []
        labels_branch, diameters_branch, pipe_types_branch = [], [], []
        
        # 2. For each segment (between two physical points), collect NGL points and interpolate
        for i in range(len(points) - 1):
            p_start = points[i]
            p_end = points[i+1]
            
            dist = abs(p_end["chainage"] - p_start["chainage"])
            # Slope uses OUT level of start manhole and IN level of end manhole
            slope = (p_start["il_out"] - p_end["il_in"]) / dist if dist > 0.0001 else 0.0
            
            slopes.append(slope)
            labels_branch.append(p_start["node"])
            diameters_branch.append(p_start["dia"])
            pipe_types_branch.append(p_start["pt"])
            
            s_ch, s_ngl, s_il = [], [], []
            
            # Snap start/end chainages to NGL points
            snapped_start = p_start["chainage"]
            snapped_end = p_end["chainage"]
            for p in ngl_pts:
                if abs(p["chainage"] - p_start["chainage"]) <= 0.002: snapped_start = p["chainage"]
                if abs(p["chainage"] - p_end["chainage"]) <= 0.002: snapped_end = p["chainage"]

            # Add points from NGL file that are within segment
            for p in ngl_pts:
                p_ch = p["chainage"]
                if snapped_start - 0.0001 <= p_ch <= snapped_end + 0.0001:
                    s_ch.append(p_ch)
                    s_ngl.append(p["ngl"])
                    s_il.append(round(p_start["il_out"] - slope * (p_ch - snapped_start), 3))
            
            # Ensure boundaries are present
            if not s_ch or abs(s_ch[0] - snapped_start) > 0.001:
                s_ch.insert(0, snapped_start)
                s_ngl.insert(0, mhc_cache.get(p_start["node"].lower(), get_ngl_at_chainage(snapped_start, ngl_pts)))
                s_il.insert(0, p_start["il_out"])
            if abs(s_ch[-1] - snapped_end) > 0.001:
                s_ch.append(snapped_end)
                s_ngl.append(mhc_cache.get(p_end["node"].lower(), get_ngl_at_chainage(snapped_end, ngl_pts)))
                s_il.append(p_end["il_in"])
            
            s_il[-1] = p_end["il_in"]

            every_ch.append(s_ch)
            every_ngl.append(s_ngl)
            every_il.append(s_il)
            
        all_labels.append(labels_branch + [points[-1]["node"]])
        all_every_ch.append(every_ch)
        all_every_ngl.append(every_ngl)
        all_every_il.append(every_il)
        all_diameters.append(diameters_branch)
        all_pipe_types.append(pipe_types_branch)
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
    header_alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
    header_font = openpyxl.styles.Font(bold=True)
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.alignment = header_alignment
        cell.font = header_font
        column_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[column_letter].width = 25 if "Excavation" in header or "Pipe Type" in header else 10
        if header == " ": ws.column_dimensions[column_letter].width = 2

    data_start_row = 3
    current_row = data_start_row
    skip_rows = []
    
    for b_idx in range(len(chainages)):
        if b_idx > 0:
            skip_rows.append(current_row)
            current_row += 1
        branch_labels = nodeLabels[b_idx]
        branch_ch, branch_ngl, branch_il = chainages[b_idx], NGLs[b_idx], ILs[b_idx]
        branch_dia, branch_od, branch_pt = innerDiameters[b_idx], outsideDiameters[b_idx], pipeTypes[b_idx]
        branch_start_row = current_row
        
        for s_idx in range(len(branch_ch)):
            seg_ch, seg_ngl, seg_il = branch_ch[s_idx], branch_ngl[s_idx], branch_il[s_idx]
            for p_idx in range(len(seg_ch)):
                if s_idx > 0 and p_idx == 0: continue
                row = current_row
                if s_idx == 0 and p_idx == 0: ws.cell(row=row, column=1, value=branch_labels[0])
                elif p_idx == len(seg_ch) - 1: ws.cell(row=row, column=1, value=branch_labels[s_idx + 1])
                
                ws.cell(row=row, column=2, value=branch_pt[s_idx])
                ws.cell(row=row, column=3, value=branch_dia[s_idx] / 1000)
                ws.cell(row=row, column=4, value=branch_od[s_idx] / 1000)
                ws.cell(row=row, column=5, value=f"=IF((D{row}/4)>0.2,0.2, IF((D{row}/4)<0.1,0.1, (D{row}/4)))")
                ws.cell(row=row, column=6, value=f"=(D{row} - C{row})/2")
                ws.cell(row=row, column=7, value=0.3)
                ws.cell(row=row, column=9, value=seg_ch[p_idx])
                
                if row == branch_start_row: ws.cell(row=row, column=10, value=0)
                else: ws.cell(row=row, column=10, value=f"=IFERROR(I{row}-I{row-1}, 0)")
                    
                ws.cell(row=row, column=11, value=seg_ngl[p_idx])
                ws.cell(row=row, column=12, value=seg_il[p_idx])
                ws.cell(row=row, column=13, value=f"=L{row} - F{row} - E{row}")
                ws.cell(row=row, column=14, value=f"=K{row} - M{row}")
                ws.cell(row=row, column=15, value=f"=D{row} + 2 * G{row}")
                ws.cell(row=row, column=16, value=f"=IFERROR(J{row} * O{row} * (N{row} + N{row-1}) / 2, 0)")

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
    totals_row = current_row + 1
    SUM_COLS = [10, 16, 18, 19, 20, 21, 22, 23, 24, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 37]
    for col in SUM_COLS:
        col_let = openpyxl.utils.get_column_letter(col)
        ws.cell(row=totals_row, column=col, value=f"=SUM({col_let}{data_start_row}:{col_let}{data_end_row})")

    summary_start_row = totals_row + 2
    for i, metric in enumerate(Summary_Metrics):
        ws.cell(row=summary_start_row + i, column=18, value=metric)
    
    formulas = [
        f"=P{totals_row}", f"=S{summary_start_row} * T{summary_start_row+1}", " ",
        f"=AH{totals_row} * T{summary_start_row+3}", f"=AI{totals_row} * T{summary_start_row+4}",
        f"=AH{totals_row} * T{summary_start_row+5}", f"=AI{totals_row} * T{summary_start_row+6}",
        " ", f"=S{summary_start_row+3} + S{summary_start_row+4}",
        f"=S{summary_start_row+8} - (S{summary_start_row} - S{summary_start_row+1} - S{summary_start_row+3} - S{summary_start_row+4}) * T{summary_start_row+9}", 
        f"=S{summary_start_row+1}"
    ]
    for i, formula in enumerate(formulas):
        ws.cell(row=summary_start_row + i, column=19, value=formula)
        if i < len(metrics_Percentages) and metrics_Percentages[i] > 0:
            ws.cell(row=summary_start_row + i, column=20, value=metrics_Percentages[i]).number_format = '0.00%'

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
