# Quantify Utility: Engineering Rules & Calculations

This document outlines the technical rules, data parsing logic, and engineering calculations used in the **Quantify** utility to generate sewer network quantity takeoffs.

---

## 1. Data Parsing Rules

The utility consumes three primary file types exported from design software:

*   **INV (.inv):** Contains manhole/node invert levels (IL), cumulative chainages, pipe diameters, and pipe types.
*   **NGL (.ngl):** Contains high-density ground profile points (Chainage and Level) along the pipe route.
*   **MHC (.mhc):** Contains specific ground levels (NGL) assigned to named nodes (manholes).

### A. Backdrop (Vertical Drop) Detection
*   **Rule:** A backdrop is identified when two consecutive rows in an INV file share the **exact same chainage**.
*   **Logic:**
    *   Row 1: Represents the **Inlet** side of the manhole (`IL_in`).
    *   Row 2 (Duplicate Chainage): Represents the **Outlet** side or the drop magnitude.
    *   **Calculation:** If the value in Row 2 is a "sentinel" (typically small values like `-0.05` or `-0.025`), the `IL_out` is calculated as `Row 1 IL + Row 2 Value`. If the value is large, it is treated as an absolute elevation for the outlet side.
*   **Constraint:** Detection is **boundless** and magnitude-independent.

### B. Chainage Alignment (Snapping)
*   **Rule:** To resolve floating-point discrepancies between INV and NGL files (e.g., `488.424` vs `488.425`), the script "snaps" node chainages.
*   **Logic:** If a node chainage in the INV file is within **0.002m** of a point in the NGL file, the NGL chainage is used. This ensures the ground profile and pipe profile are perfectly synchronized.

---

## 2. Core Engineering Logic

### A. Segment Definition
A "Segment" is defined as the pipe length between two physical points (either two manholes or a manhole and an intermediate NGL point).

### B. Slope Calculation
The slope ($S$) for a pipe segment between Node A and Node B is calculated as:
$$S = \frac{IL_{out\_A} - IL_{in\_B}}{Distance_{A 	o B}}$$
*Where $IL_{out\_A}$ is the outlet level of the upstream node (after any backdrop).*

### C. Invert Level Interpolation
For any intermediate point $P$ at distance $d$ from the start of the segment:
$$IL_P = IL_{start} - (S 	imes d)$$
*This ensures a constant grade across the entire length of a single pipe.*

---

## 3. Spreadsheet Calculations (Excel Formulas)

The following formulas are applied to every row in `Quantified_sewer.xlsx`:

| Column | Metric | Formula / Logic |
| :--- | :--- | :--- |
| **E** | **Bedding Depth** | `=IF((OD/4)>0.2, 0.2, IF((OD/4)<0.1, 0.1, (OD/4)))` <br> *(Bounds: 100mm min, 200mm max)* |
| **F** | **Pipe Thickness** | `(Outside Diameter - Inner Diameter) / 2` |
| **G** | **Working Space** | Constant: `0.300m` (applied to both sides) |
| **J** | **Distance** | `Chainage_current - Chainage_previous` |
| **M** | **Trench Level** | `Invert Level - Pipe Thickness - Bedding Depth` |
| **N** | **Trench Depth** | `NGL - Trench Level` |
| **O** | **Trench Width** | `Outside Diameter + (2 * Working Space)` |
| **P** | **Excavation Vol** | `Distance * Trench Width * Average(Depth_current, Depth_prev)` |

---

## 4. Categorization Rules (Depth Bands)

Excavation volumes are automatically sorted into columns based on **Average Depth** and **Trench Width** to support standard Bill of Quantities (BOQ) formats:

*   **Group 1 (Width < 1.0m):** Sorted into bands: `0-1m`, `1-2m`, `2-3m`, `3-4m`, `4-5m`, `5-6m`, and `>6m`.
*   **Group 2 (Width 1.0m to 2.0m):** Sorted into identical depth bands in separate columns.

**Condition Example (0-1m Band):**
`IF(AND(Avg_Depth >= 0, Avg_Depth < 1, Trench_Width < 1), Volume, 0)`

---

## 5. Backfilling Rules

*   **Bedding Backfill (Col AH):** Calculates the volume of bedding material required, subtracting the volume displaced by the pipe.
    $$Vol = (Distance 	imes Width 	imes (OD + BeddingDepth + Thickness)) - (\pi 	imes Radius^2 	imes Distance)$$
*   **Selected Backfill (Col AI):** Volume of selected material above the bedding (fixed at 200mm depth by default).
*   **General Backfill (Col AK):** The remaining trench volume to be backfilled with shared/excavated material.
    $$Vol = (Total Excavation - Pipe Displacement) - BeddingBackfill - SelectedBackfill$$

---

## 6. Summary Section Calculations

Located at the bottom of the spreadsheet, these metrics provide the final totals for the entire network:

1.  **Hard Rock:** `Total Excavation * 20%` (Default provision).
2.  **Granular Fill (Reuse):** `Total Bedding Backfill * 30%`.
3.  **Imported Material:** `(Total Bedding + Selected Backfill) - Reuse Volumes`.
4.  **Excess Material:** `Total Excavation - (Backfill used - Import Provision)`.
