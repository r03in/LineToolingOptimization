# CLAUDE.md — LineToolingOptimization

## Project Overview

Models production line allocation for a manufacturing facility producing multiple products
on shared lines. Optimizes: (1) how many lines per year, (2) which products on which lines,
(3) volume per line, (4) total tooling investment (mechanical + optical sets).
Goal: minimize tooling while meeting demand and respecting time-based capacity constraints.

## Current Status

The project has two solvers:

| Solver | Algorithm | Tooling sets | Notes |
|--------|-----------|-------------|-------|
| `solver_greedy.py` | Greedy heuristic (year-by-year) | 38 sets | Original; preserved as reference |
| `solver_ilp.py` | Mixed-Integer Linear Program (CBC) | Target: ~28 sets | **Active solver** |

The ILP solver is the primary tool going forward. It closes the 10-set gap over the greedy
approach by optimising all years simultaneously and exploiting tooling-sharing matrices.

## Inputs (Blad1 sheet)

- Volume Demand: A1, rows 3-19, cols B-L (years 2025-2041, 10 products)
- Production Setup: A21, B22-B27 (hours/shift=7.2, shifts=3, days=6, weeks=48 -> 22,394,880 sec/yr)
- Cycle Time: A29, rows 31-45, cols B-L (15 lines x 10 products, default 12s)
- OEE: A47, rows 49-63, cols B-L (15 lines x 10 products, default 85%)
- Mechanical Tooling Matrix: A67, rows 70-79, cols C-L (10x10, currently identity)
- Optical Tooling Matrix: A81, rows 84-93, cols C-L (10x10, currently identity)

## Output Grid

- Location: A95, data rows 98-114, cols C-EV (150 columns = 15 lines x 10 products)
- Column mapping: Line 1 = C-L, Line 2 = M-V, Line 3 = W-AF, Line 4 = AG-AP, ...
- Each cell = units allocated to that product on that line for that year
- Validation: rows 116-135 with sum formulas and OK/MISMATCH checks

## Key Constraints

HARD: (1) Demand satisfaction, (2) Time capacity: SUM(units*cycle/OEE) <= avail_seconds per line,
(3) Line 1 = validation line: MUST run ALL active products every year.
SOFT (priority order): (1) Minimize lines, (2) Minimize tooling sets, (3) Minimize mixing.

## Algorithm (Planned-Role Approach)

### Line 1 (Validation)

- If total demand fits on 1 line -> put everything on Line 1
- If small products (all except dominant) fit on Line 1 -> put ALL small products first,
  fill remaining time with dominant product (avoids overflow lines for small products)
- Otherwise -> proportional fill across all products

### Lines 2+ (Overflow)

- Phase A: Reuse lines already tooled for the product (score +1M, no new tooling)
- Phase B: Pack onto already-open lines (score +500K, avoids new line)
- Scoring: +1M if already tooled, +500K if line already open, -line_number (prefer lower lines)

### Look-ahead Tooling

During a line's intro year, check next 3 years. If a product will be needed, pre-add
tooling (cheaper during initial supplier validation).

## Allocation Results

### Theoretical Optimum (from Summary sheet — hand-computed / Excel Solver)

```
Line | Role              | Products    | Intro | Tooling
L1   | Validation        | P1,P2,P3,P4 | 2029  | 4 mech + 4 opt
L2   | P2 dedicated      | P2           | 2031  | 1 mech + 1 opt
L3   | Small-prod shared  | P1,P3,P4    | 2032  | 3 mech + 3 opt
L4   | P2 dedicated      | P2           | 2033  | 1 mech + 1 opt
L5   | Small-prod shared  | P1,P3,P4    | 2034  | 3 mech + 3 opt
L6   | P2+P3 overflow    | P2,P3        | 2035  | 2 mech + 2 opt
TOTAL:                                           14 mech + 14 opt = 28 sets
```

Line count by year: 1->1->2->3->4->5->6->6->3->1 (all theory-minimum)

### Greedy Solver Output (solver.py — actual output)

```
Line | Products         | Intro | Tooling
L1   | P1,P2,P3,P4      | 2029  | 4 mech + 4 opt
L2   | P1,P2,P3         | 2031  | 3 mech + 3 opt
L3   | P1,P2,P3,P4      | 2032  | 4 mech + 4 opt
L4   | P1,P2,P3,P4      | 2033  | 4 mech + 4 opt
L5   | P1,P3            | 2034  | 2 mech + 2 opt
L6   | P3,P4            | 2035  | 2 mech + 2 opt
TOTAL:                            19 mech + 19 opt = 38 sets
```

The greedy solver is 10 sets above the theoretical minimum. The gap is primarily because
the greedy approach does not keep overflow lines product-dedicated — it mixes products
across lines, requiring more tooling sets per line.

## Key Design Decisions

1. **TIME-BASED CAPACITY**: units * cycle_time / OEE, not flat unit caps.
   Max ~1,586,304 units/yr for single product at 12s/85%.

2. **LINE 1 PRIORITIZES SMALL PRODUCTS**: When P1+P3+P4 fit on L1, give them ALL their
   demand first. This prevents small products needing their own overflow lines.

3. **DEDICATED OVERFLOW LINES**: P2 (dominant ~4x others) gets dedicated lines.
   Only mix at peak when small products slightly overflow L3+L5 capacity.

4. **3-YEAR LOOK-AHEAD**: Pre-add tooling during intro year for products needed within 3 years.

5. **TOOLING MATRICES**: 10x10 compatibility. Currently identity (no sharing).
   When updated, products sharing tooling reduce total sets needed.

## File Structure

```
LineToolingOptimization/
├── Book.xlsx             — Excel workbook: all inputs, output grid, validation, summary
├── solver_ilp.py         — ILP solver (primary) — reads Book.xlsx, writes 3 outputs
├── solver_greedy.py      — Original greedy solver (reference / fallback)
├── solver.py             — Alias kept for compatibility (same as solver_greedy.py)
├── requirements.txt      — Python dependencies: pulp, matplotlib, openpyxl
└── CLAUDE.md             — This file
```

Generated outputs (written to the same directory as the workbook):

```
tooling_summary.csv   — Allocation per line/year + physical tooling ID table
line_gantt.png        — Gantt chart: products per line across all years
Book.xlsx             — Updated with allocation grid (Blad1 rows 98+)
```

## How to Run (ILP Solver)

```bash
# 1. Install dependencies (once)
pip install -r requirements.txt

# 2. Run the ILP solver
python solver_ilp.py Book.xlsx

# 3. Check console output for tooling count and per-year summary
# 4. Open tooling_summary.csv for detailed line/year allocation + tooling IDs
# 5. Open line_gantt.png for visual overview
# 6. Open Book.xlsx → Blad1 validation rows 119-135 (all should show OK)
```

## ILP Solver — Key Features

### Optimisation Objective (priority order)
1. **Minimise total lines** open across all years (`W_LINES = 10,000`)
2. **Minimise tooling sets** — one set per tooling *family* per line, not per product (`W_TOOLING = 100`)
3. **Minimise product-line switches** — penalises moving a product to a different line year-over-year (`W_SWITCHES = 1`)

### Tooling Sharing
Reads the 10×10 mechanical and optical compatibility matrices from Blad1. Computes connected
components (union-find) — products in the same family share one physical set on a line.

- Current matrices are **identity** (no sharing) — each product = its own family
- To enable sharing: set `mech_matrix[i][j] = mech_matrix[j][i] = 1` for compatible product pairs
- The tooling count drops automatically; no solver changes needed

### Changeover OEE Penalty
Lines running 2+ products incur a flat OEE reduction (configurable):

```python
BASE_OEE               = 0.85   # single-product line OEE
CHANGEOVER_OEE_PENALTY = 0.03   # reduction for multi-product lines → effective 82%
```

Evaluated per line per year. A line that consolidates to a single product reverts to BASE_OEE.
Implemented as a linear constraint: `sum(x*ct) + avail*PENALTY*multi <= avail*BASE_OEE`.

### Physical Tooling ID Tracking
Each tooling set is named by product family and type:
- `MECH-P01` — mechanical set for the family containing P1
- `OPTI-P02` — optical set for the family containing P2

The `tooling_summary.csv` shows which line each ID sits on and the years it is active,
making it easy to track tooling movement between lines across years.

## Configurable Parameters (top of solver_ilp.py)

| Parameter | Default | Effect |
|-----------|---------|--------|
| `BASE_OEE` | 0.85 | OEE for single-product lines |
| `CHANGEOVER_OEE_PENALTY` | 0.03 | OEE reduction for multi-product lines |
| `W_LINES` | 10,000 | Objective weight: minimise open lines |
| `W_TOOLING` | 100 | Objective weight: minimise tooling sets |
| `W_SWITCHES` | 1 | Objective weight: minimise product-line changes |
| `SOLVER_TIME_LIMIT` | 300 | CBC solver wall-clock limit (seconds) |
| `NUM_LINES` | 15 | Maximum lines available |
| `NUM_PRODUCTS` | 10 | Maximum products supported |

## What Could Change

- **New products P5-P10**: add demand rows to Blad1; solver handles up to 10 automatically
- **Different cycle times**: update rows 31-45 in Blad1
- **Tooling sharing**: update the 10×10 matrices in Blad1 (rows 70-79 mech, 84-93 optical)
- **More/fewer lines**: change `NUM_LINES` in solver_ilp.py
- **Tighter time budget**: reduce `SOLVER_TIME_LIMIT` for faster (possibly suboptimal) results

## Known Limitations

- ILP may hit the time limit on large instances; increase `SOLVER_TIME_LIMIT` if needed
- No multi-year demand smoothing; each year is an independent constraint block
- No minimum lot size or campaign-length constraints
- No changeover *time* (only OEE penalty); add a changeover matrix if needed later
- Tooling movement between lines has no cost penalty currently

## Demand Summary (2025-2041, active products P1-P4 only)

| Year | P1      | P2        | P3      | P4      | Total     | Lines |
|------|---------|-----------|---------|---------|-----------|-------|
| 2029 | 1,000   | 1,000     | 1,000   | 1,000   | 4,000     | 1     |
| 2030 | 35,498  | 141,992   | 35,498  | 35,498  | 248,486   | 1     |
| 2031 | 372,560 | 1,490,238 | 372,560 | 372,560 | 2,607,918 | 2     |
| 2032 | 529,763 | 2,119,050 | 529,763 | 529,763 | 3,708,339 | 3     |
| 2033 | 693,177 | 2,772,708 | 693,177 | 693,177 | 4,852,239 | 4     |
| 2034 | 1,001,694| 4,006,776| 1,001,694| 1,001,694| 7,011,858| 5    |
| 2035 | 1,338,697| 5,354,788| 1,338,697| 1,338,697| 9,370,879| 6    |
| 2036 | 1,286,832| 5,417,328| 1,286,832| 1,286,832| 9,277,824| 6    |
| 2037 | 610,237 | 2,440,948 | 610,237 | 610,237 | 4,271,659 | 3     |
| 2038 | 173,752 | 695,008   | 173,752 | 173,752 | 1,216,264 | 1     |
