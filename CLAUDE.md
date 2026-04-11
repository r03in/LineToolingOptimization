# CLAUDE.md — LineToolingOptimization

## Project Overview

Models production line allocation for a manufacturing facility producing multiple products
on shared lines. Optimizes: (1) how many lines per year, (2) which products on which lines,
(3) volume per line, (4) total tooling investment (mechanical + optical sets).
Goal: minimize tooling while meeting demand and respecting time-based capacity constraints.

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
├── Book.xlsx           — Excel workbook with all input data, output grid, and summary
├── solver.py           — Python optimizer (greedy heuristic, reads/writes Book.xlsx)
└── CLAUDE.md           — This file (project documentation and algorithm reference)
```

## How to Re-run

1. Modify demand/cycle times/OEE/tooling on Blad1
2. Run: `python solver.py Book.xlsx`
3. Check validation rows 119-135 on Blad1 (all should show OK)
4. Review Summary sheet for updated line count, utilization, tooling

## What Could Change

- New products P5-P10: just add demand data, solver handles up to 10
- Different cycle times per product/line: update the matrix rows 31-45
- Tooling sharing: change identity matrices to reflect product families
- More/fewer lines: change `NUM_LINES` in solver.py (currently 15)
- Different look-ahead: change `LOOK_AHEAD_YEARS` (currently 3)

## Known Limitations

- Greedy heuristic, not global optimizer (early year decisions affect later)
- Tooling sharing matrices read but not fully exploited in scoring yet
- No changeover time modeling between products on same line
- No multi-year demand smoothing; each year is solved independently

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
