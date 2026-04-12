# CLAUDE.md — LineToolingOptimization

## Project Overview

Models production line allocation for a manufacturing facility producing multiple products
on shared lines. Core questions: (1) how many lines per year, (2) which products on which
lines, (3) volume per line. Secondary question (solver_ilp only): total tooling investment.
Goal: meet demand at minimum cost while respecting time-based capacity and validation constraints.

---

## Three Solvers — Summary

| Solver | Workbook | Algorithm | Tracks tooling | Time limit | Current result |
|--------|----------|-----------|---------------|------------|---------------|
| `solver_greedy.py` | `Solver.xlsx` | Year-by-year greedy | Yes (mech + opt) | N/A | 38 sets, reference only |
| `solver_ilp.py` | `Solver.xlsx` | Global MIP (all years) | Yes (mech + opt) | 600 s | 7 lines, 24 sets, $25.96M |
| `solver_simple.py` | `Simple.xlsx` | Global MIP (all years) | **No** | 300 s | 7 lines, $23.0M |

The greedy solver is preserved as a reference baseline. The two ILP solvers are the active
tools. `solver_ilp` gives the full picture including tooling capital costs. `solver_simple`
is the lean version — easier to interpret, faster to tune, and a good starting point for
exploring the line allocation question before adding tooling complexity.

---

## Shared Assumptions (all three solvers)

These assumptions apply everywhere unless explicitly noted otherwise.

### Capacity model
- **Time-based**: capacity = `avail_seconds × OEE / cycle_time` (units/year).
  A line running a single product at 12 s/unit, 85% OEE = **1,586,304 units/year**.
- **Available seconds**: `hours/shift × shifts/day × days/week × weeks/year × 3600`.
  Default schedule = 7.2 h × 3 shifts × 6 days × 48 weeks = **22,394,880 s/year**.
- **Cycle times**: per-product per-line matrix (default 12 s/unit for all).
- **OEE**: per-product per-line matrix (default 0.85 for all).

### Changeover OEE penalty
Lines running **2 or more** products in the same year incur a flat OEE reduction:

```
BASE_OEE               = 0.85   (single-product line)
CHANGEOVER_OEE_PENALTY = 0.03   (multi-product reduction → effective 82%)
```

Implemented as a linear constraint: `sum(x×ct) + avail×PENALTY×multi ≤ avail×BASE_OEE`.
The `multi[l,i]` binary is forced to 1 when 2+ products are assigned; the solver sets it
to 0 on single-product lines (the relaxed 85% OEE gives more capacity, so it's naturally
optimal to do so).

### Line 1 validation
Each product that is **ever** demanded must run on Line 1 at least **once** across the
entire planning horizon. Once it has run there, it is certified for life — no annual
re-validation required.

*Rationale*: Line 1 opens first (2029, tiny demand = 4,000 units total). All active products
easily fit on Line 1 in that year, so validation is a free by-product of a line you needed
anyway. The cost is only tooling, not idle-line capital.

### Lines don't close
Once a line is commissioned (opens in any year), it remains open for all subsequent years
— `o[l,i] ≥ o[l,i−1]`. A commissioned production line is a physical asset; decommissioning
is not modelled. Consequence: lines opened early pay running cost in every subsequent year
even when demand later falls below single-line capacity.

### Minimum opening utilisation (lines 2+)
A new line (first year it transitions from closed to open) must achieve **≥ 30%** of raw
available seconds in that year. Line 1 is exempt (opens for validation at tiny demand).

*Why 30%*: Blocks the most egregious premature openings (observed case: a line opening at
28.4% when existing lines still had spare capacity) while remaining comfortably below the
~47% worst-case overflow that occurs when a new line is genuinely needed, avoiding
infeasibility.

### Demand-year filtering
Only years with total demand > 0 are included in the ILP. The Demand sheet covers
2025–2041 (17 years), but only 2029–2038 have non-zero demand. Including zero-demand years
causes infeasibility: the no-close constraint forces `o[l,i] = 1` while zero demand forces
`o[l,i] = 0` — a direct contradiction.

### Independent year blocks
Each year is an independent capacity constraint block. There is no minimum lot size,
campaign length, or demand smoothing across years. A product can appear on a line in 2031,
disappear in 2032–2033, and reappear in 2034 with no carry-over cost (within the
assumptions of each solver).

### Fixed parameters
- 10 products supported (P1–P10); P5–P10 currently zero demand
- 15 lines available
- Demand horizon: 2025–2041 (active: 2029–2038)

---

## solver_greedy.py — Greedy Heuristic

**Status**: Reference baseline. Not actively developed.

**Algorithm**: Year-by-year. For each year, checks if current line capacity is exceeded.
If so, opens the cheapest additional line and greedily assigns overflow products to it.
Does not look ahead or globally optimise — decisions made in year Y cannot be revised
for year Y+1.

**Objective**: Minimise tooling sets (proxy for minimising cost).

**Result with current data**:
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

The greedy approach mixes products across overflow lines, requiring more tooling sets per
line than necessary. No premature line opening, but no global tooling optimisation.

---

## solver_ilp.py — Full ILP Solver

**Status**: Active. Workbook: `Solver.xlsx`. Run: `python solver_ilp.py Solver.xlsx`.

**Algorithm**: Mixed-Integer Linear Program solved with PuLP/CBC. All years are optimised
simultaneously. Tracks mechanical and optical tooling families, late introduction costs, and
tooling permanence.

### Decision variables
| Variable | Type | Meaning |
|---|---|---|
| `x[p,l,i]` | Continuous ≥ 0 | Units of product p on line l in year i |
| `u[p,l,i]` | Binary | Product p assigned to line l in year i |
| `o[l,i]` | Binary | Line l open (producing) in year i |
| `mul[l,i]` | Binary | Line l runs 2+ products in year i |
| `tm[f,l]` | Binary | Mech tooling family f ever placed on line l |
| `to[f,l]` | Binary | Optical tooling family f ever placed on line l |
| `ever_open[l]` | Binary | Line l ever commissioned (one-time capital) |
| `late_intro[p,l]` | Binary | Product p first introduced to line l after it opened |
| `ever_been_u[p,l,i]` | Binary | Product p has run on line l at any point up to year i |

### Objective — minimise total USD cost
```
min  cost_line       × Σ ever_open[l]
   + cost_running    × Σ o[l,i]
   + (cost_upgrade + cost_validation) × Σ late_intro[p,l]
   + cost_mech       × Σ tm[f,l]
   + cost_opt        × Σ to[f,l]
```

### Cost defaults (editable in Solver.xlsx → Parameters, rows 50-55)
| Item | Default | Notes |
|------|---------|-------|
| One-time line capital | $0 | Absorbed into annual running cost |
| Annual running cost | $500,000/yr | $3.5M capital / 7-year life |
| Hardware upgrade (late intro) | $500,000 | Per product added to a live line |
| Re-validation (late intro) | $100,000 | Per late product intro event |
| Mechanical tooling set | $110,000 | Per family per line |
| Optical tooling set | $220,000 | Per family per line |

### Constraints (in addition to shared assumptions)
1. **Demand satisfaction**: total units = demand for every product-year.
2. **Zero-demand suppression**: if `d[p,yr] = 0`, force `u[p,l,i] = 0`.
3. **x ≤ M·u**: production only when product is assigned.
4. **Line open iff producing**: `o[l,i] ↔ Σ_p u[p,l,i] > 0`.
5. **Multi-product flag**: `Σ_p u[p,l,i] − 1 ≤ (P−1) × mul[l,i]`.
6. **Capacity with OEE penalty** (shared).
7. **Line 1 validation** (shared).
8. **Lines don't close** (shared).
9. **Min opening utilisation** (shared).
10. **Tooling permanence**: `tm[f,l] ≥ u[p,l,i]` for every product p in family f — once placed, never removed.
11. **ever_open**: `ever_open[l] ≥ o[l,i]` for every year.
12. **ever_been_u**: cumulative non-decreasing indicator; `ever_been_u[p,l,i] ≥ u[p,l,i]` and `ever_been_u[p,l,i] ≥ ever_been_u[p,l,i−1]`.
13. **late_intro**: fires when `u[p,l,i] = 1` AND `ever_been_u[p,l,i−1] = 0` AND `o[l,i−1] = 1`. Does NOT re-fire if the product stops and restarts — `ever_been_u` tracks history permanently.

### Tooling families
Reads the 10×10 mech and optical compatibility matrices from the Tooling sheet. Computes
connected components (union-find) — products in the same family share one physical set per line.
Currently identity matrices (no sharing). Updating the matrices is the main lever for reducing
tooling counts without changing the solver logic.

### Current result (600-second run, not proven optimal)
- **7 lines**, 44 line-years, 24 tooling sets (12 mech + 12 opt)
- **$25.96M** (running $22M + tooling $3.96M)
- LP lower bound: $18.04M → gap 44%
- Theory minimum with 6 lines: ~$24.6M (40 line-years + 28 tooling sets)

The ILP finds fewer tooling sets than the theoretical minimum (24 vs 28) by using more
product-dedicated lines, but pays one extra line-year. Given the 44% LP gap, the solver has
not yet proven optimality — more time would likely find the 6-line solution.

---

## solver_simple.py — Simplified ILP

**Status**: Active. Workbook: `Simple.xlsx`. Run: `python solver_simple.py Simple.xlsx`.

**Purpose**: Answer the line allocation question without tooling complexity. Useful for:
- Quickly iterating on demand forecasts or schedule changes
- Understanding line count and changeover trade-offs independently of tooling
- A cleaner starting point for adding new constraints without entangling tooling logic

**What is dropped vs solver_ilp**:
- No tooling variables (`tm`, `to`, `ever_open`)
- No late introduction variables (`late_intro`, `ever_been_u`)
- No tooling matrices or tooling cost parameters
- No upgrade / validation costs
- No Report sheet or Gantt chart

### Decision variables
| Variable | Type | Meaning |
|---|---|---|
| `x[p,l,i]` | Continuous ≥ 0 | Units of product p on line l in year i |
| `u[p,l,i]` | Binary | Product p assigned to line l in year i |
| `o[l,i]` | Binary | Line l open in year i |
| `mul[l,i]` | Binary | Line l runs 2+ products in year i |

### Objective — minimise running cost + changeover penalty
```
min  cost_running    × Σ o[l,i]
   + cost_changeover × Σ mul[l,i]
```

Two levers (editable in Simple.xlsx → Parameters, rows 51-52):
- **running cost** ($500K/yr default): discourages opening lines before capacity demands it.
- **changeover penalty** ($100K/mixed-line-yr default): soft preference for single-product
  lines (Rule 3). Increase to force more dedicated lines; decrease to allow more mixing.

### Constraints
Same shared assumptions as solver_ilp (demand satisfaction, capacity with OEE penalty, Line 1
validation, lines don't close, min opening utilisation), but without constraints 10–13
(tooling permanence, ever_open, ever_been_u, late_intro).

**Problem size**: 3,300 variables vs 5,265 in solver_ilp (37% smaller).

### Current result (300-second run, not proven optimal)
- **7 lines**, 44 line-years, 10 mixed-line-years
- **$23.0M** (running $22M + changeover penalty $1M)
- LP lower bound: $18.73M → gap 23%

The smaller problem size gives a tighter LP bound and faster convergence. Given more time,
the solver would likely find the 6-line optimal solution.

---

## Theoretical Minimum (hand-computed)

With current demand data (P1–P4, 2029–2038), the minimum-cost allocation is:

```
Line | Role              | Products     | Intro | Lines open by year
L1   | Validation        | P1,P2,P3,P4  | 2029  | 2029: 1
L2   | P2 dedicated      | P2           | 2031  | 2030: 1
L3   | Small-prod shared | P1,P3,P4     | 2032  | 2031: 2
L4   | P2 dedicated      | P2           | 2033  | 2032: 3
L5   | Small-prod shared | P1,P3,P4     | 2034  | 2033: 4
L6   | P2+P3 overflow    | P2,P3        | 2035  | 2034: 5  2035: 6  2036: 6
```

- 6 lines, 40 line-years, 14 mech + 14 opt = 28 tooling sets
- Estimated cost (solver_ilp model): ~$24.6M
- Note: with the no-close constraint, lines opened in 2035 remain open through 2038
  even as demand falls — this is factored into the theory cost.

---

## File Structure

```
LineToolingOptimization/
├── Solver.xlsx             — ILP workbook (6 sheets: Instructions/Demand/Parameters/Tooling/Allocation/Report)
├── solver_ilp.py           — Full ILP solver: reads/writes Solver.xlsx; tracks tooling
├── create_solver_xlsx.py   — One-time: generates Solver.xlsx from scratch (or copies from Solver.xlsx)
│
├── Simple.xlsx             — Simple workbook (3 sheets: Demand/Parameters/Allocation)
├── solver_simple.py        — Simplified ILP: no tooling tracking; reads/writes Simple.xlsx
├── create_simple_xlsx.py   — One-time: generates Simple.xlsx, auto-copies from Solver.xlsx
│
├── solver_greedy.py        — Original greedy solver (reference / fallback)
├── solver.py               — Legacy alias → solver_greedy.py
│
├── requirements.txt        — pulp, matplotlib, openpyxl
├── tooling_summary.csv     — Side output from solver_ilp: allocation + tooling IDs
├── line_gantt.png          — Side output from solver_ilp: Gantt chart
└── CLAUDE.md               — This file
```

---

## How to Run

### First-time setup
```bash
pip install -r requirements.txt
```

### solver_ilp  (full model with tooling)
```bash
# Generate workbook once (or to reset to defaults):
python create_solver_xlsx.py                  # built-in defaults
python create_solver_xlsx.py Book.xlsx        # copy from legacy workbook

# Edit yellow cells in Solver.xlsx (Demand, Parameters, Tooling sheets), then:
python solver_ilp.py Solver.xlsx
```

### solver_simple  (lean model, no tooling)
```bash
# Generate workbook once:
python create_simple_xlsx.py                  # built-in defaults
python create_simple_xlsx.py Solver.xlsx      # copy demand/OEE/CT from Solver.xlsx

# Edit yellow cells in Simple.xlsx (Demand, Parameters sheets), then:
python solver_simple.py Simple.xlsx
```

### Key solver constants (in code, not in workbook)
| Constant | Both solvers | Value |
|----------|-------------|-------|
| `BASE_OEE` | ✓ | 0.85 |
| `CHANGEOVER_OEE_PENALTY` | ✓ | 0.03 |
| `MIN_OPENING_UTIL` | ✓ | 0.30 |
| `NUM_LINES` | ✓ | 15 |
| `NUM_PRODUCTS` | ✓ | 10 |
| `SOLVER_TIME_LIMIT` | solver_ilp: 600 s / solver_simple: 300 s | — |

---

## Workbook Layout Reference

### Solver.xlsx (solver_ilp)

**Inputs (yellow cells):**
| Sheet | Contents | Key rows / cols |
|-------|----------|-----------------|
| Demand | Annual units per product, 2025-2041 | rows 5-21, col 1=year, cols 2-11=P1-P10 |
| Parameters | Schedule | rows 4-7 col B |
| Parameters | Cycle times (s/unit) per line | rows 14-28, cols 2-11 |
| Parameters | OEE per line per product | rows 33-47, cols 2-11 |
| Parameters | Cost parameters | rows 50-55 col B |
| Tooling | Mech compatibility matrix 10×10 | rows 7-16, cols 2-11 |
| Tooling | Optical compatibility matrix 10×10 | rows 21-30, cols 2-11 |

**Outputs (overwritten by solver):**
| Sheet | Contents |
|-------|----------|
| Allocation | Units per (product, line, year); utilisation % |
| Report | Line summary, tooling ID registry, demand validation |

### Simple.xlsx (solver_simple)

**Inputs (yellow cells):**
| Sheet | Contents | Key rows / cols |
|-------|----------|-----------------|
| Demand | Same format as Solver.xlsx | rows 5-21, col 1=year, cols 2-11=P1-P10 |
| Parameters | Schedule | rows 4-7 col B |
| Parameters | Cycle times (s/unit) per line | rows 14-28, cols 2-11 |
| Parameters | OEE per line per product | rows 34-48, cols 2-11 |
| Parameters | Running cost, changeover penalty | rows 51-52 col B |

**Outputs (overwritten by solver):**
| Sheet | Contents |
|-------|----------|
| Allocation | Product mix per line per year (green=single, amber=multi); open-line count; utilisation % |

---

## What to Refine Next

### High priority
- **solver_ilp: prove optimality** — increase `SOLVER_TIME_LIMIT` or warm-start with the
  simple solver's solution. Current gap: 44%; theory min is ~$24.6M vs $25.96M found.
- **solver_simple: prove optimality** — current gap 23%. Likely reaches 6-line solution
  with 600-second limit.
- **Per-line cost variation** (solver_ilp) — currently all lines share the same cost.
  Add per-line cost tables to the Parameters sheet and update `load_inputs()`.

### Medium priority
- **Tooling sharing matrices** — currently identity (each product = its own family).
  Updating the 10×10 matrices in Solver.xlsx is the main lever for reducing tooling counts.
- **solver_simple: changeover penalty calibration** — $100K/mixed-line-yr is a placeholder.
  Calibrate against the actual OEE impact: 3% × avail × running_rate × (production_cost/unit).

### Low priority / future
- Multi-year demand smoothing (currently each year is independent)
- Minimum lot size / campaign length constraints
- Changeover *time* penalty (currently only OEE impact; add a changeover time matrix)
- Warm-start solver_ilp from solver_simple output to reduce gap faster

---

## Demand Summary (active years, P1-P4)

| Year | P1 | P2 | P3 | P4 | Total | Lines (theory) |
|------|----|----|----|----|-------|----------------|
| 2029 | 1,000 | 1,000 | 1,000 | 1,000 | 4,000 | 1 |
| 2030 | 35,498 | 141,992 | 35,498 | 35,498 | 248,486 | 1 |
| 2031 | 372,560 | 1,490,238 | 372,560 | 372,560 | 2,607,918 | 2 |
| 2032 | 529,763 | 2,119,050 | 529,763 | 529,763 | 3,708,339 | 3 |
| 2033 | 693,177 | 2,772,708 | 693,177 | 693,177 | 4,852,239 | 4 |
| 2034 | 1,001,694 | 4,006,776 | 1,001,694 | 1,001,694 | 7,011,858 | 5 |
| 2035 | 1,338,697 | 5,354,788 | 1,338,697 | 1,338,697 | 9,370,879 | 6 |
| 2036 | 1,286,832 | 5,417,328 | 1,286,832 | 1,286,832 | 9,277,824 | 6 |
| 2037 | 610,237 | 2,440,948 | 610,237 | 610,237 | 4,271,659 | 3 (no-close → 6) |
| 2038 | 173,752 | 695,008 | 173,752 | 173,752 | 1,216,264 | 1 (no-close → 6) |
