#!/usr/bin/env python3
"""
LineToolingOptimization — ILP Solver
Replaces the greedy heuristic with a Mixed-Integer Linear Program (PuLP/CBC).

Improvements over solver_greedy.py:
  - Optimises ALL years simultaneously (global, not myopic)
  - Exploits tooling-sharing matrices (counts families, not raw products)
  - Models changeover-OEE penalty for multi-product lines (configurable)
  - Tracks named physical tooling IDs (MECH-P01 / OPTI-P01) per line per year
  - Outputs summary CSV + Gantt PNG in addition to the Excel grid

Priority order in objective:
  1. Minimise total lines open across all years
  2. Minimise tooling sets + product-line switches

Usage: python solver_ilp.py [workbook.xlsx]
Requirements: pip install pulp matplotlib openpyxl
"""
import sys
import csv
import argparse
from pathlib import Path
from collections import defaultdict

import openpyxl
import pulp
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ── Sheet names (Solver.xlsx layout — created by create_solver_xlsx.py) ───────
DEMAND_SHEET  = 'Demand'
PARAMS_SHEET  = 'Parameters'
TOOLING_SHEET = 'Tooling'
ALLOC_SHEET   = 'Allocation'
REPORT_SHEET  = 'Report'

NUM_PRODUCTS = 10
NUM_LINES    = 15

# 'Demand' sheet positions
DEM_DATA_ROW = 5    # first data row (year 2025); col 1=year, col 2=P1 … col 11=P10

# 'Parameters' sheet positions
PAR_HOURS_ROW    = 4    # hours/shift value in col B
PAR_SHIFTS_ROW   = 5
PAR_DAYS_ROW     = 6
PAR_WEEKS_ROW    = 7
PAR_VAL_COL      = 2
PAR_CT_DATA_ROW  = 14   # first cycle-time data row (Line 1); col 2=P1 … col 11=P10
PAR_OEE_DATA_ROW = 33   # first OEE data row (Line 1)
PAR_COST_LINE_ROW = 50  # cost of a new production line
PAR_COST_UPGR_ROW = 51  # cost of upgrading an established line with a new product
PAR_COST_MECH_ROW = 52  # cost per mechanical tooling set
PAR_COST_OPT_ROW  = 53  # cost per optical tooling set
PAR_COST_VALD_ROW = 54  # validation cost per late product introduction

# 'Tooling' sheet positions
TOOLING_MECH_ROW = 7    # first mech matrix data row (P1); col 2=P1 … col 11=P10
TOOLING_OPT_ROW  = 21   # first optical matrix data row (P1)

# 'Allocation' sheet output positions
ALLOC_HDR_ROW  = 5
ALLOC_DATA_ROW = 6

# ── Solver configuration (edit these to change behaviour) ─────────────────────
BASE_OEE               = 0.85   # OEE for single-product lines (default)
CHANGEOVER_OEE_PENALTY = 0.03   # OEE reduction for lines running 2+ products
SOLVER_TIME_LIMIT      = 300    # CBC wall-clock limit in seconds

# Default costs (USD) — overridden by values in the Parameters sheet
DEFAULT_COST_LINE       = 3_500_000   # one-time capital per new line opened
DEFAULT_COST_UPGRADE    =   500_000   # hardware upgrade per product added to established line
DEFAULT_COST_MECH       =   110_000   # per mechanical tooling set purchased
DEFAULT_COST_OPT        =   220_000   # per optical tooling set purchased
DEFAULT_COST_VALIDATION =   100_000   # process validation per late product intro on a line

# ── Product colours for Gantt chart ───────────────────────────────────────────
PRODUCT_COLORS = [
    '#4C72B0', '#DD8452', '#55A868', '#C44E52',
    '#8172B3', '#937860', '#DA8BC3', '#8C8C8C',
    '#CCB974', '#64B5CD',
]


# ─────────────────────────────────────────────────────────────────────────────
# INPUT LOADING  (reads from Solver.xlsx multi-sheet layout)
# ─────────────────────────────────────────────────────────────────────────────
def load_inputs(wb):
    """Read all inputs from the Solver.xlsx sheet layout and return a dict."""
    ws_d = wb[DEMAND_SHEET]
    ws_p = wb[PARAMS_SHEET]
    ws_t = wb[TOOLING_SHEET]

    # Production setup → available seconds per year
    setup = {
        'hours':  float(ws_p.cell(row=PAR_HOURS_ROW,  column=PAR_VAL_COL).value or 0),
        'shifts': float(ws_p.cell(row=PAR_SHIFTS_ROW, column=PAR_VAL_COL).value or 0),
        'days':   float(ws_p.cell(row=PAR_DAYS_ROW,   column=PAR_VAL_COL).value or 0),
        'weeks':  float(ws_p.cell(row=PAR_WEEKS_ROW,  column=PAR_VAL_COL).value or 0),
    }
    avail = setup['hours'] * setup['shifts'] * setup['days'] * setup['weeks'] * 3600

    # Demand: scan rows from DEM_DATA_ROW until year column is empty
    years, demand = [], {}
    for r in range(DEM_DATA_ROW, DEM_DATA_ROW + 50):
        yr = ws_d.cell(row=r, column=1).value
        if yr is None:
            break
        yr = int(yr)
        years.append(yr)
        demand[yr] = [int(ws_d.cell(row=r, column=2 + p).value or 0)
                      for p in range(10)]

    # Cycle times and OEE: col 2 = P1 … col 11 = P10
    ct  = [[ws_p.cell(row=PAR_CT_DATA_ROW  + l, column=2 + p).value or 12
            for p in range(10)] for l in range(15)]
    oee = [[ws_p.cell(row=PAR_OEE_DATA_ROW + l, column=2 + p).value or 0.85
            for p in range(10)] for l in range(15)]

    # Tooling compatibility matrices: col 2 = P1 … col 11 = P10
    mt  = [[ws_t.cell(row=TOOLING_MECH_ROW + i, column=2 + j).value or 0
            for j in range(10)] for i in range(10)]
    ot  = [[ws_t.cell(row=TOOLING_OPT_ROW  + i, column=2 + j).value or 0
            for j in range(10)] for i in range(10)]

    # Cost parameters (fall back to defaults if cells are empty)
    def _cost(row, default):
        v = ws_p.cell(row=row, column=PAR_VAL_COL).value
        return float(v) if v is not None else default

    costs = {
        'line':       _cost(PAR_COST_LINE_ROW, DEFAULT_COST_LINE),
        'upgrade':    _cost(PAR_COST_UPGR_ROW, DEFAULT_COST_UPGRADE),
        'mech':       _cost(PAR_COST_MECH_ROW, DEFAULT_COST_MECH),
        'opt':        _cost(PAR_COST_OPT_ROW,  DEFAULT_COST_OPT),
        'validation': _cost(PAR_COST_VALD_ROW, DEFAULT_COST_VALIDATION),
    }

    return {'years': years, 'demand': demand, 'avail': avail,
            'ct': ct, 'oee': oee, 'mt': mt, 'ot': ot, 'costs': costs}


# ─────────────────────────────────────────────────────────────────────────────
# TOOLING FAMILIES  (union-find on compatibility matrix)
# ─────────────────────────────────────────────────────────────────────────────
def compute_families(matrix, n):
    """
    Return a list of frozensets, each being a connected component of the
    product-compatibility graph defined by *matrix* (1 = compatible / shared).
    Products that share tooling belong to the same family; one physical set
    serves the whole family on a given line.
    """
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    for i in range(n):
        for j in range(n):
            if matrix[i][j]:
                union(i, j)

    groups: dict[int, set] = defaultdict(set)
    for i in range(n):
        groups[find(i)].add(i)
    return [frozenset(s) for s in groups.values()]


# ─────────────────────────────────────────────────────────────────────────────
# ILP BUILD & SOLVE
# ─────────────────────────────────────────────────────────────────────────────
def build_and_solve(inp):
    """
    Build the Mixed-Integer Linear Program and solve with CBC.

    Returns (alloc, tooled, intro, mech_sets, opt_sets, mech_fams, opt_fams)
    or None on failure.

    alloc  : dict (year, line, product) -> int units
    tooled : list[set]  line -> set of product indices ever run on that line
    intro  : dict       line -> first year the line was used
    """
    years  = inp['years']
    demand = inp['demand']
    avail  = inp['avail']
    ct     = inp['ct']

    P = NUM_PRODUCTS
    L = NUM_LINES
    Y = len(years)

    # Active products per year (demand > 0)
    active = {yr: [p for p in range(P) if demand[yr][p] > 0] for yr in years}

    # Big-M: max units any single product could demand in any year
    M_big = max(demand[yr][p] for yr in years for p in range(P)) + 1

    # Tooling families from compatibility matrices
    mech_fams = compute_families(inp['mt'], P)
    opt_fams  = compute_families(inp['ot'], P)
    MF = len(mech_fams)
    OF = len(opt_fams)

    # Product -> family index lookup
    prod_to_mf = {}
    for f, fam in enumerate(mech_fams):
        for p in fam:
            prod_to_mf[p] = f
    prod_to_of = {}
    for f, fam in enumerate(opt_fams):
        for p in fam:
            prod_to_of[p] = f

    prob = pulp.LpProblem('LineTooling', pulp.LpMinimize)

    # ── Decision variables ────────────────────────────────────────────────────

    # x[p,l,i]  continuous: units of product p on line l in year index i
    x = pulp.LpVariable.dicts(
        'x',
        [(p, l, i) for p in range(P) for l in range(L) for i in range(Y)],
        lowBound=0, cat='Continuous')

    # u[p,l,i]  binary: 1 if product p produced on line l in year i
    u = pulp.LpVariable.dicts(
        'u',
        [(p, l, i) for p in range(P) for l in range(L) for i in range(Y)],
        cat='Binary')

    # o[l,i]  binary: 1 if line l is open (any product) in year i
    o = pulp.LpVariable.dicts(
        'open',
        [(l, i) for l in range(L) for i in range(Y)],
        cat='Binary')

    # mul[l,i]  binary: 1 if line l runs 2+ products in year i  (OEE penalty)
    mul = pulp.LpVariable.dicts(
        'multi',
        [(l, i) for l in range(L) for i in range(Y)],
        cat='Binary')

    # tm[f,l]  binary: 1 if mech family f is ever on line l (permanent once added)
    tm = pulp.LpVariable.dicts(
        'tm',
        [(f, l) for f in range(MF) for l in range(L)],
        cat='Binary')

    # to_[f,l]  binary: 1 if optical family f is ever on line l
    to_ = pulp.LpVariable.dicts(
        'to',
        [(f, l) for f in range(OF) for l in range(L)],
        cat='Binary')

    # ever_open[l]  binary: 1 if line l is opened at any point (one-time capital)
    ever_open = pulp.LpVariable.dicts(
        'ever_open',
        [l for l in range(L)],
        cat='Binary')

    # late_intro[p,l]  binary: 1 if product p is introduced to line l after
    # the line's initial commissioning year (triggers upgrade + validation cost)
    late_intro = pulp.LpVariable.dicts(
        'late_intro',
        [(p, l) for p in range(P) for l in range(L)],
        cat='Binary')

    costs = inp['costs']

    # ── Objective  (minimise total USD cost) ──────────────────────────────────
    prob += (
        costs['line']       * pulp.lpSum(ever_open[l] for l in range(L))
      + (costs['upgrade'] + costs['validation'])
                            * pulp.lpSum(late_intro[p, l]
                                         for p in range(P) for l in range(L))
      + costs['mech']      * pulp.lpSum(tm[f, l]
                                         for f in range(MF) for l in range(L))
      + costs['opt']       * pulp.lpSum(to_[f, l]
                                         for f in range(OF) for l in range(L))
    )

    # ── Constraints ───────────────────────────────────────────────────────────
    for i, yr in enumerate(years):
        d   = demand[yr]
        act = active[yr]

        # 1. Demand satisfaction: all demand must be allocated
        for p in range(P):
            prob += (pulp.lpSum(x[p, l, i] for l in range(L)) == d[p],
                     f'dem_{p}_{i}')

        # 2. If demand = 0, forbid production flags (tightens the model)
        for p in range(P):
            if d[p] == 0:
                for l in range(L):
                    prob += (u[p, l, i] == 0, f'zero_{p}_{l}_{i}')

        for l in range(L):
            # 3. Production only if flagged: x <= M * u
            for p in range(P):
                prob += (x[p, l, i] <= M_big * u[p, l, i], f'xu_{p}_{l}_{i}')

            # 4. Line open iff at least one product flag is set
            for p in range(P):
                prob += (o[l, i] >= u[p, l, i], f'open_ge_{p}_{l}_{i}')
            prob += (o[l, i] <= pulp.lpSum(u[p, l, i] for p in range(P)),
                     f'open_le_{l}_{i}')

            # 5. Multi-product flag: forced to 1 when 2+ products present
            #    sum_p u[p,l,i] - 1 <= (P-1) * mul[l,i]
            prob += (pulp.lpSum(u[p, l, i] for p in range(P)) - 1
                     <= (P - 1) * mul[l, i],
                     f'multi_{l}_{i}')
            # multi can only be 1 when line is open
            prob += (mul[l, i] <= o[l, i], f'multi_le_open_{l}_{i}')

            # 6. Capacity with linearised OEE penalty
            #    sum_p(x * ct) + avail * PENALTY * mul <= avail * BASE_OEE
            #    (mul on LHS keeps the constraint linear)
            prob += (
                pulp.lpSum(x[p, l, i] * ct[l][p] for p in range(P))
                + avail * CHANGEOVER_OEE_PENALTY * mul[l, i]
                <= avail * BASE_OEE,
                f'cap_{l}_{i}'
            )

        # 7. Line 1 validation constraint:
        #    ALL active products MUST be produced on Line 1 every year
        for p in act:
            prob += (u[p, 0, i] == 1, f'val1_{p}_{i}')

    # 8. Tooling permanence: if product runs on line in any year, family tooling needed
    for f, fam in enumerate(mech_fams):
        for l in range(L):
            for p in fam:
                for i in range(Y):
                    prob += (tm[f, l] >= u[p, l, i], f'tm_{f}_{l}_{p}_{i}')

    for f, fam in enumerate(opt_fams):
        for l in range(L):
            for p in fam:
                for i in range(Y):
                    prob += (to_[f, l] >= u[p, l, i], f'to_{f}_{l}_{p}_{i}')

    # 9. ever_open: line counts as opened once any year-slot is active
    for l in range(L):
        for i in range(Y):
            prob += (ever_open[l] >= o[l, i], f'ever_open_{l}_{i}')

    # 10. late_intro: fires when product p arrives on line l after the line
    #     was already open (u goes 0→1 while line was open in prior year).
    #     Constraint: late_intro[p,l] >= u[p,l,i] - u[p,l,i-1] + o[l,i-1] - 1
    #     RHS = 1 only when product newly assigned AND line already existed.
    for p in range(P):
        for l in range(L):
            for i in range(1, Y):
                prob += (late_intro[p, l]
                         >= u[p, l, i] - u[p, l, i - 1] + o[l, i - 1] - 1,
                         f'late_intro_{p}_{l}_{i}')

    # ── Solve ─────────────────────────────────────────────────────────────────
    n_vars  = len(prob.variables())
    n_cons  = len(prob.constraints)
    print(f'  Variables: {n_vars:,}   Constraints: {n_cons:,}')

    solver = pulp.PULP_CBC_CMD(timeLimit=SOLVER_TIME_LIMIT, msg=1)
    prob.solve(solver)

    status = pulp.LpStatus[prob.status]
    obj_val = pulp.value(prob.objective)
    print(f'  Solver status: {status}   Objective: {obj_val:,.0f}')

    # Accept Optimal (1) or time-limit with feasible (-2)
    if prob.status not in (1, -2):
        print('  ERROR: solver did not find a feasible solution.')
        return None

    # ── Extract solution ───────────────────────────────────────────────────────
    alloc  = {}
    tooled = [set() for _ in range(L)]
    intro  = {}

    for i, yr in enumerate(years):
        for l in range(L):
            for p in range(P):
                v = pulp.value(x[p, l, i])
                if v is not None and v > 0.5:
                    units = round(v)
                    alloc[(yr, l, p)] = units
                    tooled[l].add(p)
                    if l not in intro:
                        intro[l] = yr

    mech_sets = sum(
        1 for f in range(MF) for l in range(L)
        if pulp.value(tm[f, l]) is not None and pulp.value(tm[f, l]) > 0.5
    )
    opt_sets = sum(
        1 for f in range(OF) for l in range(L)
        if pulp.value(to_[f, l]) is not None and pulp.value(to_[f, l]) > 0.5
    )
    n_lines_opened = sum(
        1 for l in range(L)
        if pulp.value(ever_open[l]) is not None and pulp.value(ever_open[l]) > 0.5
    )
    n_late_intros = sum(
        1 for p in range(P) for l in range(L)
        if pulp.value(late_intro[p, l]) is not None and pulp.value(late_intro[p, l]) > 0.5
    )

    return (alloc, tooled, intro,
            mech_sets, opt_sets, mech_fams, opt_fams,
            n_lines_opened, n_late_intros)


# ─────────────────────────────────────────────────────────────────────────────
# TOOLING ID SYSTEM
# ─────────────────────────────────────────────────────────────────────────────
def compute_tooling_ids(alloc, tooled, intro, mech_fams, opt_fams, years):
    """
    For each physical tooling set (one per family per line), return a record:
      id         e.g. 'MECH-P01'  (named by lowest product number in family)
      type       'MECH' or 'OPTI'
      line       1-based line number
      intro      first year line was opened
      year_range '2031-2038'
      years      list of years the tooling was active
    """
    records = []

    for l in range(NUM_LINES):
        if not tooled[l]:
            continue

        # Mech families present on this line
        seen_mf: set[int] = set()
        for p in tooled[l]:
            for f, fam in enumerate(mech_fams):
                if p in fam:
                    seen_mf.add(f)

        for f in sorted(seen_mf):
            fam = mech_fams[f]
            rep = min(fam) + 1              # lowest product index → P-number
            tid = f'MECH-P{rep:02d}'
            active_years = [yr for yr in years
                            if any(alloc.get((yr, l, p), 0) > 0 for p in fam)]
            if active_years:
                records.append({
                    'id': tid, 'type': 'MECH', 'line': l + 1,
                    'intro': intro.get(l, '?'),
                    'years': active_years,
                    'year_range': f'{active_years[0]}-{active_years[-1]}',
                })

        # Optical families present on this line
        seen_of: set[int] = set()
        for p in tooled[l]:
            for f, fam in enumerate(opt_fams):
                if p in fam:
                    seen_of.add(f)

        for f in sorted(seen_of):
            fam = opt_fams[f]
            rep = min(fam) + 1
            tid = f'OPTI-P{rep:02d}'
            active_years = [yr for yr in years
                            if any(alloc.get((yr, l, p), 0) > 0 for p in fam)]
            if active_years:
                records.append({
                    'id': tid, 'type': 'OPTI', 'line': l + 1,
                    'intro': intro.get(l, '?'),
                    'years': active_years,
                    'year_range': f'{active_years[0]}-{active_years[-1]}',
                })

    records.sort(key=lambda r: (r['id'], r['line']))
    return records


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT  (formatted table written to Allocation sheet)
# ─────────────────────────────────────────────────────────────────────────────
def _ofill(hex_color):
    from openpyxl.styles import PatternFill
    return PatternFill('solid', fgColor=hex_color)

def _ofont(bold=False, size=10, color='1F1F1F'):
    from openpyxl.styles import Font
    return Font(name='Calibri', bold=bold, size=size, color=color)

def _oborder():
    from openpyxl.styles import Border, Side
    t = Side(style='thin', color='B8CCE4')
    return Border(left=t, right=t, top=t, bottom=t)

def _oal(h='left'):
    from openpyxl.styles import Alignment
    return Alignment(horizontal=h, vertical='center')


def write_output(wb, alloc, tooled, intro, inp, mech_sets, opt_sets, ok):
    """Write formatted allocation table to the Allocation sheet in wb."""
    import datetime
    ws     = wb[ALLOC_SHEET]
    years  = inp['years']
    demand = inp['demand']
    ct     = inp['ct']
    avail  = inp['avail']

    # Clear old data rows (rows ALLOC_DATA_ROW onward)
    for r in range(ALLOC_DATA_ROW, ALLOC_DATA_ROW + 500):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            cell.value  = None
            cell.fill   = _ofill('FFFFFF')
            cell.border = _oborder()

    # Update subtitle row (row 2) with run metadata
    ts     = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M')
    status = '✓  All checks passed' if ok else '⚠  Issues found — see console'
    from openpyxl.styles import Font
    ws.cell(row=2, column=1).value = (
        f'Last run: {ts}   |   {status}   |   '
        f'{mech_sets} mech + {opt_sets} opt = {mech_sets + opt_sets} tooling sets'
    )
    ws.cell(row=2, column=1).font = Font(name='Calibri', italic=True,
                                         size=10, color='595959')

    C_EVEN = 'E2EFDA'   # light green
    C_ODD  = 'F0F8F0'   # slightly lighter green
    row = ALLOC_DATA_ROW
    alt = False
    for yr in years:
        d = demand.get(yr, [0] * 10)
        if sum(d) == 0:
            continue
        active_lines = sorted({l for (y2, l, _) in alloc if y2 == yr})
        for l in active_lines:
            units  = [alloc.get((yr, l, p), 0) for p in range(NUM_PRODUCTS)]
            total  = sum(units)
            prods  = '+'.join(f'P{p+1}' for p in range(NUM_PRODUCTS)
                              if units[p] > 0)
            n_prod = sum(1 for p in range(NUM_PRODUCTS) if units[p] > 0)
            t_used = sum(units[p] * ct[l][p] for p in range(NUM_PRODUCTS))
            util   = round(t_used / avail * 100, 1)
            oee_e  = int(round((BASE_OEE - CHANGEOVER_OEE_PENALTY
                                if n_prod > 1 else BASE_OEE) * 100))
            bg = C_ODD if alt else C_EVEN

            vals = (
                [yr, f'L{l+1}', intro.get(l, ''), prods]
                + [units[p] if units[p] > 0 else None
                   for p in range(NUM_PRODUCTS)]
                + [total, util, oee_e]
            )
            for ci, v in enumerate(vals, 1):
                cell            = ws.cell(row=row, column=ci)
                cell.value      = v
                cell.fill       = _ofill(bg)
                cell.font       = _ofont(bold=(ci == 1))
                cell.border     = _oborder()
                cell.alignment  = _oal(
                    'center' if ci <= 3 else
                    'right'  if ci >= 5 else
                    'left'
                )
                if ci >= 5 and ci <= 14:
                    cell.number_format = '#,##0'
                elif ci == 15:
                    cell.number_format = '#,##0'
                elif ci == 16:
                    cell.number_format = '0.0'
            ws.row_dimensions[row].height = 15
            row += 1
            alt = not alt

    return row - ALLOC_DATA_ROW


# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY CSV
# ─────────────────────────────────────────────────────────────────────────────
def write_summary_csv(alloc, tooling_records, inp, intro, out_path):
    years  = inp['years']
    avail  = inp['avail']
    ct     = inp['ct']
    demand = inp['demand']

    with open(out_path, 'w', newline='') as f:
        w = csv.writer(f)

        # Section 1 – allocation per (line, year)
        w.writerow(['=== ALLOCATION ==='])
        w.writerow(['Year', 'Line', 'Products', 'Utilisation_%', 'OEE_used_%']
                   + [f'P{p+1}_units' for p in range(NUM_PRODUCTS)])

        for yr in years:
            d = demand.get(yr, [0] * 10)
            if sum(d) == 0:
                continue
            active_lines = sorted({l for (y2, l, _) in alloc if y2 == yr})
            for l in active_lines:
                units  = [alloc.get((yr, l, p), 0) for p in range(NUM_PRODUCTS)]
                prods  = [f'P{p+1}' for p in range(NUM_PRODUCTS) if units[p] > 0]
                t_used = sum(units[p] * ct[l][p] for p in range(NUM_PRODUCTS))
                util   = t_used / avail * 100
                n_prod = len(prods)
                oee    = (BASE_OEE - CHANGEOVER_OEE_PENALTY if n_prod > 1
                          else BASE_OEE) * 100
                w.writerow([yr, f'L{l+1}', '+'.join(prods),
                             f'{util:.1f}', f'{oee:.0f}'] + units)

        w.writerow([])
        # Section 2 – tooling ID table
        w.writerow(['=== TOOLING IDs ==='])
        w.writerow(['Tooling_ID', 'Type', 'Line', 'Intro_Year', 'Years_Active'])
        for r in tooling_records:
            w.writerow([r['id'], r['type'], f"L{r['line']}", r['intro'],
                        r['year_range']])

    print(f'  Summary CSV: {out_path}')


# ─────────────────────────────────────────────────────────────────────────────
# GANTT CHART
# ─────────────────────────────────────────────────────────────────────────────
def plot_gantt(alloc, inp, intro, out_path):
    years = inp['years']

    active_lines = sorted({l for (_, l, _) in alloc})
    if not active_lines:
        return

    fig_h = max(4, len(active_lines) * 1.4 + 2)
    fig, ax = plt.subplots(figsize=(18, fig_h))
    ax.set_title('Production Line Allocation — Gantt', fontsize=14, pad=12)
    ax.set_xlabel('Year', fontsize=11)
    ax.set_ylabel('Line', fontsize=11)

    bar_height = 0.65
    yticks, ylabels = [], []

    for row_idx, l in enumerate(active_lines):
        yticks.append(row_idx)
        ylabels.append(f'L{l+1}')

        for yr in years:
            units = [alloc.get((yr, l, p), 0) for p in range(NUM_PRODUCTS)]
            total = sum(units)
            if total == 0:
                continue

            left = yr - 0.4
            for p in range(NUM_PRODUCTS):
                if units[p] == 0:
                    continue
                width = units[p] / total * 0.8  # 0.8 = bar width per year cell
                ax.barh(row_idx, width, left=left, height=bar_height,
                        color=PRODUCT_COLORS[p % len(PRODUCT_COLORS)], alpha=0.85,
                        edgecolor='white', linewidth=0.4)
                left += width

            # Mark intro year
            if intro.get(l) == yr:
                ax.annotate(
                    f'intro {yr}',
                    xy=(yr, row_idx + bar_height / 2 + 0.05),
                    xytext=(0, 3), textcoords='offset points',
                    fontsize=6.5, ha='center', va='bottom', color='#333333',
                )

    ax.set_yticks(yticks)
    ax.set_yticklabels(ylabels, fontsize=10)
    ax.set_xticks(years)
    ax.set_xticklabels([str(y) for y in years], rotation=45, ha='right', fontsize=9)
    ax.set_xlim(min(years) - 0.6, max(years) + 0.6)
    ax.set_ylim(-0.6, len(active_lines) - 0.4)

    legend_patches = [
        mpatches.Patch(color=PRODUCT_COLORS[p], label=f'P{p+1}', alpha=0.85)
        for p in range(NUM_PRODUCTS)
    ]
    ax.legend(handles=legend_patches, loc='upper right', ncol=5,
              fontsize=8, framealpha=0.9)
    ax.grid(axis='x', alpha=0.25, linestyle='--')
    ax.invert_yaxis()

    plt.tight_layout()
    plt.savefig(out_path, dpi=150)
    plt.close()
    print(f'  Gantt chart: {out_path}')


# ─────────────────────────────────────────────────────────────────────────────
# REPORT SHEET
# ─────────────────────────────────────────────────────────────────────────────
def write_report_sheet(wb, alloc, tooled, intro, mech_fams, opt_fams,
                       mech_sets, opt_sets, n_lines_opened, n_late_intros,
                       inp, ok):
    """Write lines summary, tooling ID registry, and demand validation."""
    import datetime
    ws     = wb[REPORT_SHEET]
    years  = inp['years']
    demand = inp['demand']
    ct     = inp['ct']
    avail  = inp['avail']

    # Clear everything from row 2 downward
    for r in range(2, 200):
        for c in range(1, 10):
            ws.cell(row=r, column=c).value  = None
            ws.cell(row=r, column=c).fill   = _ofill('FFFFFF')
            ws.cell(row=r, column=c).border = _oborder()

    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
    def _rfill(h): return PatternFill('solid', fgColor=h)
    def _rfont(bold=False, sz=10, color='1F1F1F'):
        return Font(name='Calibri', bold=bold, size=sz, color=color)
    def _rbd():
        t = Side(style='thin', color='B8CCE4')
        return Border(left=t, right=t, top=t, bottom=t)
    def _ral(h='left'):
        return Alignment(horizontal=h, vertical='center')
    def _sect_r(row, text, ncols=7):
        ws.cell(row=row, column=1).value = text
        ws.cell(row=row, column=1).font  = Font(name='Calibri', bold=True,
                                                size=11, color='FFFFFF')
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = _rfill('2E75B6')
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=ncols)
        ws.row_dimensions[row].height = 20
    def _hdr_r(row, labels):
        for i, lbl in enumerate(labels):
            c = ws.cell(row=row, column=1 + i)
            c.value, c.font, c.fill, c.border, c.alignment = \
                lbl, _rfont(bold=True), _rfill('BDD7EE'), _rbd(), _ral('center')
        ws.row_dimensions[row].height = 18
    def _data_r(row, vals, bg='E2EFDA'):
        for i, v in enumerate(vals):
            c = ws.cell(row=row, column=1 + i)
            c.value, c.fill, c.font, c.border = v, _rfill(bg), _rfont(), _rbd()
            c.alignment = _ral('right' if isinstance(v, (int, float)) else 'left')
        ws.row_dimensions[row].height = 15

    costs = inp['costs']
    total_cost = (
        costs['line']       * n_lines_opened
      + (costs['upgrade'] + costs['validation']) * n_late_intros
      + costs['mech']      * mech_sets
      + costs['opt']       * opt_sets
    )

    ts = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M')
    ws.cell(row=2, column=1).value = (
        f'Last run: {ts}   |   '
        f'{"✓  All OK" if ok else "⚠  Issues found"}   |   '
        f'{mech_sets} mech + {opt_sets} opt = {mech_sets + opt_sets} sets   |   '
        f'Total cost: ${total_cost:,.0f}'
    )
    ws.cell(row=2, column=1).font = Font(name='Calibri', italic=True,
                                         size=10, color='595959')
    ws.merge_cells('A2:G2')
    ws.row_dimensions[3].height = 6

    # ── Cost breakdown ────────────────────────────────────────────────────────
    _sect_r(4, '  COST BREAKDOWN')
    _hdr_r(5, ['Cost item', 'Unit cost (USD)', 'Qty', 'Total (USD)'])
    cost_items = [
        ('New production lines',        costs['line'],       n_lines_opened,
         costs['line'] * n_lines_opened),
        ('Line upgrades (new products)', costs['upgrade'],   n_late_intros,
         costs['upgrade'] * n_late_intros),
        ('Validation events',            costs['validation'], n_late_intros,
         costs['validation'] * n_late_intros),
        ('Mechanical tooling sets',      costs['mech'],      mech_sets,
         costs['mech'] * mech_sets),
        ('Optical tooling sets',         costs['opt'],       opt_sets,
         costs['opt'] * opt_sets),
    ]
    for ci, (label, unit_cost, qty, subtotal) in enumerate(cost_items):
        bg = 'F0F8F0' if ci % 2 else 'E2EFDA'
        _data_r(6 + ci, [label, unit_cost, qty, subtotal], bg=bg)
    # Total row
    total_row = 6 + len(cost_items)
    _data_r(total_row,
            ['TOTAL', '', '', total_cost],
            bg='BDD7EE')
    ws.cell(row=total_row, column=1).font = _rfont(bold=True, sz=10)
    ws.cell(row=total_row, column=4).font = _rfont(bold=True, sz=10)
    for col in [2, 4]:
        ws.cell(row=total_row - 5, column=col).number_format = '$#,##0'
    for ci in range(len(cost_items)):
        for col in [2, 4]:
            ws.cell(row=6 + ci, column=col).number_format = '$#,##0'
    ws.cell(row=total_row, column=4).number_format = '$#,##0'

    r = total_row + 2  # spacer before next section

    # ── Lines summary ─────────────────────────────────────────────────────────
    _sect_r(r, '  LINES SUMMARY'); r += 1
    _hdr_r(r, ['Line', 'Intro', 'Products', 'Tooling sets', 'Eff. OEE %',
               'Peak util %', 'Peak year']); r += 1
    for l in sorted(intro.keys()):
        ps     = sorted(f'P{p+1}' for p in tooled[l])
        n_prod = len(ps)
        oee    = int(round((BASE_OEE - CHANGEOVER_OEE_PENALTY
                            if n_prod > 1 else BASE_OEE) * 100))
        # Count tooling sets for this line
        mf_cnt = sum(1 for f, fam in enumerate(mech_fams) if fam & tooled[l])
        of_cnt = sum(1 for f, fam in enumerate(opt_fams)  if fam & tooled[l])
        # Peak utilisation
        peak_util, peak_yr = 0.0, ''
        for yr in years:
            d = demand.get(yr, [0] * 10)
            if sum(d) == 0:
                continue
            t = sum(alloc.get((yr, l, p), 0) * ct[l][p] for p in range(NUM_PRODUCTS))
            u = t / avail * 100
            if u > peak_util:
                peak_util, peak_yr = u, yr
        bg = 'F0F8F0' if r % 2 else 'E2EFDA'
        _data_r(r, [f'L{l+1}', intro[l], '+'.join(ps),
                    f'{mf_cnt}m + {of_cnt}o', oee,
                    round(peak_util, 1), peak_yr], bg=bg)
        r += 1

    r += 1  # spacer
    # ── Tooling ID registry ───────────────────────────────────────────────────
    _sect_r(r, '  TOOLING ID REGISTRY'); r += 1
    _hdr_r(r, ['Tooling ID', 'Type', 'Line', 'Intro year', 'Active years']); r += 1
    tooling_records = compute_tooling_ids(
        alloc, tooled, intro, mech_fams, opt_fams, years)
    for i, rec in enumerate(tooling_records):
        bg = 'F0F8F0' if i % 2 else 'E2EFDA'
        _data_r(r, [rec['id'], rec['type'], f"L{rec['line']}",
                    rec['intro'], rec['year_range']], bg=bg)
        r += 1

    r += 1  # spacer
    # ── Demand validation ─────────────────────────────────────────────────────
    _sect_r(r, '  DEMAND VALIDATION'); r += 1
    _hdr_r(r, ['Year', 'Product', 'Demand', 'Allocated', 'Difference', 'Status']); r += 1
    issues = 0
    for yr in years:
        d = demand.get(yr, [0] * 10)
        if sum(d) == 0:
            continue
        for p in range(NUM_PRODUCTS):
            alloc_p = sum(alloc.get((yr, l, p), 0) for l in range(NUM_LINES))
            diff    = alloc_p - d[p]
            if d[p] == 0 and alloc_p == 0:
                continue
            ok_row = abs(diff) <= 1
            if not ok_row:
                issues += 1
            bg = 'E2EFDA' if ok_row else 'FFC7CE'
            _data_r(r, [yr, f'P{p+1}', d[p], alloc_p, diff,
                        'OK' if ok_row else 'MISMATCH'], bg=bg)
            r += 1
    if issues == 0:
        ws.cell(row=r, column=1).value = '✓  All demand satisfied within ±1 unit tolerance.'
        ws.cell(row=r, column=1).font = Font(name='Calibri', italic=True,
                                             size=10, color='2E75B6')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)


# ─────────────────────────────────────────────────────────────────────────────
# VERIFICATION  (demand met & capacity not exceeded)
# ─────────────────────────────────────────────────────────────────────────────
def verify(alloc, inp):
    ok = True
    for yr in inp['years']:
        d = inp['demand'].get(yr, [0] * 10)
        if sum(d) == 0:
            continue

        # Demand check
        ta = [0] * 10
        for l in range(NUM_LINES):
            for p in range(10):
                ta[p] += alloc.get((yr, l, p), 0)
        for p in range(10):
            if abs(d[p] - ta[p]) > 1:
                print(f'  X {yr} P{p+1}: demand={d[p]:,}  alloc={ta[p]:,}')
                ok = False

        # Capacity check: multi-product lines use the reduced OEE capacity
        for l in range(NUM_LINES):
            units_l = [alloc.get((yr, l, p), 0) for p in range(10)]
            t = sum(units_l[p] * inp['ct'][l][p] for p in range(10))
            n_prod = sum(1 for p in range(10) if units_l[p] > 0)
            oee_eff = (BASE_OEE - CHANGEOVER_OEE_PENALTY if n_prod > 1 else BASE_OEE)
            cap = inp['avail'] * oee_eff
            if t > cap * 1.001:
                print(f'  X {yr} Line {l+1}: capacity exceeded '
                      f'({t:,.0f} > {cap:,.0f}, OEE={oee_eff*100:.0f}%)')
                ok = False

    return ok


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='Production line ILP optimizer')
    parser.add_argument('workbook', nargs='?', default='Solver.xlsx')
    args = parser.parse_args()

    path = Path(args.workbook)
    if not path.exists():
        print(f'Error: {path} not found')
        print('  Run  python create_solver_xlsx.py  to generate Solver.xlsx first.')
        sys.exit(1)

    print(f'Loading {path} ...')
    wb  = openpyxl.load_workbook(str(path))
    inp = load_inputs(wb)

    years = inp['years']
    print(f'  Available seconds/year : {inp["avail"]:,.0f}')
    print(f'  Years                  : {years[0]}-{years[-1]}  ({len(years)} years)')
    print(f'  BASE_OEE               : {BASE_OEE*100:.0f}%')
    print(f'  CHANGEOVER_OEE_PENALTY : {CHANGEOVER_OEE_PENALTY*100:.0f}%')

    mf_preview = compute_families(inp['mt'], NUM_PRODUCTS)
    of_preview = compute_families(inp['ot'], NUM_PRODUCTS)
    print(f'  Mech tooling families  : {len(mf_preview)}'
          f'  (sizes: {sorted(len(f) for f in mf_preview)})')
    print(f'  Opt  tooling families  : {len(of_preview)}'
          f'  (sizes: {sorted(len(f) for f in of_preview)})')
    c = inp['costs']
    print(f'  Cost — line: ${c["line"]:,.0f}  upgrade: ${c["upgrade"]:,.0f}  '
          f'mech: ${c["mech"]:,.0f}  opt: ${c["opt"]:,.0f}  '
          f'validation: ${c["validation"]:,.0f}')

    print('\nBuilding & solving ILP ...')
    result = build_and_solve(inp)
    if result is None:
        print('Solver failed — no solution found.')
        sys.exit(1)

    alloc, tooled, intro, mech_sets, opt_sets, mech_fams, opt_fams, \
        n_lines_opened, n_late_intros = result

    print('\nVerifying solution ...')
    ok = verify(alloc, inp)
    print(f'  {"All checks passed" if ok else "Issues found — see above"}')

    total_sets = mech_sets + opt_sets
    print(f'\nTooling: {mech_sets} mech + {opt_sets} opt = {total_sets} sets')
    print(f'  (greedy baseline: 19+19=38  |  theory min: 14+14=28)')

    costs = inp['costs']
    cost_lines  = costs['line']       * n_lines_opened
    cost_late   = (costs['upgrade'] + costs['validation']) * n_late_intros
    cost_mech   = costs['mech']      * mech_sets
    cost_opt    = costs['opt']       * opt_sets
    total_cost  = cost_lines + cost_late + cost_mech + cost_opt
    print(f'\nCost breakdown:')
    print(f'  Lines opened      : {n_lines_opened}  × ${costs["line"]:,.0f}  = ${cost_lines:,.0f}')
    print(f'  Late intros       : {n_late_intros}  × ${costs["upgrade"]+costs["validation"]:,.0f}  = ${cost_late:,.0f}')
    print(f'  Mech tooling sets : {mech_sets}  × ${costs["mech"]:,.0f}  = ${cost_mech:,.0f}')
    print(f'  Opt  tooling sets : {opt_sets}  × ${costs["opt"]:,.0f}  = ${cost_opt:,.0f}')
    print(f'  ─────────────────────────────────────────────────')
    print(f'  TOTAL             :                    ${total_cost:,.0f}')

    print('\nLines used:')
    for l in sorted(intro.keys()):
        ps    = sorted(f'P{p+1}' for p in tooled[l])
        n_p   = len(ps)
        oee_e = (BASE_OEE - CHANGEOVER_OEE_PENALTY if n_p > 1 else BASE_OEE) * 100
        print(f'  Line {l+1:2d}  intro={intro[l]}  products={ps}'
              f'  effective_OEE={oee_e:.0f}%')

    print('\nPer-year summary:')
    for yr in years:
        d = inp['demand'].get(yr, [0] * 10)
        if sum(d) == 0:
            continue
        active_lines = sorted({l for (y2, l, _) in alloc if y2 == yr})
        parts = [f'L{l+1}=[{",".join(f"P{p+1}" for p in range(NUM_PRODUCTS) if alloc.get((yr,l,p),0)>0)}]'
                 for l in active_lines]
        print(f'  {yr}: {", ".join(parts)}  —  {len(active_lines)} line(s)')

    print('\nWriting outputs to workbook ...')
    outdir = path.parent

    n_rows = write_output(wb, alloc, tooled, intro, inp, mech_sets, opt_sets, ok)
    print(f'  Allocation sheet: {n_rows} rows written')

    write_report_sheet(wb, alloc, tooled, intro, mech_fams, opt_fams,
                       mech_sets, opt_sets, n_lines_opened, n_late_intros,
                       inp, ok)
    print('  Report sheet: written')

    tooling_records = compute_tooling_ids(
        alloc, tooled, intro, mech_fams, opt_fams, years)
    write_summary_csv(alloc, tooling_records, inp, intro,
                      str(outdir / 'tooling_summary.csv'))
    plot_gantt(alloc, inp, intro, str(outdir / 'line_gantt.png'))

    print(f'\nSaving {path} ...')
    wb.save(str(path))
    print('Done!')


if __name__ == '__main__':
    main()
