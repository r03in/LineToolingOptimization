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

# ── Sheet / grid constants (match Book.xlsx layout) ───────────────────────────
SHEET_NAME   = 'Blad1'
NUM_PRODUCTS = 10
NUM_LINES    = 15
DEMAND_START = 3
DEMAND_END   = 19
CT_START     = 31   # cycle-time matrix: row CT_START + line_idx, col 3 + prod_idx
OEE_START    = 49   # OEE matrix
MECH_START   = 70   # mechanical tooling compatibility matrix
OPT_START    = 84   # optical  tooling compatibility matrix
OUTPUT_START = 98   # first data row of output grid

# ── Solver configuration (edit these to change behaviour) ─────────────────────
BASE_OEE               = 0.85   # OEE for single-product lines (default)
CHANGEOVER_OEE_PENALTY = 0.03   # OEE reduction for lines running 2+ products
W_LINES                = 10_000 # objective weight: minimise open lines
W_TOOLING              = 100    # objective weight: minimise tooling sets
W_SWITCHES             = 1      # objective weight: minimise product-line changes
SOLVER_TIME_LIMIT      = 300    # CBC wall-clock limit in seconds

# ── Product colours for Gantt chart ───────────────────────────────────────────
PRODUCT_COLORS = [
    '#4C72B0', '#DD8452', '#55A868', '#C44E52',
    '#8172B3', '#937860', '#DA8BC3', '#8C8C8C',
    '#CCB974', '#64B5CD',
]


# ─────────────────────────────────────────────────────────────────────────────
# INPUT LOADING
# ─────────────────────────────────────────────────────────────────────────────
def load_inputs(wb):
    """Read all inputs from the Blad1 sheet and return a dict."""
    ws = wb[SHEET_NAME]
    setup = {
        'hours':  float(ws['B22'].value or 0),
        'shifts': float(ws['B23'].value or 0),
        'days':   float(ws['B24'].value or 0),
        'weeks':  float(ws['B25'].value or 0),
    }
    avail = setup['hours'] * setup['shifts'] * setup['days'] * setup['weeks'] * 3600

    years, demand = [], {}
    for r in range(DEMAND_START, DEMAND_END + 1):
        yr = ws.cell(row=r, column=2).value
        if yr is None:
            continue
        yr = int(yr)
        years.append(yr)
        demand[yr] = [int(ws.cell(row=r, column=c).value or 0) for c in range(3, 13)]

    ct  = [[ws.cell(row=CT_START  + l, column=3 + p).value or 12   for p in range(10)] for l in range(15)]
    oee = [[ws.cell(row=OEE_START + l, column=3 + p).value or 0.85 for p in range(10)] for l in range(15)]
    mt  = [[ws.cell(row=MECH_START + i, column=3 + j).value or 0   for j in range(10)] for i in range(10)]
    ot  = [[ws.cell(row=OPT_START  + i, column=3 + j).value or 0   for j in range(10)] for i in range(10)]

    return {'years': years, 'demand': demand, 'avail': avail,
            'ct': ct, 'oee': oee, 'mt': mt, 'ot': ot}


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

    # chg[p,l,i]  binary: 1 if u[p,l,i] changed from u[p,l,i-1]  (i >= 1)
    chg = pulp.LpVariable.dicts(
        'chg',
        [(p, l, i) for p in range(P) for l in range(L) for i in range(1, Y)],
        cat='Binary')

    # ── Objective ─────────────────────────────────────────────────────────────
    prob += (
        W_LINES   * pulp.lpSum(o[l, i]    for l in range(L) for i in range(Y))
      + W_TOOLING * pulp.lpSum(tm[f, l]   for f in range(MF) for l in range(L))
      + W_TOOLING * pulp.lpSum(to_[f, l]  for f in range(OF) for l in range(L))
      + W_SWITCHES * pulp.lpSum(chg[p, l, i]
                                for p in range(P) for l in range(L)
                                for i in range(1, Y))
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

    # 9. Switch detection: chg = 1 when u[p,l,i] differs from u[p,l,i-1]
    for p in range(P):
        for l in range(L):
            for i in range(1, Y):
                prob += (chg[p, l, i] >= u[p, l, i] - u[p, l, i - 1],
                         f'chgA_{p}_{l}_{i}')
                prob += (chg[p, l, i] >= u[p, l, i - 1] - u[p, l, i],
                         f'chgB_{p}_{l}_{i}')

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

    return alloc, tooled, intro, mech_sets, opt_sets, mech_fams, opt_fams


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
# EXCEL OUTPUT  (unchanged grid format for downstream compatibility)
# ─────────────────────────────────────────────────────────────────────────────
def write_output(ws, alloc, years):
    # Clear grid first
    for r in range(OUTPUT_START, OUTPUT_START + len(years)):
        for c in range(3, 3 + NUM_LINES * NUM_PRODUCTS):
            ws.cell(row=r, column=c).value = None
    n = 0
    for yr in years:
        row = OUTPUT_START + (yr - years[0])
        for line in range(NUM_LINES):
            for prod in range(NUM_PRODUCTS):
                v = alloc.get((yr, line, prod), 0)
                if v > 0:
                    ws.cell(row=row, column=3 + line * 10 + prod).value = v
                    n += 1
    return n


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
    parser.add_argument('workbook', nargs='?', default='Book.xlsx')
    args = parser.parse_args()

    path = Path(args.workbook)
    if not path.exists():
        print(f'Error: {path} not found')
        sys.exit(1)

    print(f'Loading {path} ...')
    wb  = openpyxl.load_workbook(str(path))
    inp = load_inputs(wb)

    years = inp['years']
    print(f'  Available seconds/year : {inp["avail"]:,.0f}')
    print(f'  Years                  : {years[0]}-{years[-1]}  ({len(years)} years)')
    print(f'  BASE_OEE               : {BASE_OEE*100:.0f}%')
    print(f'  CHANGEOVER_OEE_PENALTY : {CHANGEOVER_OEE_PENALTY*100:.0f}%')

    # Quick sanity: show tooling family counts
    mf_preview = compute_families(inp['mt'], NUM_PRODUCTS)
    of_preview = compute_families(inp['ot'], NUM_PRODUCTS)
    print(f'  Mech tooling families  : {len(mf_preview)}'
          f'  (sizes: {sorted(len(f) for f in mf_preview)})')
    print(f'  Opt  tooling families  : {len(of_preview)}'
          f'  (sizes: {sorted(len(f) for f in of_preview)})')

    print('\nBuilding & solving ILP ...')
    result = build_and_solve(inp)
    if result is None:
        print('Solver failed — no solution found.')
        sys.exit(1)

    alloc, tooled, intro, mech_sets, opt_sets, mech_fams, opt_fams = result

    print('\nVerifying solution ...')
    ok = verify(alloc, inp)
    print(f'  {"All checks passed" if ok else "Issues found — see above"}')

    total_sets = mech_sets + opt_sets
    print(f'\nTooling: {mech_sets} mech + {opt_sets} opt = {total_sets} sets')
    print(f'  (greedy baseline: 19 mech + 19 opt = 38 sets  |  theory min: 14+14=28)')

    print('\nLines used:')
    for l in sorted(intro.keys()):
        ps     = sorted(f'P{p+1}' for p in tooled[l])
        n_prod = len(ps)
        oee    = (BASE_OEE - CHANGEOVER_OEE_PENALTY if n_prod > 1 else BASE_OEE) * 100
        print(f'  Line {l+1:2d}  intro={intro[l]}  products={ps}'
              f'  effective_OEE={oee:.0f}%')

    print('\nPer-year summary:')
    for yr in years:
        d = inp['demand'].get(yr, [0] * 10)
        if sum(d) == 0:
            continue
        active_lines = sorted({l for (y2, l, _) in alloc if y2 == yr})
        parts = []
        for l in active_lines:
            prods = [f'P{p+1}' for p in range(NUM_PRODUCTS)
                     if alloc.get((yr, l, p), 0) > 0]
            parts.append(f'L{l+1}=[{",".join(prods)}]')
        print(f'  {yr}: {", ".join(parts)}  —  {len(active_lines)} line(s)')

    # Build tooling ID records
    tooling_records = compute_tooling_ids(
        alloc, tooled, intro, mech_fams, opt_fams, years)

    print('\nGenerating outputs ...')
    outdir = path.parent

    ws = wb[SHEET_NAME]
    n_cells = write_output(ws, alloc, years)
    print(f'  Excel grid: {n_cells} cells written')

    write_summary_csv(alloc, tooling_records, inp, intro,
                      str(outdir / 'tooling_summary.csv'))
    plot_gantt(alloc, inp, intro, str(outdir / 'line_gantt.png'))

    print(f'\nSaving {path} ...')
    wb.save(str(path))
    print('Done!')


if __name__ == '__main__':
    main()
