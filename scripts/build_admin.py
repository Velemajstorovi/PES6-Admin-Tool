"""Build outputs/pes6_admin_tool.xlsx in a single pass.

Consolidates four earlier steps (build_admin_v2.py, add_remaining_sheets.py,
make_live_sheets.py, final_fix.py) into one script. All derived sheets
(Clubs, Nations, Team Composition) are written with live Excel formulas from
the start - no intermediate data_only reads, no LibreOffice recalc needed.

Usage:
    python scripts/build_admin.py
"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SCRIPTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPTS_DIR.parent
sys.path.insert(0, str(SCRIPTS_DIR))

from admin_tool_builder import (
    ABILITY_WEIGHTS, CSV_HEADERS, P, PLAYER_COLS, POSITIONS,
    ability_bonus, ability_count, attacking_prowess, b1, ball_winning,
    best_b1, best_position, cond_multiplier, defending_prowess,
    effective_b1, explosive_power, finishing, kicking_power, ovr_formula,
    player_class, pull_formula, speed,
)
from verify_source import verify_source

SRC = REPO_ROOT / 'data' / 'source.xlsx'
OUT = REPO_ROOT / 'outputs' / 'pes6_admin_tool.xlsx'

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

ROWS = 4900  # Players / CSV Paste row capacity

MIN_CLUB_POOL = 11
MIN_NATION_POOL = 16

# Tier thresholds for Top-18 Effective B1 (tighter than raw OVR because of
# ability and condition multipliers applied to Effective B1).
TIER_THRESHOLDS = [
    ('BANNED', 92.5),
    ('A', 89.0),
    ('B', 86.0),
    ('C', 84.0),
    ('D', 82.0),
    ('E', 79.5),
    # F: anything below
]

TIER_COLORS = {
    'BANNED': '000000', 'A': 'C00000', 'B': 'ED7D31', 'C': 'FFC000',
    'D': '70AD47', 'E': '5B9BD5', 'F': '808080',
}
TIER_TEXT = {
    'BANNED': 'FFFFFF', 'A': 'FFFFFF', 'B': 'FFFFFF', 'C': '000000',
    'D': 'FFFFFF', 'E': 'FFFFFF', 'F': 'FFFFFF',
}

# ---------------------------------------------------------------------------
# MANAGER ROSTER / COMPETITIONS / TIER RULES
# ---------------------------------------------------------------------------
# Single source of truth for the draft league. Edit these lists and rebuild
# when roster / competitions / rules change.

# (manager_name, team, league [1 or 2], initial_tier [A/B/C/D])
# Manager names are placeholders - replace with real names.
MANAGERS: list[tuple[str, str, int, str]] = [
    # Liga 1 (13 managers)
    ('M-L1-01', 'Aston Villa',          1, 'C'),
    ('M-L1-02', 'Como',                 1, 'C'),
    ('M-L1-03', 'Crvena Zvezda',        1, 'C'),
    ('M-L1-04', 'Crystal Palace',       1, 'C'),
    ('M-L1-05', 'Fulham',               1, 'C'),
    ('M-L1-06', 'Olympique Lyonnais',   1, 'C'),
    ('M-L1-07', "Borussia M'gladbach",  1, 'C'),
    ('M-L1-08', 'Fortuna D�sseldorf', 1, 'C'),
    ('M-L1-09', 'Roma',                 1, 'C'),
    ('M-L1-10', 'Slavia Praha',         1, 'C'),
    ('M-L1-11', 'Torino',               1, 'C'),
    ('M-L1-12', 'Bochum',               1, 'C'),
    ('M-L1-13', 'Villarreal',           1, 'C'),
    # Liga 2 (14 managers)
    ('M-L2-01', 'Milan',                2, 'C'),
    ('M-L2-02', 'Atalanta',             2, 'C'),
    ('M-L2-03', 'Athletic',             2, 'C'),
    ('M-L2-04', 'Bayer Leverkusen',     2, 'C'),
    ('M-L2-05', 'Leipzig',              2, 'C'),
    ('M-L2-06', 'Chelsea',              2, 'C'),
    ('M-L2-07', 'Fenerbahce',           2, 'C'),
    ('M-L2-08', 'Juventus',             2, 'C'),
    ('M-L2-09', 'Napoli',               2, 'C'),
    ('M-L2-10', 'Newcastle',            2, 'C'),
    ('M-L2-11', 'Nottingham Forest',    2, 'C'),
    ('M-L2-12', 'Olympique Marseille',  2, 'C'),
    ('M-L2-13', 'Betis',                2, 'C'),
    ('M-L2-14', 'Schalke',              2, 'C'),
]

# Derived from MANAGERS - used to highlight owned clubs in Clubs / Team Composition.
OWNED_TEAMS = {team for _, team, _, _ in MANAGERS}

# (competition_name, coefficient, active_this_season)
# Hierarchy: League (1.00) > PL / Cup (primary cups) > Europa / Conference
# (25% less than PL) > Friendly / SP-WC (25% less than Europa).
# Admins toggle Active per season when a competition is actually played.
COMPETITIONS: list[tuple[str, float, bool]] = [
    ('Liga 1',       1.00, True),
    ('Liga 2',       1.00, True),
    ('PL',           0.80, True),
    ('Cup',          0.70, True),
    ('Europa',       0.60, False),
    ('Conference',   0.60, False),
    ('Friendly Cup', 0.45, False),
    ('SP-World Cup', 0.45, False),
]

# Cup result vocab, best -> worst. Used as the dropdown list for Season Results
# result columns and for indexing relative achievement strength.
RESULT_LEVELS = ['Winner', 'Final', 'Semi', 'QF', 'R16', 'Group', 'DNQ']

TIERS = ['A', 'B', 'C', 'D']

# Outcome-based league finishing bands. Order here = worst-to-best index 0..3.
# (Band index 0 = best = Champion zone.)
BAND_LABELS = ['Champion', 'Playoff', 'Mid-safe', 'Playout']

# Tier -> suggested band index, per league. L2 is weaker overall so each tier's
# expected finish is bumped up (floored at Champion = 0).
#   L1: A -> Champion, B -> Playoff, C -> Mid-safe, D -> Playout (direct)
#   L2: A -> Champion, B -> Champion, C -> Playoff, D -> Mid-safe (shifted up 1)
TIER_BAND_BY_LEAGUE = {
    1: {'A': 0, 'B': 1, 'C': 2, 'D': 3},
    2: {'A': 0, 'B': 0, 'C': 1, 'D': 2},
}

# Tier -> minimum (1-based) RESULT_LEVELS index needed in a cup for that cup's
# coefficient to contribute to the bump score. Lower index = better result.
#   B: Winner only                 (index 1)
#   C: Semi or better              (index <= 3)
#   D: R16 or better               (index <= 5)
# Tier A managers cannot be bumped up (already at top).
CUP_BUMP_MAX_IDX_1B = {'B': 1, 'C': 3, 'D': 5}

# Cup bump triggers when SUM of eligible cups' coefficients >= this floor.
# 0.50 means: a single secondary cup (>=0.60) triggers a bump; two tertiary
# cups (0.45 + 0.45 = 0.90) trigger; one tertiary alone does not.
BUMP_SCORE_FLOOR = 0.50

# Season Results "Outcome" dropdown vocab. Auto tier shift:
#   Promoted  -> +1 tier shift (in addition to league delta + cup bump)
#   Relegated -> -1 tier shift
#   Stayed / blank -> no shift
OUTCOME_LEVELS = ['Stayed', 'Promoted', 'Relegated']
OUTCOME_SHIFT = {'Promoted': 1, 'Relegated': -1, 'Stayed': 0}

# Draft lottery weights - weight of getting #1 pick for the NEW tier after
# tier movement. Worse tier (D) has much higher weight (NBA lottery style).
LOTTERY_WEIGHTS = {'A': 1, 'B': 2, 'C': 4, 'D': 8}

# Blank rows to seed in Season Results for admin entry.
SEASON_RESULTS_BLANK_ROWS = 200

# Which CSV column of `1. Paste cvs` holds what.
SRC_CLUB_COL = CSV_HEADERS.index('CLUB')          # 80 (0-indexed)
SRC_NATION_COL = CSV_HEADERS.index('NATIONALITY') # 77

# Source.xlsx's actual header for the club column (differs from our internal
# CSV_HEADERS). Used only when reading via pandas by column name.
SRC_CLUB_HEADER = 'CLUB TEAM'

# ---------------------------------------------------------------------------
# STYLING HELPERS
# ---------------------------------------------------------------------------

BOLD_W = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
HEAD_FILL = PatternFill('solid', start_color='1F4E78')
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
WRAP_TOP = Alignment(wrap_text=True, vertical='top')


def head_row(ws, row: int, n_cols: int) -> None:
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD_W
        cell.fill = HEAD_FILL
        cell.alignment = CENTER


# ---------------------------------------------------------------------------
# FORMULA HELPERS
# ---------------------------------------------------------------------------

def players_range(col_name: str) -> str:
    """Absolute reference to a column of the Players sheet, rows 2..ROWS+1."""
    col = P[col_name]
    return f"Players!${col}$2:${col}${ROWS + 1}"


def top18_formula(value_range: str, group_range: str, group_ref: str) -> str:
    """Average of the top-18 values in value_range where group_range = group_ref.

    Uses N(+range) to coerce empty-string cells to 0 - without this, LARGE and
    SUMPRODUCT return #VALUE! because the Players sheet leaves unused rows as
    empty strings (the ``IF(Name="","",...)`` pattern).
    """
    count = f"MIN(18,COUNTIF({group_range},{group_ref}))"
    large_sum = (
        f"SUMPRODUCT(LARGE(({group_range}={group_ref})"
        f"*N(+{value_range}),ROW($1:$18)))"
    )
    return f"=IFERROR(ROUND({large_sum}/{count},2),0)"


def tier_formula(value_cell: str) -> str:
    """Nested IF returning tier label (BANNED / A-F) from a numeric cell."""
    expr = f'"F"'
    for tier, cutoff in reversed(TIER_THRESHOLDS):
        expr = f'IF({value_cell}>={cutoff},"{tier}",{expr})'
    return f"={expr}"


def apply_tier_coloring(ws, cell_range: str) -> None:
    for tier, fill_color in TIER_COLORS.items():
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(
                operator='equal',
                formula=[f'"{tier}"'],
                fill=PatternFill('solid', start_color=fill_color),
                font=Font(bold=True, color=TIER_TEXT[tier]),
            ),
        )


def apply_effb1_colorscale(ws, cell_range: str) -> None:
    ws.conditional_formatting.add(
        cell_range,
        ColorScaleRule(
            start_type='min', start_color='F8D7DA',
            mid_type='percentile', mid_value=50, mid_color='FFFFFF',
            end_type='max', end_color='375623',
        ),
    )


def tier_legend_row(ws, row: int) -> None:
    ws.cell(row=row, column=1, value='Tiers:').font = Font(bold=True)
    legend = TIER_THRESHOLDS + [('F', None)]
    labels = {'BANNED': '>=92.5', 'A': '89.0+', 'B': '86.0+', 'C': '84.0+',
              'D': '82.0+', 'E': '79.5+', 'F': '<79.5'}
    for i, (tier, _) in enumerate(legend):
        cell = ws.cell(row=row, column=2 + i, value=f"{tier}: {labels[tier]}")
        cell.fill = PatternFill('solid', start_color=TIER_COLORS[tier])
        cell.font = Font(bold=True, color=TIER_TEXT[tier])
        cell.alignment = CENTER


# ---------------------------------------------------------------------------
# SHEET BUILDERS
# ---------------------------------------------------------------------------

def build_readme(ws) -> None:
    ws['A1'] = 'PES6 ADMIN TOOL'
    ws['A1'].font = Font(size=18, bold=True, color='1F4E78')
    ws.merge_cells('A1:C1')

    sections = [
        ('', ''),
        ('WORKFLOW', 'h'),
        ('Step 1:', 'Export player data from PES Fan Editor as CSV.'),
        ('Step 2:', 'Open this workbook, go to sheet "CSV Paste".'),
        ('Step 3:', 'Delete old data (keep row 1 headers). Paste new CSV at A2.'),
        ('Step 4:', 'Press F9 to recalculate. All other sheets refresh automatically.'),
        ('', ''),
        ('KEY METRICS', 'h'),
        ('OVR', 'Max B1 across all 11 positions. Official overall rating.'),
        ('B1 per position', "Position-specific rating using the patch-maker's exact formula."),
        ('Ability Bonus', 'Multiplicative bonus to Best B1 based on special abilities, weighted per position. Max ~15%.'),
        ('COND Multiplier', 'COND 7 -> x1.06, 6 -> x1.03, 5 -> x1.00, 4 -> x0.98, 3 -> x0.96, <=2 -> x0.94.'),
        ('Effective B1', 'Best B1 * (1 + Ability Bonus) * COND Multiplier. The true player-value metric.'),
        ('Class', 'A (>=90), B (84-89), C (78-83), D (<78), based on OVR.'),
        ('Tier (teams)', 'A-F, based on top-18 average Effective B1.'),
        ('', ''),
        ('SHEETS', 'h'),
        ('1. CSV Paste', 'Raw CSV data. Paste new patch exports here.'),
        ('2. Players', 'Per-player analysis (formulas pulling from CSV Paste).'),
        ('3. Clubs', 'Club rankings and tiers (live formulas).'),
        ('4. Nations', 'National team rankings (live formulas).'),
        ('5. Team Composition', 'Class distribution (A/B/C/D counts) per club (live formulas).'),
        ('6. Draft Rules', 'Reference sheet: class/group/tier rules.'),
        ('7. Managers', 'Roster: manager <-> team <-> league <-> initial tier. Source of truth.'),
        ('8. Competitions', 'League, PL, Cup, Europa/Conference, Friendly, SP-World Cup with importance.'),
        ('9. Expectations', 'Per-manager expected league finish (quartile), admin-overridable.'),
        ('10. Season Results', 'ADMIN DATA ENTRY: one row per manager per season with final placements.'),
        ('11. Tier Movement', 'Auto-computed: compares actual vs expected, outputs new tier for next draft.'),
        ('12. Draft Order', 'NBA-style lottery weights by new tier. Admin runs draw, fills Actual Pick.'),
        ('', ''),
        ('TIER MOVEMENT LOGIC', 'h'),
        ('Bands', 'Outcome-based: Champion (top 3), Playoff (4-8), Mid-safe (middle), Playout (bottom 4).'),
        ('Expected (L1)', 'Tier A -> Champion, B -> Playoff, C -> Mid-safe, D -> Playout.'),
        ('Expected (L2)', 'One band stricter: A -> Champion, B -> Champion, C -> Playoff, D -> Mid-safe.'),
        ('League Delta', 'Actual Band vs Expected: better = +1, worse = -1, same = 0.'),
        ('Cup Bump', '+1 if SUM of eligible-cup coefficients >= 0.50 (eligible = Active AND result meets tier threshold: B=Winner, C=Semi+, D=R16+). A: no bump.'),
        ('Coefficients', 'League 1.00, PL 0.80, Cup 0.70, Europa / Conference 0.60, Friendly / SP-World Cup 0.45. Admin-editable in Competitions sheet.'),
        ('Anti-gaming', 'Cup bump is ignored when league delta is negative (winning a friendly does not save a tanked league).'),
        ('Outcome Shift', '+1 if Promoted, -1 if Relegated, 0 if Stayed. Entered by admin in Season Results after playoff / playout KOs resolve.'),
        ('Final Delta', 'League component capped at +-1; Outcome Shift adds on top. Total range: -2 (tanked and relegated) to +2 (dominant and promoted).'),
        ('', ''),
        ('ABILITY WEIGHTS PER POSITION', 'h'),
    ]
    r = 2
    for label, val in sections:
        if val == 'h':
            cell = ws.cell(row=r, column=1, value=label)
            cell.font = Font(size=12, bold=True, color='1F4E78')
            cell.fill = PatternFill('solid', start_color='DDEBF7')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        elif label or val:
            if label:
                ws.cell(row=r, column=1, value=label).font = Font(bold=True)
            if val:
                ws.cell(row=r, column=2, value=val).alignment = WRAP_TOP
        r += 1

    ws.cell(row=r, column=1, value='Position').font = Font(bold=True)
    ws.cell(row=r, column=2, value='Top weighted abilities').font = Font(bold=True)
    head_row(ws, r, 2)
    r += 1
    for pos in POSITIONS:
        weights = sorted(ABILITY_WEIGHTS.get(pos, {}).items(), key=lambda kv: -kv[1])[:6]
        ws.cell(row=r, column=1, value=pos).font = Font(bold=True)
        ws.cell(row=r, column=2, value=', '.join(f"{k} ({v:.3f})" for k, v in weights))
        r += 1

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 110


def build_csv_paste(ws, df_source: pd.DataFrame) -> None:
    for i, header in enumerate(CSV_HEADERS, 1):
        ws.cell(row=1, column=i, value=header)
    head_row(ws, 1, len(CSV_HEADERS))

    for row_idx, row in enumerate(df_source.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row, start=1):
            if col_idx > len(CSV_HEADERS):
                break
            if pd.notna(val):
                if hasattr(val, 'item'):
                    val = val.item()
                ws.cell(row=row_idx, column=col_idx, value=val)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(CSV_HEADERS))}1"
    ws.freeze_panes = 'B2'


def build_players(ws) -> None:
    # Header row
    for i, header in enumerate(PLAYER_COLS, 1):
        ws.cell(row=1, column=i, value=header)
    head_row(ws, 1, len(PLAYER_COLS))

    # Mapping from Players column name to CSV_HEADERS source column.
    csv_to_p_map = {
        'Name': 'NAME', 'Nationality': 'NATIONALITY', 'Club': 'CLUB',
        'Age': 'AGE', 'Foot': 'STRONG FOOT', 'Registered Pos': 'REGISTERED POSITION',
        'ATTACK': 'ATTACK', 'DEFENSE': 'DEFENSE', 'BALANCE': 'BALANCE',
        'STAMINA': 'STAMINA', 'TOP SPEED': 'TOP SPEED', 'ACCELERATION': 'ACCELERATION',
        'RESPONSE': 'RESPONSE', 'AGILITY': 'AGILITY',
        'DRIBBLE ACC': 'DRIBBLE ACCURACY', 'DRIBBLE SPD': 'DRIBBLE SPEED',
        'SHORT PASS ACC': 'SHORT PASS ACCURACY', 'SHORT PASS SPD': 'SHORT PASS SPEED',
        'LONG PASS ACC': 'LONG PASS ACCURACY', 'LONG PASS SPD': 'LONG PASS SPEED',
        'SHOT ACC': 'SHOT ACCURACY', 'SHOT PWR': 'SHOT POWER', 'SHOT TEC': 'SHOT TECHNIQUE',
        'FREE KICK': 'FREE KICK ACCURACY', 'SWERVE': 'SWERVE', 'HEADING': 'HEADING',
        'JUMP': 'JUMP', 'TECHNIQUE': 'TECHNIQUE', 'AGGRESSION': 'AGGRESSION',
        'MENTALITY': 'MENTALITY', 'GOAL KEEPING': 'GOAL KEEPING',
        'TEAM WORK': 'TEAM WORK', 'CONDITION': 'CONDITION / FITNESS',
        'REACTION': 'REACTION', 'PLAYMAKING': 'PLAYMAKING', 'PASSING': 'PASSING',
        'SCORING': 'SCORING', '1-1 SCORING': '1-1 SCORING', 'POST PLAYER': 'POST PLAYER',
        'LINES': 'LINES', 'MIDDLE SHOOTING': 'MIDDLE SHOOTING', 'SIDE': 'SIDE',
        'CENTRE': 'CENTRE', 'PENALTIES': 'PENALTIES', '1-TOUCH PASS': '1-TOUCH PASS',
        'OUTSIDE': 'OUTSIDE', 'MARKING': 'MARKING', 'SLIDING': 'SLIDING',
        'COVERING': 'COVERING', 'D-LINE CONTROL': 'D-LINE CONTROL',
        'PENALTY STOPPER': 'PENALTY STOPPER', '1-ON-1 STOPPER': '1-ON-1 STOPPER',
        'LONG THROW': 'LONG THROW',
    }
    composite_fn = {
        'Attacking Prowess': attacking_prowess, 'Finishing': finishing,
        'Speed': speed, 'Explosive Power': explosive_power,
        'Kicking Power': kicking_power, 'Defending Prowess': defending_prowess,
        'Ball Winning': ball_winning,
    }
    b1_map = {f'{pos} B1': pos for pos in POSITIONS}

    for row_idx in range(2, ROWS + 2):
        for pcol_name in PLAYER_COLS:
            col = P[pcol_name]
            cell = ws[f"{col}{row_idx}"]
            if pcol_name in csv_to_p_map:
                cell.value = pull_formula(pcol_name, csv_to_p_map[pcol_name], row_idx)
            elif pcol_name in composite_fn:
                cell.value = composite_fn[pcol_name](row_idx)
            elif pcol_name in b1_map:
                cell.value = b1(b1_map[pcol_name], row_idx)
            elif pcol_name == 'OVR':
                cell.value = ovr_formula(row_idx)
            elif pcol_name == 'Best B1':
                cell.value = best_b1(row_idx)
            elif pcol_name == 'Best Position':
                cell.value = best_position(row_idx)
            elif pcol_name == 'Ability ★':
                cell.value = ability_count(row_idx)
            elif pcol_name == 'Ability Bonus':
                cell.value = ability_bonus(row_idx)
            elif pcol_name == 'COND Multiplier':
                cell.value = cond_multiplier(row_idx)
            elif pcol_name == 'Effective B1':
                cell.value = effective_b1(row_idx)
            elif pcol_name == 'Class':
                cell.value = player_class(row_idx)
        if row_idx % 1000 == 0:
            print(f"  Players rows: {row_idx}/{ROWS + 1}")

    ws.auto_filter.ref = f"A1:{get_column_letter(len(PLAYER_COLS))}1"
    ws.freeze_panes = 'B2'

    # Number formats
    for col_name in ['OVR', 'Best B1', 'Effective B1'] + [f'{p} B1' for p in POSITIONS]:
        col = P[col_name]
        for row in range(2, ROWS + 2):
            ws[f"{col}{row}"].number_format = '0.00'
    for col_name, fmt in [('Ability Bonus', '0.000'), ('COND Multiplier', '0.00')]:
        col = P[col_name]
        for row in range(2, ROWS + 2):
            ws[f"{col}{row}"].number_format = fmt

    col_widths = {
        'Name': 22, 'Nationality': 14, 'Club': 22, 'Registered Pos': 10,
        'OVR': 7, 'Class': 6, 'Best B1': 9, 'Effective B1': 10,
        'Best Position': 11, 'Ability ★': 7, 'CONDITION': 9,
        'Ability Bonus': 10, 'COND Multiplier': 10,
    }
    for name, w in col_widths.items():
        ws.column_dimensions[P[name]].width = w


def build_clubs(ws, clubs: list[str]) -> None:
    ws['A1'] = 'CLUB RANKINGS - top-18 Effective B1, auto-refreshes from Players'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:K1')
    ws['A2'] = ('All values are live Excel formulas. When a new CSV is pasted '
                'on "CSV Paste", press F9 to recalculate.')
    ws['A2'].font = Font(italic=True, color='808080')

    tier_legend_row(ws, 3)

    headers = ['Rank', 'Club', 'Tier', 'Top-18 Eff B1', 'Top-18 OVR', 'Pool',
               'A (#)', 'B (#)', 'C (#)', 'D (#)', 'Owned']
    for i, h in enumerate(headers, 1):
        ws.cell(row=5, column=i, value=h)
    head_row(ws, 5, len(headers))

    club_rng = players_range('Club')
    effb1_rng = players_range('Effective B1')
    ovr_rng = players_range('OVR')
    class_rng = players_range('Class')

    first_row, last_row = 6, 5 + len(clubs)
    rank_range = f"$D${first_row}:$D${last_row}"

    for idx, club in enumerate(clubs):
        row = first_row + idx
        club_ref = f'B{row}'

        ws.cell(row=row, column=1, value=f'=RANK(D{row},{rank_range})').alignment = CENTER
        club_cell = ws.cell(row=row, column=2, value=club)
        club_cell.alignment = LEFT
        if club in OWNED_TEAMS:
            club_cell.font = Font(bold=True)
            club_cell.fill = PatternFill('solid', start_color='FFF2CC')

        ws.cell(row=row, column=3, value=tier_formula(f'D{row}')).alignment = CENTER
        v = ws.cell(row=row, column=4, value=top18_formula(effb1_rng, club_rng, club_ref))
        v.number_format = '0.00'
        v.alignment = CENTER
        v = ws.cell(row=row, column=5, value=top18_formula(ovr_rng, club_rng, club_ref))
        v.number_format = '0.00'
        v.alignment = CENTER
        ws.cell(row=row, column=6, value=f'=COUNTIF({club_rng},{club_ref})').alignment = CENTER
        for i, cls in enumerate(['A', 'B', 'C', 'D']):
            ws.cell(row=row, column=7 + i,
                    value=f'=COUNTIFS({club_rng},{club_ref},{class_rng},"{cls}")').alignment = CENTER
        ws.cell(row=row, column=11,
                value='YES' if club in OWNED_TEAMS else '').alignment = CENTER

    apply_tier_coloring(ws, f'C{first_row}:C{last_row}')
    apply_effb1_colorscale(ws, f'D{first_row}:D{last_row}')

    ws.auto_filter.ref = f"A5:{get_column_letter(len(headers))}{last_row}"
    ws.freeze_panes = 'C6'
    for i, w in enumerate([6, 28, 9, 14, 14, 7, 7, 7, 7, 7, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_nations(ws, nations: list[str]) -> None:
    ws['A1'] = 'NATIONAL TEAM RANKINGS - top-18 Effective B1 (>=16 players)'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:G1')

    tier_legend_row(ws, 3)

    headers = ['Rank', 'Nation', 'Tier', 'Top-18 Eff B1', 'Top-18 OVR', 'Pool']
    for i, h in enumerate(headers, 1):
        ws.cell(row=5, column=i, value=h)
    head_row(ws, 5, len(headers))

    nat_rng = players_range('Nationality')
    effb1_rng = players_range('Effective B1')
    ovr_rng = players_range('OVR')

    first_row, last_row = 6, 5 + len(nations)
    rank_range = f"$D${first_row}:$D${last_row}"

    for idx, nat in enumerate(nations):
        row = first_row + idx
        nat_ref = f'B{row}'

        ws.cell(row=row, column=1, value=f'=RANK(D{row},{rank_range})').alignment = CENTER
        ws.cell(row=row, column=2, value=nat).alignment = LEFT
        ws.cell(row=row, column=3, value=tier_formula(f'D{row}')).alignment = CENTER
        v = ws.cell(row=row, column=4, value=top18_formula(effb1_rng, nat_rng, nat_ref))
        v.number_format = '0.00'
        v.alignment = CENTER
        v = ws.cell(row=row, column=5, value=top18_formula(ovr_rng, nat_rng, nat_ref))
        v.number_format = '0.00'
        v.alignment = CENTER
        ws.cell(row=row, column=6, value=f'=COUNTIF({nat_rng},{nat_ref})').alignment = CENTER

    apply_tier_coloring(ws, f'C{first_row}:C{last_row}')
    apply_effb1_colorscale(ws, f'D{first_row}:D{last_row}')

    ws.auto_filter.ref = f"A5:{get_column_letter(len(headers))}{last_row}"
    ws.freeze_panes = 'C6'
    for i, w in enumerate([6, 26, 9, 14, 14, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_team_composition(ws, clubs: list[str]) -> None:
    ws['A1'] = 'TEAM COMPOSITION - class breakdown per club (auto-refresh)'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:L1')
    ws['A2'] = ("For admin use: audit draft picks, verify manager A-caps, "
                "edit teams after draft.")
    ws['A2'].font = Font(italic=True, color='808080')

    headers = ['Rank', 'Club', 'Tier', 'Top-18 Eff B1', 'Total',
               'A (90+)', 'B (84-89)', 'C (78-83)', 'D (<78)', 'A%', 'B%', 'Owned']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    club_rng = players_range('Club')
    effb1_rng = players_range('Effective B1')
    class_rng = players_range('Class')

    first_row, last_row = 5, 4 + len(clubs)
    rank_range = f"$D${first_row}:$D${last_row}"

    for idx, club in enumerate(clubs):
        row = first_row + idx
        club_ref = f'B{row}'

        ws.cell(row=row, column=1, value=f'=RANK(D{row},{rank_range})').alignment = CENTER
        club_cell = ws.cell(row=row, column=2, value=club)
        if club in OWNED_TEAMS:
            club_cell.font = Font(bold=True)
            club_cell.fill = PatternFill('solid', start_color='FFF2CC')
        ws.cell(row=row, column=3, value=tier_formula(f'D{row}')).alignment = CENTER
        v = ws.cell(row=row, column=4, value=top18_formula(effb1_rng, club_rng, club_ref))
        v.number_format = '0.00'
        v.alignment = CENTER
        ws.cell(row=row, column=5, value=f'=COUNTIF({club_rng},{club_ref})').alignment = CENTER
        for i, cls in enumerate(['A', 'B', 'C', 'D']):
            ws.cell(row=row, column=6 + i,
                    value=f'=COUNTIFS({club_rng},{club_ref},{class_rng},"{cls}")').alignment = CENTER
        ws.cell(row=row, column=10, value=f'=IFERROR(F{row}/E{row},0)').number_format = '0%'
        ws.cell(row=row, column=10).alignment = CENTER
        ws.cell(row=row, column=11, value=f'=IFERROR(G{row}/E{row},0)').number_format = '0%'
        ws.cell(row=row, column=11).alignment = CENTER
        ws.cell(row=row, column=12,
                value='YES' if club in OWNED_TEAMS else '').alignment = CENTER

    apply_tier_coloring(ws, f'C{first_row}:C{last_row}')
    # Highlight A-class counts >= 3 (admin attention)
    ws.conditional_formatting.add(
        f'F{first_row}:F{last_row}',
        CellIsRule(operator='greaterThanOrEqual', formula=['3'],
                   fill=PatternFill('solid', start_color='FFE699'),
                   font=Font(bold=True, color='C00000')),
    )

    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{last_row}"
    ws.freeze_panes = 'C5'
    for i, w in enumerate([6, 28, 9, 14, 8, 9, 10, 10, 9, 7, 7, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_draft_rules(ws) -> None:
    ws['A1'] = 'DRAFT RULES & SYSTEM REFERENCE'
    ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
    ws.merge_cells('A1:D1')

    rules = [
        ('', ''),
        ('PLAYER CLASSES (based on OVR)', 'h'),
        ('A', 'OVR 90-94'),
        ('B', 'OVR 84-89'),
        ('C', 'OVR 78-83'),
        ('D', 'OVR <= 77'),
        ('Banned', 'OVR >= 95 (cannot be drafted)'),
        ('', ''),
        ('MANAGER GROUPS', 'h'),
        ('Group A (top)', 'Max 1 A lifetime | Per season: 1A + 2B + 1C + 1D'),
        ('Group B (mid-strong)', 'Max 2 A lifetime | Per season: 2A + 1B + 1C + 1D'),
        ('Group C (mid)', 'Max 3 A lifetime | Per season: 2A + 2B + 1 (C or D)'),
        ('Group D (new)', 'Max 3 A lifetime | Per season: 2A + 3 (B/C/D free)'),
        ('', ''),
        ('POSITION RULE', 'h'),
        ('', 'No 2 A-class players in the same color position '
             '(e.g., not 2 A-class CFs).'),
        ('', 'Color groups: BLUE = defenders, GREEN = midfielders, '
             'RED = attackers, GK.'),
        ('', ''),
        ('BONUSES FOR WEAK TEAMS', 'h'),
        ('Tier F team', '+1 A + 1 B bonus player'),
        ('Tier E team', '+2 B bonus players'),
        ('', ''),
        ('PROMOTION / RELEGATION', 'h'),
        ('', 'After each season, managers move between groups by success. '
             'Best -> harder group.'),
        ('', "If a manager's A-count exceeds their new group's cap, "
             'swap an A -> B at season end.'),
        ('', ''),
        ('TEAM TIER THRESHOLDS (Top-18 Effective B1)', 'h'),
        ('BANNED', '>= 92.5'),
        ('A', '>= 89.0'),
        ('B', '>= 86.0'),
        ('C', '>= 84.0'),
        ('D', '>= 82.0'),
        ('E', '>= 79.5'),
        ('F', '< 79.5'),
    ]

    r = 2
    for label, val in rules:
        if val == 'h':
            cell = ws.cell(row=r, column=1, value=label)
            cell.font = Font(size=12, bold=True, color='1F4E78')
            cell.fill = PatternFill('solid', start_color='DDEBF7')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        elif label or val:
            if label:
                ws.cell(row=r, column=1, value=label).font = Font(bold=True)
            if val:
                ws.cell(row=r, column=2, value=val).alignment = WRAP_TOP
        r += 1

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 90


def build_managers(ws) -> None:
    ws['A1'] = 'MANAGER ROSTER'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:E1')
    ws['A2'] = ("Source of truth for manager <-> team assignments. "
                "To change roster, edit the MANAGERS constant in "
                "scripts/build_admin.py and rebuild.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:E2')

    headers = ['Manager', 'Team', 'League', 'Initial Tier', 'Active?']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    first_row = 5
    last_row = 4 + len(MANAGERS)
    for idx, (name, team, league, tier) in enumerate(MANAGERS):
        row = first_row + idx
        ws.cell(row=row, column=1, value=name).alignment = LEFT
        team_cell = ws.cell(row=row, column=2, value=team)
        team_cell.alignment = LEFT
        team_cell.font = Font(bold=True)
        team_cell.fill = PatternFill('solid', start_color='FFF2CC')
        ws.cell(row=row, column=3, value=league).alignment = CENTER
        ws.cell(row=row, column=4, value=tier).alignment = CENTER
        ws.cell(row=row, column=5, value='YES').alignment = CENTER

    apply_tier_coloring(ws, f'D{first_row}:D{last_row}')

    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{last_row}"
    ws.freeze_panes = 'A5'
    for col, width in enumerate([14, 28, 8, 14, 10], 1):
        ws.column_dimensions[get_column_letter(col)].width = width


def build_competitions(ws) -> None:
    ws['A1'] = 'COMPETITIONS'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:C1')
    ws['A2'] = ("Coefficient is each cup's weight in the tier-movement cup-bump "
                "score. Set Active to YES when the competition is played this "
                "season - inactive cups contribute 0 to the bump regardless of "
                "result. League rows are here for reference; they drive tier "
                "movement directly, not via the cup-bump score.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:C2')

    headers = ['Competition', 'Coefficient', 'Active This Season?']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    for idx, (name, coef, active) in enumerate(COMPETITIONS):
        row = 5 + idx
        ws.cell(row=row, column=1, value=name).alignment = LEFT
        c = ws.cell(row=row, column=2, value=coef)
        c.alignment = CENTER
        c.number_format = '0.00'
        ws.cell(row=row, column=3,
                value='YES' if active else 'NO').alignment = CENTER

    # Active-flag dropdown so admins can toggle without typos.
    active_dv = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(active_dv)
    active_dv.add(f'C5:C{4 + len(COMPETITIONS)}')

    ws.freeze_panes = 'A5'
    for col, width in enumerate([18, 14, 24], 1):
        ws.column_dimensions[get_column_letter(col)].width = width


def build_expectations(ws) -> None:
    ws['A1'] = 'EXPECTATIONS - per-manager expected league finish'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:F1')
    ws['A2'] = ("Bands are outcome-based: Champion (top 3), Playoff (4-8), "
                "Mid-safe (middle), Playout (bottom 4). Suggested Band uses "
                "both Tier AND League - L2 expectations are one band stricter "
                "than L1 at the same tier (since L2 is overall weaker). "
                "Override Band is admin-typed. Effective Band = Override if "
                "set, otherwise Suggested.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:F2')

    headers = ['Manager', 'Current Tier', 'Current League',
               'Suggested Band', 'Override Band', 'Effective Band']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    n_mgr = len(MANAGERS)
    mgr_first = 5
    mgr_last = 4 + n_mgr
    mgrs_name_rng = f'Managers!$A$5:$A${mgr_last}'
    mgrs_tier_rng = f'Managers!$D$5:$D${mgr_last}'
    mgrs_league_rng = f'Managers!$C$5:$C${mgr_last}'

    # Build the suggested-band IF chain from TIER_BAND_BY_LEAGUE.
    # Outer: IF(league=1, <L1 block>, IF(league=2, <L2 block>, ""))
    # Each block: IF(tier="A", label_A, IF(tier="B", ...))
    def tier_to_label_chain(tier_cell: str, league: int) -> str:
        mapping = TIER_BAND_BY_LEAGUE[league]
        expr = '""'
        for t in reversed(TIERS):
            label = BAND_LABELS[mapping[t]]
            expr = f'IF({tier_cell}="{t}","{label}",{expr})'
        return expr

    for idx, (name, _team, _lg, _tier) in enumerate(MANAGERS):
        row = mgr_first + idx
        mgr_ref = f'A{row}'
        tier_ref = f'B{row}'
        league_ref = f'C{row}'
        sug_ref = f'D{row}'
        ovr_ref = f'E{row}'

        ws.cell(row=row, column=1, value=name).alignment = LEFT

        # Current Tier from Managers
        ws.cell(row=row, column=2, value=(
            f'=IFERROR(INDEX({mgrs_tier_rng},'
            f'MATCH({mgr_ref},{mgrs_name_rng},0)),"")'
        )).alignment = CENTER

        # Current League from Managers
        ws.cell(row=row, column=3, value=(
            f'=IFERROR(INDEX({mgrs_league_rng},'
            f'MATCH({mgr_ref},{mgrs_name_rng},0)),"")'
        )).alignment = CENTER

        # Suggested band: depends on both tier and league
        l1_chain = tier_to_label_chain(tier_ref, 1)
        l2_chain = tier_to_label_chain(tier_ref, 2)
        ws.cell(row=row, column=4, value=(
            f'=IF({league_ref}=1,{l1_chain},'
            f'IF({league_ref}=2,{l2_chain},""))'
        )).alignment = CENTER

        # Override: blank for admin input
        ws.cell(row=row, column=5, value=None).alignment = CENTER

        # Effective
        ws.cell(row=row, column=6, value=(
            f'=IF({ovr_ref}="",{sug_ref},{ovr_ref})'
        )).alignment = CENTER

    # Override accepts BAND_LABELS only
    band_dv = DataValidation(
        type='list',
        formula1='"' + ','.join(BAND_LABELS) + '"',
        allow_blank=True,
    )
    ws.add_data_validation(band_dv)
    band_dv.add(f'E{mgr_first}:E{mgr_last}')

    apply_tier_coloring(ws, f'B{mgr_first}:B{mgr_last}')

    ws.freeze_panes = 'A5'
    for col, width in enumerate([14, 13, 14, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(col)].width = width


def build_season_results(ws) -> None:
    ws['A1'] = 'SEASON RESULTS - admin data entry (append-only log)'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:K1')
    ws['A2'] = ("One row per manager per season. League Finish = final rank "
                "after playoff/playout. Cup columns use the dropdown list. "
                "Outcome = Promoted / Relegated / Stayed (after playoff / "
                "playout KO resolves). Tier Movement auto-uses the LATEST "
                "Season for each manager.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:K2')

    cup_names = [c for c, _, _ in COMPETITIONS if c not in ('Liga 1', 'Liga 2')]
    # Layout: Season | Manager | League Finish | <cup cols...> | Outcome
    headers = ['Season', 'Manager', 'League Finish'] + cup_names + ['Outcome']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    data_first = 5
    data_last = 4 + SEASON_RESULTS_BLANK_ROWS
    mgr_last = 4 + len(MANAGERS)
    outcome_col = 3 + len(cup_names) + 1  # column index for Outcome

    # Manager dropdown from Managers sheet
    mgr_dv = DataValidation(
        type='list',
        formula1=f'=Managers!$A$5:$A${mgr_last}',
        allow_blank=True,
    )
    ws.add_data_validation(mgr_dv)
    mgr_dv.add(f'B{data_first}:B{data_last}')

    # Cup result dropdown from RESULT_LEVELS
    result_dv = DataValidation(
        type='list',
        formula1='"' + ','.join(RESULT_LEVELS) + '"',
        allow_blank=True,
    )
    ws.add_data_validation(result_dv)
    for ci in range(4, 4 + len(cup_names)):
        col = get_column_letter(ci)
        result_dv.add(f'{col}{data_first}:{col}{data_last}')

    # Outcome dropdown
    outcome_dv = DataValidation(
        type='list',
        formula1='"' + ','.join(OUTCOME_LEVELS) + '"',
        allow_blank=True,
    )
    ws.add_data_validation(outcome_dv)
    out_letter = get_column_letter(outcome_col)
    outcome_dv.add(f'{out_letter}{data_first}:{out_letter}{data_last}')

    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{data_last}"
    ws.freeze_panes = 'C5'
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 13
    for ci in range(4, 4 + len(cup_names)):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    ws.column_dimensions[out_letter].width = 12


def build_tier_movement(ws) -> None:
    ws['A1'] = 'TIER MOVEMENT - computed from latest Season Results'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:M1')
    ws['A2'] = ("Uses each manager's MOST RECENT season in Season Results. "
                "League Delta: +1 if Actual Band better than Expected, -1 if "
                "worse, 0 if same. Cup Bump: +1 if SUM of eligible-cup "
                "coefficients >= 0.5 (eligible = Active AND result reaches the "
                "tier threshold: B=Winner, C=Semi+, D=R16+). Cup bump is "
                "suppressed when league delta is negative. League component "
                "capped at +-1. Outcome Shift adds +1 for Promoted / -1 for "
                "Relegated on top of league component; combined Final Delta "
                "can be +-2 in promotion/relegation cases.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:M2')

    headers = ['Manager', 'League', 'Current Tier', 'Latest Season',
               'League Finish', 'League Size', 'Actual Band', 'Expected Band',
               'League Delta', 'Cup Bump', 'Outcome Shift',
               'Final Delta', 'New Tier']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    n_mgr = len(MANAGERS)
    mgr_first = 5
    mgr_last = 4 + n_mgr
    mgrs_name_rng = f'Managers!$A$5:$A${mgr_last}'
    mgrs_tier_rng = f'Managers!$D$5:$D${mgr_last}'
    mgrs_league_rng = f'Managers!$C$5:$C${mgr_last}'
    exp_name_rng = f'Expectations!$A$5:$A${mgr_last}'
    exp_eff_rng = f'Expectations!$F$5:$F${mgr_last}'  # Effective Band is col F now

    sr_first = 5
    sr_last = 4 + SEASON_RESULTS_BLANK_ROWS
    sr_season = f"'Season Results'!$A${sr_first}:$A${sr_last}"
    sr_mgr = f"'Season Results'!$B${sr_first}:$B${sr_last}"
    sr_finish = f"'Season Results'!$C${sr_first}:$C${sr_last}"

    cup_names = [c for c, _, _ in COMPETITIONS if c not in ('Liga 1', 'Liga 2')]
    sr_cup_rngs = []
    for i in range(len(cup_names)):
        col = get_column_letter(4 + i)
        sr_cup_rngs.append(
            f"'Season Results'!${col}${sr_first}:${col}${sr_last}"
        )

    # Outcome column in Season Results (after the cup columns).
    outcome_col_idx = 3 + len(cup_names) + 1
    outcome_col_letter = get_column_letter(outcome_col_idx)
    sr_outcome = f"'Season Results'!${outcome_col_letter}${sr_first}:${outcome_col_letter}${sr_last}"

    # Competitions sheet rows map (for Active + Coefficient lookups).
    comp_row_map = {n: 5 + i for i, (n, _, _) in enumerate(COMPETITIONS)}

    # Array literals for MATCH-based lookups.
    result_levels_literal = '{' + ','.join(f'"{r}"' for r in RESULT_LEVELS) + '}'
    band_labels_literal = '{' + ','.join(f'"{b}"' for b in BAND_LABELS) + '}'

    for idx, (name, _team, _lg, _tier) in enumerate(MANAGERS):
        row = mgr_first + idx
        mgr_ref = f'A{row}'
        league_ref = f'B{row}'
        tier_ref = f'C{row}'
        latest_ref = f'D{row}'
        finish_ref = f'E{row}'
        size_ref = f'F{row}'
        actband_ref = f'G{row}'
        expband_ref = f'H{row}'
        ldelta_ref = f'I{row}'
        bump_ref = f'J{row}'
        oshift_ref = f'K{row}'
        fdelta_ref = f'L{row}'

        ws.cell(row=row, column=1, value=name).alignment = LEFT

        ws.cell(row=row, column=2, value=(
            f'=IFERROR(INDEX({mgrs_league_rng},'
            f'MATCH({mgr_ref},{mgrs_name_rng},0)),"")'
        )).alignment = CENTER

        ws.cell(row=row, column=3, value=(
            f'=IFERROR(INDEX({mgrs_tier_rng},'
            f'MATCH({mgr_ref},{mgrs_name_rng},0)),"")'
        )).alignment = CENTER

        # Latest Season via SUMPRODUCT(MAX((cond)*range)) - portable.
        ws.cell(row=row, column=4, value=(
            f'=IF(COUNTIF({sr_mgr},{mgr_ref})=0,"",'
            f'SUMPRODUCT(MAX(({sr_mgr}={mgr_ref})*{sr_season})))'
        )).alignment = CENTER

        # League Finish for that latest season.
        ws.cell(row=row, column=5, value=(
            f'=IF({latest_ref}="","",'
            f'SUMIFS({sr_finish},{sr_season},{latest_ref},{sr_mgr},{mgr_ref}))'
        )).alignment = CENTER

        # League Size.
        ws.cell(row=row, column=6, value=(
            f'=COUNTIF({mgrs_league_rng},{league_ref})'
        )).alignment = CENTER

        # Actual Band: outcome-based label. Playout wins ties (bottom-4 priority).
        ws.cell(row=row, column=7, value=(
            f'=IF(OR({finish_ref}="",{finish_ref}=0),"",'
            f'IF({finish_ref}>={size_ref}-3,"Playout",'
            f'IF({finish_ref}<=3,"Champion",'
            f'IF({finish_ref}<=8,"Playoff","Mid-safe"))))'
        )).alignment = CENTER

        # Expected Band from Expectations.Effective Band.
        ws.cell(row=row, column=8, value=(
            f'=IFERROR(INDEX({exp_eff_rng},'
            f'MATCH({mgr_ref},{exp_name_rng},0)),"")'
        )).alignment = CENTER

        # League Delta via band-index comparison (lower index = better).
        act_idx = f'MATCH({actband_ref},{band_labels_literal},0)'
        exp_idx = f'MATCH({expband_ref},{band_labels_literal},0)'
        ws.cell(row=row, column=9, value=(
            f'=IF(OR({actband_ref}="",{expband_ref}=""),"",'
            f'IF({act_idx}<{exp_idx},1,'
            f'IF({act_idx}>{exp_idx},-1,0)))'
        )).alignment = CENTER

        # Cup Bump: coefficient-weighted sum across active cups where result
        # meets tier threshold. +1 if sum >= BUMP_SCORE_FLOOR, else 0.
        contrib_terms = []
        for cup_name, cup_rng in zip(cup_names, sr_cup_rngs):
            comp_row = comp_row_map[cup_name]
            active_cell = f'Competitions!$C${comp_row}'
            coef_cell = f'Competitions!$B${comp_row}'
            cell_pull = (
                f'IFERROR(LOOKUP(2,1/(({sr_season}={latest_ref})*'
                f'({sr_mgr}={mgr_ref})),{cup_rng}),"")'
            )
            match_idx = (
                f'IFERROR(MATCH({cell_pull},{result_levels_literal},0),999)'
            )
            # Tier-dependent eligibility (B<=1, C<=3, D<=5, A never).
            b_thr = CUP_BUMP_MAX_IDX_1B['B']
            c_thr = CUP_BUMP_MAX_IDX_1B['C']
            d_thr = CUP_BUMP_MAX_IDX_1B['D']
            eligible = (
                f'IF({tier_ref}="B",{match_idx}<={b_thr},'
                f'IF({tier_ref}="C",{match_idx}<={c_thr},'
                f'IF({tier_ref}="D",{match_idx}<={d_thr},FALSE)))'
            )
            # Contribution = (active=YES) * (eligible) * coefficient.
            contrib = f'(({active_cell}="YES")*({eligible})*{coef_cell})'
            contrib_terms.append(contrib)
        bump_score = '(' + '+'.join(contrib_terms) + ')'
        ws.cell(row=row, column=10, value=(
            f'=IF({latest_ref}="",0,'
            f'IF({bump_score}>={BUMP_SCORE_FLOOR},1,0))'
        )).alignment = CENTER

        # Outcome Shift from Season Results.Outcome for this season.
        outcome_pull = (
            f'IFERROR(LOOKUP(2,1/(({sr_season}={latest_ref})*'
            f'({sr_mgr}={mgr_ref})),{sr_outcome}),"")'
        )
        ws.cell(row=row, column=11, value=(
            f'=IF({latest_ref}="",0,'
            f'IF({outcome_pull}="Promoted",1,'
            f'IF({outcome_pull}="Relegated",-1,0)))'
        )).alignment = CENTER

        # Final Delta: clamp(league_delta + cup_bump_if_eligible, -1, +1)
        # THEN add outcome shift (total can be +-2 for promotion/relegation).
        ws.cell(row=row, column=12, value=(
            f'=IF({ldelta_ref}="",{oshift_ref},'
            f'MAX(-1,MIN(1,{ldelta_ref}+'
            f'IF({ldelta_ref}<0,0,{bump_ref})))+{oshift_ref})'
        )).alignment = CENTER

        # New Tier: A=1..D=4; +1 delta = tier UP (lower index). Clamp 1..4.
        old_idx = (
            f'IF({tier_ref}="A",1,IF({tier_ref}="B",2,'
            f'IF({tier_ref}="C",3,IF({tier_ref}="D",4,0))))'
        )
        new_idx = f'MAX(1,MIN(4,{old_idx}-{fdelta_ref}))'
        ws.cell(row=row, column=13, value=(
            f'=IF(AND({fdelta_ref}=0,{tier_ref}<>""),{tier_ref},'
            f'IF({tier_ref}="","",'
            f'CHOOSE({new_idx},"A","B","C","D")))'
        )).alignment = CENTER

    apply_tier_coloring(ws, f'C{mgr_first}:C{mgr_last}')
    apply_tier_coloring(ws, f'M{mgr_first}:M{mgr_last}')

    ws.freeze_panes = 'A5'
    widths = [14, 8, 10, 13, 13, 11, 12, 13, 12, 10, 13, 12, 10]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def build_draft_order(ws) -> None:
    ws['A1'] = 'DRAFT ORDER - NBA-style lottery weights by new tier'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:G1')
    ws['A2'] = ("Lower tier = better draft odds. Lottery Weight is each manager's "
                "share of the #1-pick probability. Reverse Rank shows the straight "
                "reverse-order pick (ties = same rank). Run the draw externally "
                "and record the outcome in Actual Pick.")
    ws['A2'].font = Font(italic=True, color='808080')
    ws.merge_cells('A2:G2')

    headers = ['Manager', 'Team', 'League', 'New Tier',
               'Lottery Weight', 'Reverse Rank', 'Actual Pick']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    head_row(ws, 4, len(headers))

    n_mgr = len(MANAGERS)
    mgr_first = 5
    mgr_last = 4 + n_mgr
    tm_name_rng = f"'Tier Movement'!$A$5:$A${mgr_last}"
    tm_newtier_rng = f"'Tier Movement'!$M$5:$M${mgr_last}"

    # Build nested IF for lottery weight: IF(tier="A",w_A, IF(tier="B",w_B, ...))
    def weight_formula(tier_cell: str) -> str:
        expr = '0'
        for tier in reversed(TIERS):
            expr = f'IF({tier_cell}="{tier}",{LOTTERY_WEIGHTS[tier]},{expr})'
        return expr

    weight_range = f'$E${mgr_first}:$E${mgr_last}'

    for idx, (name, team, league, _tier) in enumerate(MANAGERS):
        row = mgr_first + idx
        mgr_ref = f'A{row}'
        tier_ref = f'D{row}'
        weight_ref = f'E{row}'

        ws.cell(row=row, column=1, value=name).alignment = LEFT
        team_cell = ws.cell(row=row, column=2, value=team)
        team_cell.alignment = LEFT
        team_cell.font = Font(bold=True)
        team_cell.fill = PatternFill('solid', start_color='FFF2CC')
        ws.cell(row=row, column=3, value=league).alignment = CENTER

        ws.cell(row=row, column=4, value=(
            f'=IFERROR(INDEX({tm_newtier_rng},'
            f'MATCH({mgr_ref},{tm_name_rng},0)),"")'
        )).alignment = CENTER

        ws.cell(row=row, column=5, value=(
            f'=IF({tier_ref}="","",{weight_formula(tier_ref)})'
        )).alignment = CENTER

        ws.cell(row=row, column=6, value=(
            f'=IF({tier_ref}="","",RANK({weight_ref},{weight_range}))'
        )).alignment = CENTER

        ws.cell(row=row, column=7, value=None).alignment = CENTER

    apply_tier_coloring(ws, f'D{mgr_first}:D{mgr_last}')

    ws.freeze_panes = 'A5'
    widths = [14, 26, 8, 10, 14, 13, 12]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def _extract_clubs_and_nations(df_source: pd.DataFrame) -> tuple[list[str], list[str]]:
    """Return (clubs, nations) sorted lists, filtered by minimum pool size."""
    club_series = df_source.iloc[:, SRC_CLUB_COL].dropna()
    nation_series = df_source.iloc[:, SRC_NATION_COL].dropna()

    club_counts = Counter(
        str(c) for c in club_series if str(c) != '0' and str(c).strip()
    )
    nation_counts = Counter(
        str(n) for n in nation_series if str(n) != 'Free Nationality' and str(n).strip()
    )

    clubs = sorted(c for c, n in club_counts.items() if n >= MIN_CLUB_POOL)
    nations = sorted(n for n, count in nation_counts.items() if count >= MIN_NATION_POOL)
    return clubs, nations


def main() -> int:
    print(f"Verifying {SRC}...")
    verify_source(SRC)

    OUT.parent.mkdir(parents=True, exist_ok=True)

    print(f"Loading {SRC}...")
    df_source = pd.read_excel(SRC, sheet_name='1. Paste cvs', header=0)
    print(f"  {len(df_source)} rows loaded.")

    clubs, nations = _extract_clubs_and_nations(df_source)
    print(f"  {len(clubs)} clubs (pool >= {MIN_CLUB_POOL}), "
          f"{len(nations)} nations (pool >= {MIN_NATION_POOL}).")

    wb = Workbook()
    build_readme(wb.active)
    wb.active.title = 'README'

    print("Building CSV Paste...")
    build_csv_paste(wb.create_sheet('CSV Paste'), df_source)

    print("Building Players (formulas)...")
    build_players(wb.create_sheet('Players'))

    print(f"Building Clubs ({len(clubs)} rows)...")
    build_clubs(wb.create_sheet('Clubs'), clubs)

    print(f"Building Nations ({len(nations)} rows)...")
    build_nations(wb.create_sheet('Nations'), nations)

    print(f"Building Team Composition ({len(clubs)} rows)...")
    build_team_composition(wb.create_sheet('Team Composition'), clubs)

    print("Building Draft Rules...")
    build_draft_rules(wb.create_sheet('Draft Rules'))

    print(f"Building Managers ({len(MANAGERS)} rows)...")
    build_managers(wb.create_sheet('Managers'))

    print(f"Building Competitions ({len(COMPETITIONS)} rows)...")
    build_competitions(wb.create_sheet('Competitions'))

    print("Building Expectations...")
    build_expectations(wb.create_sheet('Expectations'))

    print(f"Building Season Results ({SEASON_RESULTS_BLANK_ROWS} blank rows)...")
    build_season_results(wb.create_sheet('Season Results'))

    print("Building Tier Movement...")
    build_tier_movement(wb.create_sheet('Tier Movement'))

    print("Building Draft Order...")
    build_draft_order(wb.create_sheet('Draft Order'))

    print(f"Saving {OUT}...")
    wb.save(OUT)
    print(f"Done. Sheets: {wb.sheetnames}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
