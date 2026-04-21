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

# Clubs currently owned by managers in the league. Highlighted across sheets.
OWNED_TEAMS = {
    # Liga 1 (13 managers)
    'Aston Villa', 'Como', 'Crvena Zvezda', 'Crystal Palace', 'Fulham',
    'Olympique Lyonnais', "Borussia M'gladbach", 'Fortuna D�sseldorf',
    'Roma', 'Slavia Praha', 'Torino', 'Bochum', 'Villarreal',
    # Liga 2 (14 managers)
    'Milan', 'Atalanta', 'Athletic', 'Bayer Leverkusen', 'Leipzig',
    'Chelsea', 'Fenerbahce', 'Juventus', 'Napoli', 'Newcastle',
    'Nottingham Forest', 'Olympique Marseille', 'Betis', 'Schalke',
}

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

    print(f"Saving {OUT}...")
    wb.save(OUT)
    print(f"Done. Sheets: {wb.sheetnames}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
