# PES6-Admin-Tool

Tooling for running a PES6 draft league: reverse-engineers the patch's
player-value (B1) formulas, ranks clubs / nations by top-18 effective B1,
auto-tiers them A-F, and tracks per-team class distribution for draft rules.

Main deliverable is `outputs/pes6_admin_tool.xlsx`, a self-refreshing workbook
with ~390k native Excel formulas. Paste a fresh CSV export from PES Fan Editor
onto the "CSV Paste" sheet and everything recalculates.

## Repo layout

```
.
├── data/
│   └── source.xlsx           # patch Excel (Phoenix 2026) - tracked in git
├── scripts/
│   ├── admin_tool_builder.py # shared config + formula builders (library)
│   ├── build_admin.py        # single-pass builder for the admin workbook
│   ├── pes6_tracker_build.py # standalone over-performer tracker (ridge)
│   └── verify_source.py      # schema guardrail for data/source.xlsx
├── prompts/
│   ├── REUSABLE_PROMPT.md
│   └── REUSABLE_TIER_COMPARISON_PROMPT.md
├── outputs/                  # generated workbooks (gitignored)
├── requirements.txt
└── .gitignore
```

## Setup

```bash
pip install -r requirements.txt
```

## Rebuilding the admin workbook

```bash
python scripts/build_admin.py
```

Output lands in `outputs/pes6_admin_tool.xlsx`. Everything - Players, Clubs,
Nations, Team Composition - is driven by live Excel formulas, so re-pasting a
new CSV on the "CSV Paste" sheet refreshes all downstream views. Press F9 in
Excel / LibreOffice if auto-recalculation is disabled.

## Regenerating the over-performer tracker

```bash
python scripts/pes6_tracker_build.py data/source.xlsx outputs/tracker.xlsx
```

Optional: `--owned-teams owned.txt` to exclude a custom roster list.

## Source file guardrail

`data/source.xlsx` must match the schema the scripts expect. Before running any
script that reads it, call the validator:

```bash
python scripts/verify_source.py                 # checks data/source.xlsx
python scripts/verify_source.py other_patch.xlsx
```

`build_admin_v2.py` and `pes6_tracker_build.py` call `verify_source()` on
startup, so a structurally-different file fails fast with a message listing the
mismatched sheets/columns. If the patch format legitimately changed, update
`SHEET_SPECS` (and `EXPECTED_CVS_HEADERS`) in `scripts/verify_source.py`.

## Reusable prompts

`prompts/REUSABLE_PROMPT.md` - regenerate the draft tracker on a new patch.
`prompts/REUSABLE_TIER_COMPARISON_PROMPT.md` - compare a human tier list vs a
data-driven one.
