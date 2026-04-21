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

`build_admin.py` and `pes6_tracker_build.py` call `verify_source()` on
startup, so a structurally-different file fails fast with a message listing the
mismatched sheets/columns. If the patch format legitimately changed, update
`SHEET_SPECS` (and `EXPECTED_CVS_HEADERS`) in `scripts/verify_source.py`.

## Tier-list comparison

There's no dedicated script for this yet. When you want to diff a human tier
list against a data-driven one (as was done for Phoenix 2026), either build a
`scripts/pes6_tier_comparison.py` or ask Claude Code to run the comparison
directly against `data/source.xlsx`. Methodology notes worth preserving:

- **Primary metric:** top-16 average OVR per club (starting XI + 5 rotation).
  Secondary validation metric: top-16 best-B1.
- **Tiers:** BANNED, A, B, C, D, E, F. Prefer natural gaps in the OVR
  distribution as boundaries; BANNED threshold is typically top-16 OVR
  `>= 89.0`.
- **BANNED caveat:** human lists often include "star-power bans" (Messi,
  Ronaldo, one-star super-clubs) that don't show up in depth-based data
  tiers. Flag these as a legitimate difference, not an error.
- **Comparison outputs:** full side-by-side, mismatches-only, major
  disagreements (`>= 2` tier diff), and accuracy headline (exact-match %,
  within-1-tier %, within-2-tier %).
- **Name-matching quirks:** patch Excel often has mojibake (`Bayern
  M�nchen` -> `Bayern München`, `Atl�tico` -> `Atlético`,
  `D�sseldorf` -> `Düsseldorf`). Short-name aliases used by community
  lists: `Leverkusen` = Bayer Leverkusen, `Inter` = Internazionale,
  `Bayern` = Bayern München, `Sociedad` = Real Sociedad, `PSG` = Paris
  Saint-Germain, `United`/`City` = Manchester United / City, `Dortmund` =
  Borussia Dortmund, `Atletico Madrid` = Atlético de Madrid.
