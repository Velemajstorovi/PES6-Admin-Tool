# Reusable prompt — PES6 over-performer draft tracker

Copy everything between the `---` lines into a fresh Claude chat, **edit the bracketed values** for your current draft, and attach the new patch Excel file.

---

I'm playing PES6 with the **[PATCH NAME]** patch (Phoenix/Shollym-style, slow/balanced gameplay) and need a comprehensive over-performer draft tracker built from the attached Excel file.

## Draft context (edit before sending)

- My team: **[YOUR TEAM]**
- Banned OVR: **≥ 95** (cannot be drafted)
- OVR classes: A = **90–94**, B = **84–89**, C = **79–83**, D = **≤ 78**
- Draft slots per season: **3×A + 2×B + 1×C + 1×D** (7 total)
- Formation target: **3-4-2-1** (alt: 3-5-2)
- Owned teams to exclude from the draft pool (my opponents' rosters):
  - Liga 1: **[comma-separated team names]**
  - Liga 2: **[comma-separated team names]**

## Work to do

1. Load the attached Excel. Three sheets matter: main data (usually `3. Paste (Value Only)`), training data (usually `Calc`) which has B1 columns, and abilities data (usually `1. Paste cvs`). If sheet names differ, detect them by content.

2. Fit ridge regression (α = 0.001) per position using the Calc sheet as training data. Target R² ≥ 0.9995 for all 11 position formulas: **GK, CB/CWP, SB, WB, DMF, CMF, SMF, AMF, WF, SS, CF**. Use the 26 standard attributes as features (ATTACK, DEFENSE, BALANCE, STAMINA, TOP SPEED, ACCELERATION, RESPONSE, AGILITY, DRIBBLE ACCURACY, DRIBBLE SPEED, SHORT/LONG PASS ACCURACY+SPEED, SHOT ACCURACY/POWER/TECHNIQUE, FREE KICK ACCURACY, SWERVE, HEADING, JUMP, TECHNIQUE, AGGRESSION, MENTALITY, GOAL KEEPING, TEAM WORK).

3. Predict B1 scores for **all** players across all positions (not just the Calc-sheet subset).

4. Extract special abilities from the abilities sheet's 20 binary flag columns: REACTION, PLAYMAKING, PASSING, SCORING, 1-1 SCORING, POST PLAYER, LINES, MIDDLE SHOOTING, SIDE, CENTRE, PENALTIES, 1-TOUCH PASS, OUTSIDE, MARKING, SLIDING, COVERING, D-LINE CONTROL, PENALTY STOPPER, 1-ON-1 STOPPER, LONG THROW.

5. Filter the pool:
   - Exclude OVR ≥ 95 and all owned-team players
   - **Primary filter:** B1 score 80–100 AND Gap (B1 − OVR) ≥ +3.0
   - **Starred exception:** also include players with Gap ≥ +1.5 if they have ≥3 special abilities OR Condition/Fitness ≥ 6
   - Minimum 10 entries per position × class cell (relax the gap threshold if the pool is thin)

6. Build a multi-sheet XLSX tracker with:
   - **README** — methodology, R² per position, formula weights, notes
   - **Master Pool** — all entries, filterable, with an `Interest` dropdown (Target / Consider / Maybe / Skip / Picked) and a `Notes` column for me to fill in
   - **Position × Class Grid** — top 3 per position×class cell, compact overview
   - **Free Agents** — unattached players only (top draft priority, no team competition)
   - **A/B/C/D-class** — four per-class sheets grouped by position
   - **Top 50 Diamonds** — biggest gaps across the entire pool
   - **GK Picks** — rank keepers by B1 within class (NOT by gap); highlight COND ≥ 6 keepers separately
   - **Draft Slot Planner** — 7 fillable slots (3A + 2B + 1C + 1D)

## Important nuances (don't skip these)

- CWP uses identical formula to CB, and WB uses identical formula to SB — no out-of-position penalty in Phoenix-style patches. So the CB/CWP and SB/WB tables share values.
- GK formula is RESPONSE + GOAL KEEPING dominated (≈ 68% of weight). This makes GK gaps very tight — typical max gap is +2 to +2.5, well below the +3 field-player threshold. **Do not rank GKs by gap.** Use B1 ranking within class plus COND ≥ 6 as the primary signals.
- Condition/Fitness 6–7 is rare (~8–9% of the patch) and is a real performance multiplier over a long season with CONDITION RANDOM enabled — flag it prominently.
- The same player appearing at multiple positions should produce multiple rows — each row is a different draft consideration (e.g., a utility midfielder may over-perform at DMF, CMF, and WF simultaneously).
- Character encoding for accented names may be mangled in the source Excel (`�` characters). Apply best-effort heuristic fixes where context is clear (e.g., `K. Mbapp�` → `K. Mbappé`, `Atl�tico de Madrid` → `Atlético de Madrid`, `_x0007_` endings → `ć`). Leave ambiguous cases as-is.

## Deliverables

1. The complete XLSX tracker (present it via the file tool so I can download it)
2. A brief summary including:
   - Total players loaded, eligible pool size after filters, total over-performer entries
   - R² per position (to confirm the ridge models reproduced the patch formula accurately)
   - Top 5–10 diamonds (biggest gaps across the entire pool)
   - Any notable changes from prior patches if recognisable

## Optional shortcut

If I also attach `pes6_tracker_build.py` alongside the Excel, you can just run it directly:
```
python pes6_tracker_build.py <patch>.xlsx <output>.xlsx --owned-teams owned.txt
```
Otherwise, build the pipeline from scratch following the methodology above.

---
