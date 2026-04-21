# Reusable prompt — PES6 tier-list comparison

Copy the block between the `---` lines into a fresh Claude chat, **edit the bracketed sections**, attach the new patch Excel (and optionally the comparison script), and send.

---

I'm playing PES6 with the **[PATCH NAME]** patch (Phoenix/Shollym-style, slow/balanced). Attached is the full patch Excel. I've built my own tier list from eye-test and memory (pasted below in [Serbian / English / language X]). I want you to:

1. Build a **data-driven tier list** from the same Excel using **top-16 average OVR** as the primary metric (= mean OVR of each club's 16 highest-OVR players — captures starting XI + 5 rotation depth). Use top-16 best-B1 (highest B1 across all 11 position formulas per player) as a secondary/validation metric.

2. Choose sensible tier thresholds that produce a distribution roughly similar in shape to mine (so we're comparing like-for-like), while using **natural gaps in the OVR distribution** as tier boundaries where possible. Use these tiers: **BANNED, A, B, C, D, E, F** (banned = significantly OP, F = weakest).

3. For the **BANNED tier**: use a clear top-16 OVR threshold (typically ≥ 89.0) that catches the "super-clubs." Note in your analysis whether my banned tier includes any teams you wouldn't have banned purely on depth — these are probably "star-power bans" (Messi/Ronaldo/etc.) rather than depth bans, and that's a legitimate difference to flag, not necessarily an error.

4. **Compare the two tier lists team-by-team** and compute:
   - Exact tier match rate (%)
   - Within-1-tier rate (%)
   - Within-2-tier rate (%)

5. **Build a multi-sheet XLSX comparison workbook** with these sheets:
   - **README** — methodology, tier thresholds used, accuracy headline numbers
   - **Full Comparison** — all ~140 clubs side by side, with columns: OVR rank, team, top-16 OVR, top-16 B1, Your tier, My tier, Tier diff, Match?, Owned?, Notes
   - **Mismatches** — only rows where our tiers differ, sorted by magnitude
   - **Major Disagreements** — only ≥2 tier differences (the genuinely surprising cases), with brief commentary on each
   - **My Tier List** — my data-driven list grouped by tier, with your tier shown alongside for cross-reference
   - **Your Tier List** — your original list grouped by tier, for reference

6. Use tier color-coding throughout (BANNED=black, A=red, B=orange, C=gold, D=green, E=blue, F=grey) and mark currently-owned teams from my draft with an "Owned?" column.

7. Present the file via the file tool and give me a written analysis covering:
   - Accuracy headline numbers
   - Tier distribution comparison (counts per tier, your vs. mine)
   - The biggest disagreements one by one, with brief explanations of why each one looks different under the data
   - Any patterns in my errors (e.g., "you systematically overrate big-brand teams" or "you underrate teams with hidden-value attackers")
   - Any teams missing from my list entirely

## My tier list (edit this block)

**BANNED (OP):**
[paste team names, comma or newline separated]

**Tier A:**
[paste]

**Tier B:**
[paste]

**Tier C:**
[paste]

**Tier D:**
[paste]

**Tier E:**
[paste]

**Tier F:**
[paste]

## Currently-owned teams (to mark in the output)

Liga 1: [list — or say "same 13 as before" if unchanged]
Liga 2: [list — or say "same 14 as before"]

## Notes on name-matching

Patch Excel names sometimes have encoding mangled (`Atl�tico`, `M�nchen`, `Bayern M?nchen`). Please normalize these transparently to display readable names while keeping the patch-name lookup working internally. A few ones to expect:
- `Bayern M\ufffdnchen` → Bayern München
- `Atl\ufffdtico de Madrid` → Atlético de Madrid
- `Fortuna D\ufffdsseldorf` → Fortuna Düsseldorf
- `K\ufffdln` → Köln
- `Alav\ufffds` → Alavés
- `M\ufffdlaga` → Málaga
- `Deportivo de La Coru\ufffda` → Deportivo de La Coruña

Short-name mappings I use for common teams (add any more you encounter):
Leverkusen=Bayer Leverkusen, RB Leipzig=Leipzig, Sociedad=Real Sociedad, PSV=PSV Eindhoven, Frankfurt=Eintracht Frankfurt, Marseille=Olympique Marseille, Lyon=Olympique Lyonnais, Sporting=Sporting CP, Club Brugge=Brugge, Wolverhampton=Wolves, Hamburg=Hamburger, Monchengladbach=Borussia M'gladbach, Kiel=Holstein Kiel, Hertha Berlin=Hertha, Hannover 96=Hannover, Athletic Bilbao=Athletic, Kairat=Kairat Almaty, Bodo-Glimt=Bodo Glimt, Düsseldorf=Fortuna Düsseldorf, Inter=Internazionale, Tottenham=Tottenham, Dortmund=Borussia Dortmund, United=Manchester United, City=Manchester City, PSG=Paris Saint-Germain, Bayern=Bayern München, Atletico Madrid=Atlético de Madrid, Al Nassr=Al-Nassr

If any team in my list doesn't map to a patch club, flag it. If any patch club is missing from my list, flag it and assign a tier based on the data.

---

**Optional shortcut:** If I also attach `pes6_tracker_build.py` from earlier sessions, you can reuse its ridge-regression pipeline rather than rebuilding B1 predictions from scratch.

---

**Tips for your tier list**

- Use a consistent criterion (e.g., "how well would this squad do in our league" or "pure squad-depth rating"). If you use one criterion for most teams and another for a few (e.g., star-power for Real Madrid/Inter Miami/Al-Nassr), call that out explicitly — it makes the disagreement meaningful rather than looking like an error.
- Paste your tier list into a fresh chat exactly as you wrote it; Claude will translate from Serbian or any other language automatically. No need to pre-translate.
- If your draft has different tier names (e.g., S/A/B/C/D/E or 1–7), just substitute — the structure is the same.
