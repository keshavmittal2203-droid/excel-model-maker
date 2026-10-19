# User-Supplied Financial Modelling Course Reference

- Title visible on the supplied playlist: **Learn Financial Modelling - Step by Step** (The Valuation School).
- Canonical URL: https://www.youtube.com/playlist?list=PL3uUjzLk6PulhRop_ffNeHyK0kprzO4cT
- Added at the user's request on 2026-10-16.
- User reports that this course explains the methodology used for the uploaded Excel workbook (`originals/Marksans Keshav.xlsx`).
- Inventory extraction completed on 2026-10-16 from live playlist metadata into `docs/course_lessons.json` (34 total lessons).

## Course Lesson Inventory & Workbook Sheet Mapping

The core 3-statement financial modelling and equity research modules (Sessions 4–22) map directly to the sheets of `originals/Marksans Keshav.xlsx`:

| Session # | Video Title | Video ID | Target Sheet(s) in `Marksans Keshav.xlsx` | Core Financial Subject |
|---|---|---|---|---|
| Session 4 | 3 Statement Financial Modelling | `-VIqSzEFvjM` | `Data Sheet`, `BS`, `Cash Flow`, `Main Financials` | Historical financial statement input mapping, P&L, BS, Cash Flow linking |
| Session 5 | Ratio Analysis | `t93qZunzHd8` | `Ratio Analysis` | Liquidity, solvency, profitability, and turnover ratios |
| Session 6 | Forecasting | `Ccv8eAqBSSQ` | `Data Sheet` (FY26–FY30), `Intrinsic Growth` | Top-line revenue projections, expense ratio schedules, asset schedules |
| Session 7 | DCF Excel Modelling | `pdveRcJucX4` | `DCF 1st way`, `DCF 2nd way` | Unlevered Free Cash Flow (FCFF), projection periods, discounting |
| Session 8 | Master WACC Modelling (Part 1) | `Pp_qhxHUziQ` | `WACC-DCF` | Weighted Average Cost of Capital, Cost of Equity (CAPM), Risk-free rate |
| Session 9 | What is Beta? Master WACC Modelling (Part 2) | `qRri_spNhHo` | `Beta`, `peerdata`, `RM` | Historical regression beta, peer un-levered and re-levered beta, Market Risk Premium |
| Session 10 | Master WACC Modelling (Part 3) | `N0VG4nbeoBI` | `WACC-DCF` | Pre-tax and post-tax cost of debt, marginal tax rate, capital structure weighting |
| Session 11 | Common Size Statement | `csv-QJTubG8` | `BS`, `Data Sheet` | Vertical percentage analysis of Balance Sheet and Income Statement |
| Session 12 | Growth Rate | `0N3zIEhaq1c` | `Intrinsic Growth` | Retention ratio, Return on Equity (ROE), sustainable intrinsic growth rate (SGR) |
| Session 13 | DCF (Part 2) | `h-rO3pNHHuw` | `DCF 1st way`, `DCF 2nd way` | Terminal value calculation (perpetual growth vs. exit multiple), Enterprise to Equity value bridge |
| Session 14 | Relative Valuation (Part 1) | `1aAh_CbnOj0` | `Comps`, `compdata`, `peerdata` | Comparable company universe, EV/EBITDA, P/E, EV/Sales multiples |
| Session 15 | Relative Valuation (Part 2) | `DpKdORBG224` | `Comps`, `sectordata` | Sector benchmarking, quartile rankings, target multiple application |
| Session 16 | Sensitivity Analysis | `cdPqgPmZQjM` | `DCF 1st way`, `DCF 2nd way` | Excel Data Tables (`{=TABLE()}`) varying WACC vs. terminal growth rate |
| Session 17 | Football Field Analysis | `O-SzKtazG-g` | Valuation summary | Range-based valuation compilation across DCF, Comps, 52-week trading |
| Sessions 18–19 | Dupont Analysis Modelling | `LIHqS2aQYtY`, `DcLvXfXPLKo` | `Ratio Analysis` | 3-stage and 5-stage DuPont decomposition (Operating margin × Asset turnover × Leverage) |
| Session 20 | Altman's Z-Score Modelling | `13YAACEFamQ` | `Ratio Analysis` | Distress prediction modeling for manufacturing / non-manufacturing firms |
| Sessions 21–34 | Profile, LBO & Special Topics | Various | Reference only | One-page company profile, LBO modeling, and IPO analysis |

## Evidence Provenance Tiers & Transcription Note

To maintain documentation integrity across autonomous contributions, course references are categorized into three explicit evidence tiers:

1. **Tier 1: Harvested Playlist Metadata (Verified)**
   - Lesson video IDs, canonical playlist URLs, and raw video titles were extracted directly from the public playlist HTML metadata into `docs/course_lessons.json` (34 total lessons).
2. **Tier 2: Inferred Sheet Mappings (Structural Hypothesis)**
   - The sheet-to-session mappings in the table above represent inferred structural alignment based on matching module titles to sheet names in `originals/Marksans Keshav.xlsx`. These serve as guidance and contextual references, not authoritative accounting proofs.
3. **Tier 3: Transcript-Verified Rules (None Currently Verified)**
   - **Access Limitation**: YouTube watch pages do not deliver raw static caption/timedtext tracks via HTTP without JavaScript hydration or authorized API tokens. Consequently, no granular modelling formulas or verbal assertions have been autonomously transcribed or verified from audio.
   - Any agent proposing a rule in `template-workflow/course-rule.example.json` format must cite local timestamped verbatim transcripts rather than asserting unverified spoken rules.

## Course Onboarding for AI Contributors

1. Use the session mapping above when establishing new model pack input mappings or financial identity checks.
2. For each extracted rule, use `template-workflow/course-rule.example.json` as the capture specification.
3. Record discrepancies and unresolved modeling decisions explicitly; course instructions provide methodology context but do not override mathematical accounting balance or independent financial validation.
