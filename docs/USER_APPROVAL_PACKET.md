# Human Review & Approval Packet: Mahindra & Mahindra Single-Target Excel Model (Issue #8)

> [!IMPORTANT]
> **HISTORICAL VALUATION SNAPSHOT NOTICE**: All financial models, closing stock prices, market capitalizations, Treasury yields, peer trading multiples, and empirical regression datasets in this release reflect historical market observations as of **Wednesday, October 15, 2025 (`2025-10-15`)**. They are **NOT** current or live market valuations and do not incorporate subsequent market events.
>
> **APPROVAL GATE STATUS**: In strict adherence to repository safety invariants (`DISALLOWED_REVIEWER_TOKENS` prohibiting agent self-promotion), the **`human_pack_approval` gate remains PENDING (`ready: false`)**. This workbook is delivered in proposal/validation state pending explicit sign-off by the human reviewer. PR #8 remains open and unmerged.

---

## 1. Single-Target Scope & Governance (Issue #8)

In direct resolution of **GitHub Issue #8**, the build process and deliverables have been restructured to focus exclusively on **Mahindra & Mahindra Limited (`MM`)**:
- **Active Deliverable**: `MM-FullTemplate.xlsx` is the sole generated company workbook delivered.
- **Decommissioned Deliverables**: Active generated deliverables for LTM, CLEAN, and EDELWEISS have been removed from `deliverables/V16_user_handoff/` and `companies/<TARGET>/template_v16/`.
- **Archive Safety**: The full 4-company handoff archive is preserved permanently at `deliverables/archive/V16_user_handoff_four_company_archive.zip` (9,682,773 bytes). Underlying company source snapshots, registries, input deltas, and test assets are retained.
- **Master Template Preservation**: The 16-sheet master template [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-gemini-mm-v17/originals/Marksans%20Keshav.xlsx) is verified byte-identical (`4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`) via `python verify_sources.py`.

---

## 2. Deliverable Verification Matrix

The delivered workbook has been built and verified natively via Microsoft Excel COM 16.0:

| Property | Target Value / Status | Verification Method |
| :--- | :--- | :--- |
| **Target Company** | Mahindra & Mahindra Limited (`MM`) | Primary Build Specification |
| **Worksheet Count** | **16 worksheets** in exact master order | openpyxl & Excel COM 16.0 |
| **Formula Errors** | **0 errors** across all 16 sheets | openpyxl `data_only=True` error scan |
| **Circular References** | **0 circular references** (`ws.CircularReference is None`) | Microsoft Excel COM 16.0 |
| **Historical BS Checks (`BS!C40:M40`)** | **0.00 Cr** across all 11 periods (FY15–FY25) | Dynamic formula `={col}38-{col}16` |
| **FY26E Balance Sheet Check (`BS!N40`)** | **`BLOCKED`** | Explicit blocked indicator without plugs |
| **Cost of Capital: Raw OLS Beta** | **1.4746** (`Beta!K7`) | Aligned 104-week Yahoo Finance regression |
| **Cost of Capital: Blume Adjusted Beta** | **1.3164** (`Beta!K13`) | `(2/3)*Raw + (1/3)*1.0` |
| **Cost of Capital: WACC** | **`Unavailable`** (`WACC-DCF!J39`) | Explicit allowlist restriction |
| **Primary Valuation (`DCF 1st way!E42`)** | **`Unavailable`** | Single source-of-truth cell |
| **Cover Link (`Model ->!C9`)** | **`Unavailable`** | Dynamic formula `='DCF 1st way'!E42` |
| **Native Sparklines** | Preserved XML `<x14:sparklineGroups>` (`C6:G6` to `C26:G26`) | `xl/worksheets/sheet3.xml` audit |
| **Deliverable Checksum (SHA-256)** | `bdfd6a9903f1c6be34d13641dbdb7ecf0a04c83e04653996239d4d720704c54d` | SHA-256 verification |
| **Deliverable File Size** | **92,153 bytes** | Filesystem audit |
| **Preview Schedules** | **15 landscape PDF files** under `previews/MM/` | Native Office COM export |

---

## 3. Detailed Audit of the 8 Repaired Defects

### Defect 1: Circular Calculation Loops in `compdata`
- **Prior Defect**: 5 circular formula loops existed between `F19:F25` and `K4:K10` (`F19 ↔ K4`, `F20 ↔ K5`, `F21 ↔ K6`, `F22 ↔ K7`, `F25 ↔ K10`) due to reverse-solving implied EBITDA from peer multiples.
- **Repair**: Completely eliminated formula reverse-solving in `compdata`. Peer comparison blocks in `Comps` are cleanly marked unavailable. Excel COM confirms `ws.CircularReference is None` across all 16 worksheets.

### Defect 2: Cash Flow Statement Integrity
- **Prior Defect**: Marksans pharmaceutical carryover items remained in historical rows, missing detail lines were treated as zero, and unsupported FY26E forecasts were calculated.
- **Repair**: Retained authentic Screener historical subtotals (Cash from Operating Activities, Cash from Investing Activities, Cash from Financing Activities). Cleared Marksans carryovers. Marked FY26E forecast column N explicitly blocked.

### Defect 3: Balance Sheet Historical Ties & Blocked Forecast
- **Prior Defect**: Balance sheet check was either hardcoded or exposed unreconciled gaps due to unsupported asset allocations.
- **Repair**: Preserved supported Screener historical totals (`Total Assets` and `Total Liabilities`), verifying that `Total Assets - Total Liabilities = 0.00 Cr` across all 11 historical periods (`BS!C40:M40`). Unavailable asset line items are left blank rather than populated with plugs. FY26E check `BS!N40` is set to `BLOCKED`.

### Defect 4: Main Financials Common-Size Errors
- **Prior Defect**: Clearing Data Sheet row 75 destroyed the Total Assets denominator, causing widespread `#DIV/0!` errors across historical common-size balance sheet rows (141–148).
- **Repair**: Preserved the Total Assets denominator row dynamically in Data Sheet (`=col61`). Cleared unsupported current asset breakdowns (rows 146–148) and blocked FY26E column J, resulting in zero formula errors across the entire sheet.

### Defect 5: Comps Sheet Layout & Peer Comparisons
- **Prior Defect**: Comps sheet contained placeholder zeros, broken `#N/A` blocks, and unallowlisted peer estimates.
- **Repair**: Clean layout displaying only supported M&M market statistics (closing price ₹3,497.20, market cap ₹4,34,917 Cr as of 15-Oct-2025). Peer valuation multiples are marked unavailable.

### Defect 6: Conglomerate DCF & Valuation Scope
- **Prior Defect**: An industrial FCFF DCF was forced onto a conglomerate whose subsidiary (MMFS) holds over ₹1.3 lakh Cr in NBFC lending assets, producing an invalid valuation.
- **Repair**: Industrial FCFF DCF is methodologically disallowed for hybrid conglomerates. Single source-of-truth cell `DCF 1st way!E42` is set to `Unavailable`. The cover sheet dynamically displays `Unavailable`. Stale numeric DCF projections and terminal calculations have been cleared.

### Defect 7: Independent Beta & WACC Recomputation
- **Prior Defect**: Beta relied on stale external assertions without verifiable return series linkage.
- **Repair**: Independently recomputed OLS regression beta from the committed Yahoo Finance 104-week return dataset (`work/sources/weekly_market_returns_104w_v2.json`) pairing M&M.NS against NIFTY 50 (`^NSEI`):
  - Covariance = `0.0016629737`, Variance = `0.0011277440`
  - Raw Beta = **1.4746023202**
  - Blume Adjusted Beta = **1.3164015468**
  - WACC is set to `Unavailable` in `WACC-DCF!J39` due to missing debt structure facts on allowlisted sources.

### Defect 8: Restoration of Native Excel Sparklines
- **Prior Defect**: Sparklines in `Ratio Analysis` were stripped or corrupted by intermediate openpyxl saves.
- **Repair**: Native `<x14:sparklineGroups>` XML is restored directly into `xl/worksheets/sheet3.xml`, defining line sparklines for rows 6 through 26 with valid data ranges `C6:G6` to `C26:G26`.

---

## 4. Strict 7-Website Allowlist Compliance

All sourced facts are verified exclusively from the websites named in the original Marksans workbook plus Screener.in:
1. **Screener.in**: Historical financial statements and balance sheets (FY15–FY25).
2. **Yahoo Finance**: Aligned 105 weekly adjusted-close prices and 104 weekly return pairs.
3. **NSE India**: Official closing quotation of ₹3,497.20 as of October 15, 2025.
4. **Trading Economics**: Benchmark sovereign bond yield proxy (6.51%).
5. **Top Stock Research**: Sector classification cross-check.
6. **Investing.com**: Secondary yield verification.
7. **Prime Investor**: Ratio consistency cross-check.

*Disallowed Sources*: Zero use of BSE, investor presentations, RBI, or unverified web searches.

---

## 5. Human Sign-Off Checklist

- [ ] Confirm `MM-FullTemplate.xlsx` opens in Microsoft Excel with 16 worksheets and zero formula errors.
- [ ] Confirm Excel displays zero circular reference warnings upon calculation (`Ctrl + Alt + F9`).
- [ ] Verify `BS!C40:M40` evaluates to `0.00 Cr` and `BS!N40` displays `BLOCKED`.
- [ ] Verify `Model ->!C9` and `DCF 1st way!E42` display `Unavailable`.
- [ ] Verify `Beta!K7` (1.4746) and `Beta!K13` (1.3164) match the independent regression results.
- [ ] Verify native sparklines render visually in `Ratio Analysis` column H.
- [ ] Approve PR #8 for merge into main branch.
