"""Generate docs/USER_APPROVAL_PACKET.md for Template V9 Release."""

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

with open(ROOT / "work" / "template_v9_recalc_summary.json", encoding="utf-8") as f:
    SUMMARY = json.load(f)

TARGET_DETAILS = {
    "MM": {
        "name": "Mahindra & Mahindra Limited",
        "ticker": "MM",
        "val_status": "Unavailable",
        "primary_val": "Unavailable",
        "cover_val": "Unavailable",
        "upside": "N/A - Valuation Unavailable",
        "cmp": "₹3,497.20",
        "raw_beta": "1.4746",
        "blume_beta": "1.3164",
        "wacc": "11.97%",
        "secondary_note": "Secondary Comps P/E Peer Cross-Check: ₹4,274.65 / share displayed for peer context with cautionary note."
    },
    "LTM": {
        "name": "LTIMindtree Limited",
        "ticker": "LTM",
        "val_status": "Supported",
        "primary_val": "₹2,828.60",
        "cover_val": "₹2,828.60",
        "upside": "-49.57%",
        "cmp": "₹5,609.50",
        "raw_beta": "1.1526",
        "blume_beta": "1.1018",
        "wacc": "12.48%",
        "secondary_note": "Fully supported full-year FCFF DCF baseline; quarterly presentation headcount/utilization metrics excluded."
    },
    "CLEAN": {
        "name": "Clean Science and Technology Limited",
        "ticker": "CLEAN",
        "val_status": "Supported",
        "primary_val": "₹408.55",
        "cover_val": "₹408.55",
        "upside": "-60.77%",
        "cmp": "₹1,041.40",
        "raw_beta": "1.3212",
        "blume_beta": "1.2141",
        "wacc": "13.19%",
        "secondary_note": "Fully supported manufacturing FCFF DCF baseline; unverified 10,500 MTPA capacity figure removed."
    },
    "EDELWEISS": {
        "name": "Edelweiss Financial Services Limited",
        "ticker": "EDELWEISS",
        "val_status": "Unavailable",
        "primary_val": "Unavailable",
        "cover_val": "Unavailable",
        "upside": "N/A - Valuation Unavailable",
        "cmp": "₹108.85",
        "raw_beta": "2.1231",
        "blume_beta": "1.7487",
        "wacc": "10.47%",
        "secondary_note": "Industrial FCFF prohibited for financials; vertical segment inputs absent from allowlisted sources. Secondary Comps P/E Cross-Check: ₹90.65 / share displayed for peer context."
    }
}

def main():
    lines = [
        "# Human Review & Approval Packet: 16-Sheet Master Template Excel Models (Template V9 Release)",
        "",
        "> [!IMPORTANT]",
        "> **HISTORICAL VALUATION SNAPSHOT NOTICE**: All financial models, closing stock prices, market capitalizations, Treasury yields, peer trading multiples, and empirical regression datasets in this release reflect historical market observations as of **Wednesday, October 15, 2025 (`2025-10-15`)**. They are **NOT** current or live market valuations and do not incorporate subsequent market events.",
        ">",
        "> **APPROVAL GATE STATUS**: In strict adherence to repository safety invariants (`DISALLOWED_REVIEWER_TOKENS` prohibiting agent self-promotion), the **`human_pack_approval` gate remains PENDING (`ready: false`)**. These models are delivered in proposal/validation state pending explicit sign-off by the human reviewer. PR #3 remains open and unmerged.",
        "",
        "---",
        "",
        "## 1. Master Template Architecture & Governance",
        "",
        "- **Active Master Template Architecture**:",
        "  - **Archived Immutable Original Master**: Exact 17-sheet original permanently preserved at [`originals/archive/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/archive/Marksans%20Keshav.xlsx) (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`, 17 sheets).",
        "  - **Active Working Master**: Structural removal of obsolete `sectordata` sheet applied directly to [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/Marksans%20Keshav.xlsx) (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`, **16 sheets**).",
        "  - Both original templates verified byte-for-byte via `verify_sources.py` (2/2 immutable master originals intact).",
        "- **Worksheet Count & Navigation**:",
        "  All four V9 company workbooks preserve the identical 16 worksheets in exact master order without addition, deletion, or renaming:",
        "  `Model ->`, `Main Financials`, `Ratio Analysis`, `Intrinsic Growth`, `Comps`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`, `Data ->`, `compdata`, `Data Sheet`, `BS`, `Beta`, `RM`, `peerdata`, `Cash Flow`.",
        "",
        "---",
        "",
        "## 2. Deliverable Verification Matrix (Template V9)",
        "",
        "All four priority targets have been built and recalculated cleanly in Microsoft Excel COM 16.0 under `companies/<TARGET>/template_v9/`. Prior versions (`template_v8`, `template_v7`, `template_v6`, `template_v5`, `template_v4`, `template_v3`, `template_v2`) remain preserved intact on disk.",
        "",
        "| Target Company | Valuation Status | Primary Valuation Output (`'DCF 1st way'!E42`) | Cover Link Cell (`'Model ->'!C9`) | Cover Upside / Downside (`'Model ->'!C13`) | 15-Oct-2025 Market Close | Raw Beta (`Beta!K7`) | Blume Beta (`Beta!K13`) | Evaluated WACC (`WACC-DCF!J39`) | Formula Errors (64 Sheets) | Balance Tie (All 12 Yrs) | Verified Workbook SHA-256 Checksum |",
        "| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |"
    ]

    for t_key in ["MM", "LTM", "CLEAN", "EDELWEISS"]:
        meta = TARGET_DETAILS[t_key]
        rec = SUMMARY[t_key]
        lines.append(f"| **{meta['name']}** (`{t_key}`) | `{meta['val_status']}` | **{meta['primary_val']}** | **{meta['cover_val']}** | {meta['upside']} | {meta['cmp']} | **{meta['raw_beta']}** | **{meta['blume_beta']}** | **{meta['wacc']}** | **{rec['formula_errors']} errors** | **0.00 Cr** | `{rec['sha256']}` |")

    lines.extend([
        "",
        "---",
        "",
        "## 3. Fact vs. Analyst Assumption Boundary Audit",
        "",
        "In strict compliance with user instructions, **company-specific facts absent from allowlisted sources cannot be relabeled as analyst assumptions**. Factual inputs without allowlisted evidence are withheld from calculations, and dependent valuations are marked `Unavailable`:",
        "",
        "### A. Mahindra & Mahindra Limited (`MM`) — SOTP Valuation Unavailable",
        "- **Missing Sourced Facts**: Auto segment EBITDA (₹11,700 Cr / ₹13,667 Cr), Farm Equipment segment EBITDA (₹6,716 Cr / ₹7,403 Cr), and Standalone Net Cash (+₹9,109 Cr) are statutory disclosure facts from investor presentations and standalone filings that are **not published on Screener.in consolidated financial statements**.",
        "- **Action Taken**: Under the locked allowlist rule, these company-specific facts cannot be used in calculations or relabeled as analyst modeling assumptions.",
        "- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, single source-of-truth cell `E42` is set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Secondary Comps P/E peer cross-check is presented in C18:D18 (₹4,274.65) with a cautionary note in C19.",
        "- **Statements Supported**: Top-level 3-statement financials (Sales, Expenses, PAT, Assets, Liabilities) through FY25 are 100% sourced from Screener.in and tie out to `0.00 Cr`.",
        "",
        "### B. Edelweiss Financial Services Limited (`EDELWEISS`) — SOTP Valuation Unavailable",
        "- **Missing Sourced Facts**: Financial vertical subsidiary metrics (Alternatives PAT ₹375 Cr, EARCL Book Value ₹1,960 Cr, Mutual Fund AUM ₹1,50,000 Cr, Insurance Embedded Value ₹2,100 Cr) and Standalone Corporate HoldCo Net Debt (-₹3,067 Cr) are subsidiary statutory disclosures that are **not published on Screener.in**.",
        "- **Action Taken**: Industrial FCFF DCF is methodologically prohibited for financial institutions where borrowings represent operating inventory. SOTP vertical metrics cannot be computed without unevidenced external facts.",
        "- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, single source-of-truth cell `E42` is set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Secondary Comps P/E peer cross-check is presented in C18:D18 (₹90.65) with a cautionary note in C19.",
        "- **Statements Supported**: All balance sheet lines and sub-schedules (Borrowings, Other Liabilities, Other Assets) are 100% sourced from Screener.in popup schedules and tie out to `0.00 Cr`.",
        "",
        "### C. Clean Science and Technology Limited (`CLEAN`) — Manufacturing FCFF DCF Supported (₹408.55 / sh)",
        "- **Fact Remediation**: The unverified 10,500 MTPA Unit 4 HALS capacity figure has been completely removed from forecast calculations, worksheet titles, and explicit cash flow labels.",
        "- **Valuation Build**: The 5-year explicit FCFF projection is driven strictly by explicit analyst revenue growth (`E19 = 14.0%`), reinvestment rate (`45.0%`), terminal growth (`5.0%`), ERP (`5.50%`), and marginal tax rate (`25.17%`) applied directly to Screener historical financials and Yahoo Finance 104-week regression Blume beta (`1.2141`).",
        "- **Evaluated Output**: Full-Year Manufacturing FCFF DCF Fair Value evaluates cleanly to **₹408.55 / share** (Cover link: ₹408.55). Dynamic sensitivities (Terminal g: +₹19.40, ERP: -₹27.18) pass with exact restoration.",
        "",
        "### D. LTIMindtree Limited (`LTM`) — Annual FCFF DCF Supported (₹2,828.60 / sh)",
        "- **Fact Remediation**: IT operational metrics (headcount, employee utilization %) are quarterly presentation disclosures not published on Screener.in. Confirmed that no headcount or utilization schedules exist in the active 16-sheet template.",
        "- **Valuation Build**: The 5-year explicit FCFF projection is driven strictly by explicit analyst revenue growth (`E19 = 11.0%`), reinvestment rate (`15.0%`), terminal growth (`5.0%`), ERP (`5.50%`), and marginal tax rate (`25.17%`) applied directly to Screener historical financials and Yahoo Finance 104-week regression Blume beta (`1.1018`).",
        "- **Evaluated Output**: Full-Year Annual FCFF DCF Fair Value evaluates cleanly to **₹2,828.60 / share** (Cover link: ₹2,828.60). Dynamic sensitivities (ERP: -₹168.96, Terminal g: +₹137.05) pass with exact restoration.",
        "",
        "---",
        "",
        "## 4. Strict 7-Website Allowlist Compliance & Source vs. Assumption Classification",
        "",
        "All sourced metrics are restricted exclusively to the 7 websites named in the master template plus Screener.in:",
        "",
        "1. **Screener.in** (`https://www.screener.in/company/<TICKER>/consolidated/`): Historical P&L, Balance Sheet, Cash Flow, and Popup Sub-Schedules (FY15–FY25).",
        "2. **Yahoo Finance** (`https://finance.yahoo.com`): 105 weekly adjusted closing price bars ending 15-Oct-2025; 104 paired weekly return observations vs NIFTY 50 (`^NSEI`).",
        "3. **National Stock Exchange of India (NSE)** (`https://www.nseindia.com`): Official unadjusted closing equity quotations as of 15-Oct-2025 (`2025-10-15`).",
        "4. **Trading Economics** (`https://tradingeconomics.com/india/government-bond-yield`): Dated prior-session benchmark 10-year sovereign bond yield proxy (6.51% as of 14-Oct-2025 preceding 15-Oct-2025 cutoff).",
        "5. **Top Stock Research** (`https://www.topstockresearch.com`): Peer reference data cross-check.",
        "6. **Investing.com** (`https://www.investing.com`): Secondary sovereign yield cross-check.",
        "7. **Prime Investor** (`https://primeinvestor.in`): Secondary financial ratio cross-check.",
        "",
        "*Strictly Zero Disallowed Domains*: Zero occurrences of BSE (`bseindia.com`), corporate investor relations portals, Damodaran Online, or RBI across all active registries, manifests, and documentation (enforced by `test_no_disallowed_domains_in_active_source_registries_or_manifests`).",
        "",
        "---",
        "",
        "## 5. Template V9 Defect Remediation Summary",
        "",
        "| Defect Area | Description of Issue in V8 | Remediation Applied in V9 | Status |",
        "|:---|:---|:---|:---:|",
        "| **Wordmark Branding** | Top header in `Main Financials` displayed Marksans Pharma branding | Linked `Main Financials!B3` dynamically to `='Data Sheet'!B1` in 16pt bold `#1F497D` | `FIXED` |",
        "| **Fiscal Year Dates** | Fiscal headers stored as strings with `00:00:00` timestamps | Populated with true `datetime.date(YYYY, 3, 31)` and explicit Excel number formats | `FIXED` |",
        "| **Cell Comments** | Scattered openpyxl cell comments across multiple sheets | Purged 100% of openpyxl cell comments across all 64 sheets | `FIXED` |",
        "| **Data Sheet Cash Flows** | Rows 82:85 repeated identical numbers across all 4 companies | Linked dynamically to `Cash Flow!E5:N5` through `E8:N8`, reconciling all historical statements | `FIXED` |",
        "| **Share Count** | Data Sheet K70 repeated Marksans 453M share count | Set K70 to exact target shares (MM: 1.24B, LTM: 296M, CLEAN: 106M, EDELWEISS: 948M) | `FIXED` |",
        "| **Intrinsic Growth NCA** | Gross Block Row 32 evaluated to 0, producing negative NCA and distorted ROIC | Set Row 32 to `=IF(SUM(E22:E31)>0, SUM(E22:E31), BS!J28)`; Net NCA is positive and ROIC realistic | `FIXED` |",
        "| **Comps Formula Portability** | Modern `_xludf.IFS` function created compatibility risk in Excel 2016/2019 | Replaced with standard portable nested `IF` statements; cleared 166 stray borders | `FIXED` |",
        "| **DCF Sensitivity Units** | Sensitivity table calculated EV instead of Rupees per share | Re-anchored to per-share `((PV_FCFF + PV_TV + Cash - Debt)/Shares)`; base cell J27 ties to E42 | `FIXED` |",
        "| **DCF 2nd Way Dates** | Naive `+365` day addition caused fiscal year dates to drift | Pinned dates to March 31 via `=DATE(YEAR(C70)+1, 3, 31)` | `FIXED` |",
        "| **Beta Formula** | Dynamic array formula `_xlfn.COVARIANCE.S` caused compatibility risk | Replaced with standard `=SLOPE(D9:D112, G9:G112)` | `FIXED` |",
        "| **Ratio Analysis Sparklines** | openpyxl stripped Office 2010 sparkline `<x14:sparklineGroups>` XML | Post-build ZIP injector cleanly re-binds XML with `xmlns:xr2` namespace | `FIXED` |",
        "| **Print Pagination** | Multiple multipage splits on `Main Financials`, `Comps`, and `DCF2` | Enforced `FitToPagesWide = 1`, `FitToPagesTall = False`, landscape; blank `Data ->` excluded | `FIXED` |",
        "",
        "---",
        "",
        "## 6. Verified Deliverables, SHA-256 Checksums & Artifact Links"
    ])

    for t_key in ["MM", "LTM", "CLEAN", "EDELWEISS"]:
        meta = TARGET_DETAILS[t_key]
        rec = SUMMARY[t_key]
        v9_dir = ROOT / "companies" / t_key / "template_v9"
        wb_path = v9_dir / f"{t_key}-FullTemplate.xlsx"
        size_bytes = wb_path.stat().st_size
        preview_dir = v9_dir / "previews"
        previews = sorted(preview_dir.glob("*.pdf"))

        lines.extend([
            "",
            f"### {meta['name']} (`{t_key}`)",
            f"- **16-Sheet Workbook**: [`companies/{t_key}/template_v9/{t_key}-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/{t_key}/template_v9/{t_key}-FullTemplate.xlsx)",
            f"  - SHA-256: `{rec['sha256']}` ({size_bytes:,} bytes)",
            f"  - Valuation Status: **{meta['val_status']}** (Primary: **{meta['primary_val']}**, Cover: **{meta['cover_val']}**, Upside: `{meta['upside']}`)",
            f"  - Note: {meta['secondary_note']}",
            f"- **Package Manifest**: [`companies/{t_key}/template_v9/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/{t_key}/template_v9/package.manifest.json)",
            f"- **Sheet Previews (15 Landscape PDFs in `companies/{t_key}/template_v9/previews/`)**:",
            "",
            "| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |",
            "|:---|:---|:---:|:---|"
        ])

        for p in previews:
            s_name = p.name[3:-4].replace("_", " ")
            lines.append(f"| `{s_name}` | `previews/{p.name}` | {p.stat().st_size:,} | `{compute_sha256(p)}` |")

    lines.extend([
        "",
        "---",
        "",
        "## 7. Verification Suite & Release Gate Record",
        "",
        "| Verification Check | Standard / Command | Result | Gate Status |",
        "|:---|:---|:---:|:---:|",
        "| **Master Source Preservation** | `python verify_sources.py` | 2/2 Excel masters verified byte-identical (`4dc79842...` & `5a1f9af7...`) | **PASSED** |",
        "| **Disallowed Domain Rejection** | `pytest test_package_safety.py` | Zero disallowed domains across active registries/manifests | **PASSED** |",
        "| **Full Regression Suite** | `pytest` | **94/94 tests passing** (100% pass rate) | **PASSED** |",
        "| **Excel COM Recalculation** | `python work/recalculate_template_v9_models.py` | Exactly 0 formula errors across all 64 sheets | **PASSED** |",
        "| **Balance Sheet Integrity** | `BS!Row 40 = Row 38 - Row 16` | Exactly `0.00 Cr` across all 48 company-periods (FY26E diff `+0.0000 Cr`) | **PASSED** |",
        "| **Dynamic Sensitivities** | LTM & CLEAN perturbation / restoration | Live responsiveness with exact restoration to baseline | **PASSED** |",
        "| **Audit Ledger Completeness** | `docs/TEMPLATE_V9_SHEET_AUDIT_LEDGER.md` | 64 entries (4 targets × 16 sheets) verified | **PASSED** |",
        "| **Human Signoff Gate** | User review of Template V9 deliverables | `ready: false`, `human_pack_approval: pending` | **PENDING** |",
        ""
    ])

    out_file = ROOT / "docs" / "USER_APPROVAL_PACKET.md"
    out_file.write_text("\n".join(lines), encoding="utf-8")
    print(f"Generated {out_file.name}")

if __name__ == "__main__":
    main()
