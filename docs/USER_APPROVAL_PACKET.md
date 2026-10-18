# Human Review & Approval Packet: 16-Sheet Master Template Excel Models (Template V8 Release)

> [!IMPORTANT]
> **HISTORICAL VALUATION SNAPSHOT NOTICE**: All financial models, closing stock prices, market capitalizations, Treasury yields, peer trading multiples, and empirical regression datasets in this release reflect historical market observations as of **Wednesday, October 15, 2025 (`2025-10-15`)**. They are **NOT** current or live market valuations and do not incorporate subsequent market events.
>
> **APPROVAL GATE STATUS**: In strict adherence to repository safety invariants (`DISALLOWED_REVIEWER_TOKENS` prohibiting agent self-promotion), the **`human_pack_approval` gate remains PENDING (`ready: false`)**. These models are delivered in proposal/validation state pending explicit sign-off by the human reviewer. PR #3 remains open and unmerged.

---

## 1. Master Template Architecture & Governance

- **Active Master Template Architecture**:
  - **Archived Immutable Original Master**: Exact 17-sheet original permanently preserved at [`originals/archive/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/archive/Marksans%20Keshav.xlsx) (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`, 17 sheets).
  - **Active Working Master**: Structural removal of obsolete `sectordata` sheet applied directly to [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/Marksans%20Keshav.xlsx) (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`, **16 sheets**).
  - Both original templates verified byte-for-byte via `verify_sources.py` (7/7 immutable master originals intact).
- **Worksheet Count & Navigation**:
  All four V8 company workbooks preserve the identical 16 worksheets in exact master order without addition, deletion, or renaming:
  `Model ->`, `Main Financials`, `Ratio Analysis`, `Intrinsic Growth`, `Comps`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`, `Data ->`, `compdata`, `Data Sheet`, `BS`, `Beta`, `RM`, `peerdata`, `Cash Flow`.

---

## 2. Deliverable Verification Matrix (Template V8)

All four priority targets have been built and recalculated cleanly in Microsoft Excel COM 16.0 under `companies/<TARGET>/template_v8/`. Prior versions (`template_v7`, `template_v6`, `template_v5`, `template_v4`, `template_v3`, `template_v2`) remain preserved intact on disk.

| Target Company | Valuation Status | Primary Valuation Output (`'DCF 1st way'!E42`) | Cover Link Cell (`'Model ->'!C9`) | Cover Upside / Downside (`'Model ->'!C13`) | 15-Oct-2025 Market Close | Raw Beta (`Beta!K7`) | Blume Beta (`Beta!K13`) | Evaluated WACC (`WACC-DCF!J39`) | Formula Errors (64 Sheets) | Balance Tie (All 12 Yrs) | Verified Workbook SHA-256 Checksum |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Mahindra & Mahindra Limited** (`MM`) | `Unavailable` | **Unavailable** | **Unavailable** | `N/A - Valuation Unavailable` | ₹3,497.20 | **1.4746** | **1.3164** | **11.97%** | **0 errors** | **0.00 Cr** | `3705342cc188f489628ad8d3b0345c932cc0ec7fecaec2f0458d08b766f9ea97` |
| **LTIMindtree Limited** (`LTM`) | `Supported` | **₹2,828.60** | **₹2,828.60** | -49.57% | ₹5,609.50 | **1.1526** | **1.1018** | **12.48%** | **0 errors** | **0.00 Cr** | `5ec03127d0535f7b7e1d05219a96a4cb873b08647960a68273166e1b8ea41e7d` |
| **Clean Science and Technology** (`CLEAN`) | `Supported` | **₹408.55** | **₹408.55** | -60.77% | ₹1,041.40 | **1.3212** | **1.2141** | **13.19%** | **0 errors** | **0.00 Cr** | `b868ea5e0b1350219ebd28a1e276d135898abf4d5e87e430ea5c05c5118e1796` |
| **Edelweiss Financial Services** (`EDELWEISS`) | `Unavailable` | **Unavailable** | **Unavailable** | `N/A - Valuation Unavailable` | ₹108.85 | **2.1231** | **1.7487** | **10.47%** | **0 errors** | **0.00 Cr** | `196d46b5e3d0283b0cba0f4f787cc9516d43b7093da3b16cef6c2b930a49c6e5` |

---

## 3. Fact vs. Analyst Assumption Boundary Audit

In strict compliance with user instructions, **company-specific facts absent from allowlisted sources cannot be relabeled as analyst assumptions**. Factual inputs without allowlisted evidence are withheld from calculations, and dependent valuations are marked `Unavailable`:

### A. Mahindra & Mahindra Limited (`MM`) — SOTP Valuation Unavailable
- **Missing Sourced Facts**: Auto segment EBITDA (₹11,700 Cr / ₹13,667 Cr), Farm Equipment segment EBITDA (₹6,716 Cr / ₹7,403 Cr), and Standalone Net Cash (+₹9,109 Cr) are statutory disclosure facts from investor presentations and standalone filings that are **not published on Screener.in consolidated financial statements**.
- **Action Taken**: Under the locked allowlist rule, these company-specific facts cannot be used in calculations or relabeled as analyst modeling assumptions.
- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, segment EBITDA rows (`E50`, `E51`), listed stakes (`E54`, `E55`), standalone cash bridge (`E56`, `F56`, `G56`), and single source-of-truth cell `E42` are set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Prior numerical SOTP fair value (₹2,742.97/sh) is removed from active valuation outputs.
- **Statements Supported**: Top-level 3-statement financials (Sales, Expenses, PAT, Assets, Liabilities) through FY25 are 100% sourced from Screener.in and tie out to `0.00 Cr`.

### B. Edelweiss Financial Services Limited (`EDELWEISS`) — SOTP Valuation Unavailable
- **Missing Sourced Facts**: Financial vertical subsidiary metrics (Alternatives PAT ₹375 Cr / ₹422 Cr, EARCL Book Value ₹1,960 Cr / ₹2,234.4 Cr, Mutual Fund AUM ₹1,50,000 Cr / ₹1,41,000 Cr, Insurance Embedded Value ₹2,100 Cr / ₹1,125 Cr) and Standalone Corporate HoldCo Net Debt (-₹3,067 Cr) are subsidiary statutory disclosures that are **not published on Screener.in**.
- **Action Taken**: Industrial FCFF DCF is methodologically prohibited for financial institutions where borrowings represent operating inventory. SOTP vertical metrics cannot be computed without unevidenced external facts.
- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, vertical financial rows `E50:E53`, HoldCo debt `E56`, and single source-of-truth cell `E42` are set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Prior numerical SOTP fair value (₹130.53/sh) is removed from active valuation outputs.
- **Statements Supported**: All balance sheet lines and sub-schedules (Borrowings, Other Liabilities, Other Assets) are 100% sourced from Screener.in popup schedules and tie out to `0.00 Cr`.

### C. Clean Science and Technology Limited (`CLEAN`) — Manufacturing FCFF DCF Supported (₹408.55 / sh)
- **Fact Remediation**: The unverified 10,500 MTPA Unit 4 HALS capacity figure has been completely removed from forecast calculations, worksheet titles (`'DCF 1st way'!B48`), and explicit cash flow labels (`B50`).
- **Valuation Build**: The 5-year explicit FCFF projection is driven strictly by explicit analyst revenue growth (`E19 = 14.0%`), reinvestment rate (`45.0%`), terminal growth (`5.0%`), ERP (`5.50%`), and marginal tax rate (`25.17%`) applied directly to Screener historical financials and Yahoo Finance 104-week regression Blume beta (`1.2141`).
- **Evaluated Output**: Full-Year Manufacturing FCFF DCF Fair Value evaluates cleanly to **₹408.55 / share** (Cover link: ₹408.55). Dynamic sensitivities (Terminal g: +₹19.40, ERP: -₹27.18) pass with exact restoration.

### D. LTIMindtree Limited (`LTM`) — Annual FCFF DCF Supported (₹2,828.60 / sh)
- **Fact Remediation**: IT operational metrics (headcount, employee utilization %) are quarterly presentation disclosures not published on Screener.in. Confirmed that no headcount or utilization schedules exist in the active 16-sheet template.
- **Valuation Build**: The 5-year explicit FCFF projection is driven strictly by explicit analyst revenue growth (`E19 = 11.0%`), reinvestment rate (`15.0%`), terminal growth (`5.0%`), ERP (`5.50%`), and marginal tax rate (`25.17%`) applied directly to Screener historical financials and Yahoo Finance 104-week regression Blume beta (`1.1018`).
- **Evaluated Output**: Full-Year Annual FCFF DCF Fair Value evaluates cleanly to **₹2,828.60 / share** (Cover link: ₹2,828.60). Dynamic sensitivities (ERP: -₹168.96, Terminal g: +₹137.05) pass with exact restoration.

---

## 4. Strict 7-Website Allowlist Compliance & Source vs. Assumption Classification

All sourced metrics are restricted exclusively to the 7 websites named in the master template plus Screener.in:

1. **Screener.in** (`https://www.screener.in/company/<TICKER>/consolidated/`): Historical P&L, Balance Sheet, Cash Flow, and Popup Sub-Schedules (FY15–FY25).
2. **Yahoo Finance** (`https://finance.yahoo.com`): 105 weekly adjusted closing price bars ending 15-Oct-2025; 104 paired weekly return observations vs NIFTY 50 (`^NSEI`).
3. **National Stock Exchange of India (NSE)** (`https://www.nseindia.com`): Official unadjusted closing equity quotations as of 15-Oct-2025 (`2025-10-15`).
4. **Trading Economics** (`https://tradingeconomics.com/india/government-bond-yield`): Dated prior-session benchmark 10-year sovereign bond yield proxy (6.51% as of 14-Oct-2025 preceding 15-Oct-2025 cutoff).
5. **Top Stock Research** (`https://www.topstockresearch.com`): Peer reference data cross-check.
6. **Investing.com** (`https://www.investing.com`): Secondary sovereign yield cross-check.
7. **Prime Investor** (`https://primeinvestor.in`): Secondary financial ratio cross-check.

*Strictly Zero Disallowed Domains*: Zero occurrences of BSE (`bseindia.com`), corporate investor relations portals, Damodaran Online, or RBI across all active registries, manifests, and documentation (enforced by `test_no_disallowed_domains_in_active_source_registries_or_manifests`).

### Sourced Facts vs. Explicit Analyst Assumptions
- **Sourced Facts**: Historical multi-year statement lines through FY25 (Screener.in); 105 weekly price bars (Yahoo Finance); 15-Oct-2025 closing stock prices (NSE / Yahoo Finance); 14-Oct-2025 10Y sovereign yield 6.51% (Trading Economics).
- **Explicit Editable Analyst Assumptions**:
  - Expected Revenue Growth Rate (`'DCF 1st way'!E19`): MM 12%, LTM 11%, CLEAN 14%, EDELWEISS 10%.
  - Terminal Perpetual Growth Rate (`'DCF 1st way'!E20`): 5.00% across all four targets.
  - Equity Risk Premium (`'WACC-DCF'!J23`): 5.50% (macroeconomic modeling assumption).
  - Marginal Corporate Tax Rate (`'WACC-DCF'!E23`, `'Data Sheet'!K29`): 25.17% (22% base + 10% surcharge + 4% cess).
  - Reinvestment Rates (`'DCF 1st way'!F11:K11`): LTM 15%, CLEAN 45%.
  - Blume Beta Weighting (`'Beta'!K8, K11`): 2/3 raw beta + 1/3 market beta (1.0).

---

## 5. Balance Sheet Integrity & Zero-Plug Reconciliation Audit

- **Exact Balance Check Formula**:
  - `BS!Row 40 = Row 38 - Row 16` across all 12 columns (`C40:N40`), evaluating `=C38-C16` through `=N38-N16`.
  - Row 38 (`Total Assets`) sums all asset components independently: Gross Block less Depreciation plus Capital WIP, Investments, Inventories, Trade Receivables, Cash & Equivalents, Short Term Loans, and Other Asset Items.
  - Row 16 (`Total Liabilities`) sums all liability and equity components independently: Equity Share Capital, Reserves & Surplus, Borrowings sub-lines, Trade Payables, Advance from Customers, and Other Liability Items.
- **Verification Across All 48 Company-Periods**:
  - Across all 4 targets (MM, LTM, CLEAN, EDELWEISS) and all 12 periods (`Mar 2015` to `Mar 2026E`), `ABS(Row 40) <= 0.01 Cr`, evaluating to exact **`0.00 Cr`** in Microsoft Excel COM.
- **Zero Residual Plugs Confirmed (`plug_free: true`)**:
  - **Historical Actuals (Cols C–M, FY15–FY25)**: Row 37 ("Other asset items") is sourced directly from Screener reported sub-schedules (or `0.0` for MM where Screener reports no Other Assets), with zero fallback (`tot_liab - curr_known` completely eradicated).
  - **Forecast FY26E (Col N)**: Row 37 is projected strictly from an independent operational driver (`=M37*1.05` for LTM, CLEAN, EDELWEISS; `0` for MM) with zero references to Total Liabilities (`N16`) or balancing tests.
  - **Cash Roll-Forward**: Closing Cash (`BS!N35`) is forecast strictly from `=M35+'Cash Flow'!N31` (opening cash plus CFO, CFI, and CFF), never calculated as a residual to force the balance sheet to tie.

---

## 6. Verified Deliverables, SHA-256 Checksums & Artifact Links

### 1. Mahindra & Mahindra Limited (`MM`)
- **16-Sheet Workbook**: [`companies/MM/template_v8/MM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v8/MM-FullTemplate.xlsx)
  - SHA-256: `3705342cc188f489628ad8d3b0345c932cc0ec7fecaec2f0458d08b766f9ea97` (114,718 bytes)
  - Valuation Status: **Unavailable** (Cover link: `Unavailable`, Upside: `N/A - Valuation Unavailable`)
- **Package Manifest**: [`companies/MM/template_v8/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v8/package.manifest.json)
- **Source Registry**: [`companies/MM/source_registry.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/source_registry.json) / [`docs/MM_SOURCE_REGISTRY.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/MM_SOURCE_REGISTRY.md)
- **Sheet Previews (16 Landscape PDFs in `companies/MM/template_v8/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model ->` | `previews/01_Model_-_.pdf` | 36,362 | `9bf258d3e3ab78de2d10de84443625a283ec4b40253bf6a1676e40ffcba98eba` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 247,636 | `41569bee303489698b9fdcb5f6ecf7338acd8b4112ac476032be016654af37ab` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 190,859 | `9d386f3492c48df390b1cda82c55f9ce792b284242a272e98d2fc53ac46b3b72` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 293,186 | `45af49b08f0b30f3da9cad5859b9e5dcfad9fb0e3a0e23c9eea58e19261c3c70` |
| `Comps` | `previews/05_Comps.pdf` | 275,944 | `b679c1616a3d0944337c5ea61eb8ecac146014aea365f5bbac2337bef13ca124` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 271,635 | `0d59fbb270b8dd5f94333d6b6018f89e3a2687842d20c54b1b6fe6ade3c95d77` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 314,379 | `2e12a57214f8ad17f0a2f1532d5b5010d63e7f4d46ac1477370a408d06586004` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 352,468 | `cd6eb94e24c60a5a984d5b3d0f83b388418b18f80cae32bc2a7d3bf1219a279f` |
| `Data ->` | `previews/09_Data_-_.pdf` | 4,846 | `a9c3b3ca6ce3dd54cead788d91618e2f718ea7c41d9dee3052bd71d8d5ddc719` |
| `compdata` | `previews/10_compdata.pdf` | 47,260 | `c280a5e5d7947d445c91054eef4bb0e1d0da6f47e5812587aaab4dd2f86f7872` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 81,687 | `950e1fd3df5f2946a8c93f07841b294481cf32daa1d9ee3e7e5553046cdc499e` |
| `BS` | `previews/12_BS.pdf` | 186,099 | `31625773361172c2e243467d2d5599b9c25e6c81e3d34eec74eb5ba1c6a3af62` |
| `Beta` | `previews/13_Beta.pdf` | 320,575 | `4499da0db3983d49ab969bd9e2d6c36c2280a0c9a936e9b25e0c3a159cfb43dd` |
| `RM` | `previews/14_RM.pdf` | 35,831 | `f83ec0009c31622dc4d871a2fb549b66521ef206750cf32c04cc15cbeee44edb` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,069 | `0e98fbeec1a16f58fe316233742b1835074a45ec747b22e31c5403890950c717` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 54,234 | `5263f68208c366e65bdd80f1e295da9cd30ed6a39dc6fb6385804b2be1c1cf32` |

### 2. LTIMindtree Limited (`LTM`)
- **16-Sheet Workbook**: [`companies/LTM/template_v8/LTM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v8/LTM-FullTemplate.xlsx)
  - SHA-256: `5ec03127d0535f7b7e1d05219a96a4cb873b08647960a68273166e1b8ea41e7d` (115,285 bytes)
  - Valuation Status: **Supported** (Cover link: **₹2,828.60 / share**, Upside: -49.57%)
- **Package Manifest**: [`companies/LTM/template_v8/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v8/package.manifest.json)
- **Source Registry**: [`companies/LTM/source_registry.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/source_registry.json) / [`docs/LTM_SOURCE_REGISTRY.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/LTM_SOURCE_REGISTRY.md)
- **Sheet Previews (16 Landscape PDFs in `companies/LTM/template_v8/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model ->` | `previews/01_Model_-_.pdf` | 35,945 | `d760f24bb1c6911cff53d539fc030e0fa8275130896a497e3fe3f2fa5464068a` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 244,996 | `cbe8df589402daa7ac5c9e63194945cdb632c99a0017e9e14ac2758e1fccb796` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 189,392 | `75f60875ea80abc7a7562e61a020eaa1261f3dec21419bf396d42cb258b0399f` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 291,876 | `3f8bf2592402ea7a33401a17fc3485f7d744f1a07b1389aa6da6b0c4c6a6b993` |
| `Comps` | `previews/05_Comps.pdf` | 275,953 | `79646f0bf328668806009acf76d7908e33ab54077e5a7866b95101598de5ca5c` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 270,438 | `0d9e5e5f55e45e9d08fa4ac0fdcf1625d225e77fe8cfd5f9278e83b21006b315` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 332,035 | `91a05755be9fec4e4d93ea64bea481bd3b77e986c3f742596706ee217df5538f` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 350,953 | `8baea91ebaf82a5d1780cdac292c9a8b3905647d55f8e1df3e3f81ba38de6507` |
| `Data ->` | `previews/09_Data_-_.pdf` | 4,846 | `acb0a4b91bc3e6d4d5cde509fea1e634d2726eec0088560146f2aef5342795d2` |
| `compdata` | `previews/10_compdata.pdf` | 47,577 | `2e0b5bec06714a61b3e5bf44cfd1eebf4819bfdc453f6ffd51d8719ef3f7672f` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 80,924 | `f1477c79f886873e14c9e9a5ec6f1d0c8f5a14a63751ac83b2d00aaf37f4d490` |
| `BS` | `previews/12_BS.pdf` | 194,018 | `7557e1a7ea69355afb3e0826a51982517dd740d028d435e2932f738f23977135` |
| `Beta` | `previews/13_Beta.pdf` | 321,608 | `7429c6f5f3916f0d08898fa134d3de932bcc70f67dad7c6eda5948e203dd5760` |
| `RM` | `previews/14_RM.pdf` | 35,831 | `0f2f98e07b48dc8d9185a7e3edc31bffbb96e53514fbb6730edbc5134beb036c` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,057 | `44a131f1cfb39243e44597e8ff47385fb00e5360e785ab9e1d346ab46ff5604a` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 60,648 | `921e92c663b9aeef43309513b35ab5da2783fb45f5d780b0fe9153350fcfd35e` |

### 3. Clean Science and Technology Limited (`CLEAN`)
- **16-Sheet Workbook**: [`companies/CLEAN/template_v8/CLEAN-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v8/CLEAN-FullTemplate.xlsx)
  - SHA-256: `b868ea5e0b1350219ebd28a1e276d135898abf4d5e87e430ea5c05c5118e1796` (114,100 bytes)
  - Valuation Status: **Supported** (Cover link: **₹408.55 / share**, Upside: -60.77%)
- **Package Manifest**: [`companies/CLEAN/template_v8/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v8/package.manifest.json)
- **Source Registry**: [`companies/CLEAN/source_registry.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/source_registry.json) / [`docs/CLEAN_SOURCE_REGISTRY.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/CLEAN_SOURCE_REGISTRY.md)
- **Sheet Previews (16 Landscape PDFs in `companies/CLEAN/template_v8/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model ->` | `previews/01_Model_-_.pdf` | 36,105 | `d1785018fd90d0e81ac629cfa844b9a8963291a295b95d2834bf71966c0bd9e1` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 248,819 | `07712414bb74f6511284cca8e9520eea7359994918ede9aa23a536eef420604a` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 191,758 | `3e57a55c0282fc83de329fa0e253c9501e697d785ec1c7894f5d0a3511b751cb` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 291,285 | `46dbd8b97ca9b4be63c0b2be86df69275fd93d371288c2e5a7fdb6a69d30214c` |
| `Comps` | `previews/05_Comps.pdf` | 273,941 | `ff172341543c16ee0e23a389664ea08b9b83b6a941b799694ac6ca6f4327e00e` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 272,189 | `8ad9710dffa32352aa8e691ebfe1c0106e19f9f7d6fa1ae9bf1fec9849706817` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 330,181 | `d34e7c1b4bec48fdcef2232f81dfe2c12936f8619dbe2ac9d21325e116fa2c1c` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 351,788 | `6f5e0f2be52930e264827b7980b011f74fb8fd5d9f760d60824ac8b31d77e266` |
| `Data ->` | `previews/09_Data_-_.pdf` | 4,846 | `1dd1481bebf863c0bd8f897b7ee7b0f96a666547c8d3e56e9f02d63fac185308` |
| `compdata` | `previews/10_compdata.pdf` | 47,685 | `3006df287fdb66f6c1c00abd4f67eafdcd0131f8885d556ef78d4bd14debd649` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 80,823 | `892ce458f003fe9a5162b3cb04a5b287f3cba89a3951e8f4225a704bf252e595` |
| `BS` | `previews/12_BS.pdf` | 191,905 | `5c983d4dfe533ce37d085ef4a0cad7697c1554e59a7e1e3200edb55f8ac6b324` |
| `Beta` | `previews/13_Beta.pdf` | 324,661 | `7ef46876f9e79278789fe6d5433567e47423c566c41330e20d3dbcdcea4ed24f` |
| `RM` | `previews/14_RM.pdf` | 35,831 | `46ddcceba145f0fccaeb652197500cff52b205b5aecb36cc4b8ab9f529c851ba` |
| `peerdata` | `previews/15_peerdata.pdf` | 35,599 | `d8b2638dde9d0587e410a0251af88cb4f353db64d70abdc018cfc661f812e4e0` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 58,253 | `bf689dc0fdfe0c660a7a4d575e0a968f64032bdd44cd2e29a6ab76befa6b7a73` |

### 4. Edelweiss Financial Services Limited (`EDELWEISS`)
- **16-Sheet Workbook**: [`companies/EDELWEISS/template_v8/EDELWEISS-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v8/EDELWEISS-FullTemplate.xlsx)
  - SHA-256: `196d46b5e3d0283b0cba0f4f787cc9516d43b7093da3b16cef6c2b930a49c6e5` (117,718 bytes)
  - Valuation Status: **Unavailable** (Cover link: `Unavailable`, Upside: `N/A - Valuation Unavailable`)
- **Package Manifest**: [`companies/EDELWEISS/template_v8/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v8/package.manifest.json)
- **Source Registry**: [`companies/EDELWEISS/source_registry.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/source_registry.json) / [`docs/EDELWEISS_SOURCE_REGISTRY.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/EDELWEISS_SOURCE_REGISTRY.md)
- **Sheet Previews (16 Landscape PDFs in `companies/EDELWEISS/template_v8/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model ->` | `previews/01_Model_-_.pdf` | 35,742 | `9edfd0dd3ef32a1109bcccdd3341b975c041399744406bfc2f1e7906d5366a88` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 249,129 | `f018c3903712155ab8cf260f0a7a2dabead38fcd36a284257e2387ea04d3e490` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 191,386 | `a98bf45a1ddbc4c4d08a761cb97bc2ab20f8ae895239e0ec300668d824618f65` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 300,650 | `7b99e797da4bcdb255914567ddb361b85d69bb32c1728deb61fbf801225ec47b` |
| `Comps` | `previews/05_Comps.pdf` | 272,990 | `6e479ac41a69b375b6a1e11857a78d80430dbbf95e31cf2ac1c2088ce1aee1b5` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 275,134 | `639b5aa6f4170adffe4e6fda02d42ad039504e87517150975cdb9c61324d0096` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 313,676 | `8a5fb6dc653ff7cee8c56b08bb14a191ebd56de99c234f09313b2c4e028766d9` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 353,248 | `693b1b47898d161f379b81aec1f6fedd86f0d731f032f53b35c82ca540d739fd` |
| `Data ->` | `previews/09_Data_-_.pdf` | 4,846 | `6b3aceea8041e32d76e9c5d5fd5e9ad41fd9ba00452f14e8648fb4aef9dcc14a` |
| `compdata` | `previews/10_compdata.pdf` | 47,653 | `60fd20c59facffaa3c7e441e2e84cf9cf88adc922fc912ad426e352ed660160f` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 80,952 | `450d0e33b5a1011ca6f1d5434dfbda5e130e38c4d8b35277252e3fb79800a71a` |
| `BS` | `previews/12_BS.pdf` | 194,197 | `0c5864c182013a478e1ba165097d8b246c8cd1ff347ec420fcbfc690b3d773f2` |
| `Beta` | `previews/13_Beta.pdf` | 323,888 | `5e42697370f27ee2ccdbe33ddd01c61d77e070cbb03bd7cedd9c125b179e3dbb` |
| `RM` | `previews/14_RM.pdf` | 35,831 | `c66300f41258cab4d67d7af20d9ef88212a0be52993cb8196ab0bd9659e796ac` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,282 | `da8991dd6f67635084039a1b7d109c0178ae0032fb77c1b66376a43c352556a6` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 60,770 | `3d8cc0dec37b71f0ff1d805d8c86dbc5f3cbda94f62cb19bc974b365b4ec367e` |

---

## 7. Verification Suite & Release Gate Record

| Verification Check | Standard / Command | Result | Gate Status |
|:---|:---|:---:|:---:|
| **Master Source Preservation** | `python verify_sources.py` | 7/7 master originals verified byte-identical | **PASSED** |
| **Disallowed Domain Rejection** | `pytest test_package_safety.py` | Zero disallowed domains across active registries/manifests | **PASSED** |
| **Full Regression Suite** | `pytest` | **94/94 tests passing** (100% pass rate in 28.40s) | **PASSED** |
| **Excel COM Recalculation** | `python work/recalculate_template_v8_models.py` | Exactly 0 formula errors across all 64 sheets | **PASSED** |
| **Balance Sheet Integrity** | `BS!Row 40 = Row 38 - Row 16` | Exactly `0.00 Cr` across all 48 company-periods | **PASSED** |
| **Dynamic Sensitivities** | LTM & CLEAN perturbation / restoration | Live responsiveness with exact restoration to baseline | **PASSED** |
| **Human Signoff Gate** | User review of Template V8 deliverables | `ready: false`, `human_pack_approval: pending` | **PENDING** |
