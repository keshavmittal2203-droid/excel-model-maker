# Human Review & Approval Packet: 16-Sheet Master Template Excel Models (Template V9 Release)

> [!IMPORTANT]
> **HISTORICAL VALUATION SNAPSHOT NOTICE**: All financial models, closing stock prices, market capitalizations, Treasury yields, peer trading multiples, and empirical regression datasets in this release reflect historical market observations as of **Wednesday, October 15, 2025 (`2025-10-15`)**. They are **NOT** current or live market valuations and do not incorporate subsequent market events.
>
> **APPROVAL GATE STATUS**: In strict adherence to repository safety invariants (`DISALLOWED_REVIEWER_TOKENS` prohibiting agent self-promotion), the **`human_pack_approval` gate remains PENDING (`ready: false`)**. These models are delivered in proposal/validation state pending explicit sign-off by the human reviewer. PR #3 remains open and unmerged.

---

## 1. Master Template Architecture & Governance

- **Active Master Template Architecture**:
  - **Archived Immutable Original Master**: Exact 17-sheet original permanently preserved at [`originals/archive/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/archive/Marksans%20Keshav.xlsx) (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`, 17 sheets).
  - **Active Working Master**: Structural removal of obsolete `sectordata` sheet applied directly to [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/Marksans%20Keshav.xlsx) (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`, **16 sheets**).
  - Both original templates verified byte-for-byte via `verify_sources.py` (2/2 immutable master originals intact).
- **Worksheet Count & Navigation**:
  All four V9 company workbooks preserve the identical 16 worksheets in exact master order without addition, deletion, or renaming:
  `Model ->`, `Main Financials`, `Ratio Analysis`, `Intrinsic Growth`, `Comps`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`, `Data ->`, `compdata`, `Data Sheet`, `BS`, `Beta`, `RM`, `peerdata`, `Cash Flow`.

---

## 2. Deliverable Verification Matrix (Template V9)

All four priority targets have been built and recalculated cleanly in Microsoft Excel COM 16.0 under `companies/<TARGET>/template_v9/`. Prior versions (`template_v8`, `template_v7`, `template_v6`, `template_v5`, `template_v4`, `template_v3`, `template_v2`) remain preserved intact on disk.

| Target Company | Valuation Status | Primary Valuation Output (`'DCF 1st way'!E42`) | Cover Link Cell (`'Model ->'!C9`) | Cover Upside / Downside (`'Model ->'!C13`) | 15-Oct-2025 Market Close | Raw Beta (`Beta!K7`) | Blume Beta (`Beta!K13`) | Evaluated WACC (`WACC-DCF!J39`) | Formula Errors (64 Sheets) | Balance Tie (All 12 Yrs) | Verified Workbook SHA-256 Checksum |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Mahindra & Mahindra Limited** (`MM`) | `Unavailable` | **Unavailable** | **Unavailable** | N/A - Valuation Unavailable | ₹3,497.20 | **1.4746** | **1.3164** | **11.97%** | **0 errors** | **0.00 Cr** | `57ab161c7c85c9258e1d6d57b5babe513c33cb7a84683d91273c9557c7e902d1` |
| **LTIMindtree Limited** (`LTM`) | `Supported` | **₹2,828.60** | **₹2,828.60** | -49.57% | ₹5,609.50 | **1.1526** | **1.1018** | **12.48%** | **0 errors** | **0.00 Cr** | `9189d9c663d16bf160f3b3a32925c18dd18571007b1fa1ef1fafa092b11c9b5d` |
| **Clean Science and Technology Limited** (`CLEAN`) | `Supported` | **₹408.55** | **₹408.55** | -60.77% | ₹1,041.40 | **1.3212** | **1.2141** | **13.19%** | **0 errors** | **0.00 Cr** | `c37fbb8935112e534f4f614bda9ec98b90ece60bbd73ed7258013ccd93ea033c` |
| **Edelweiss Financial Services Limited** (`EDELWEISS`) | `Unavailable` | **Unavailable** | **Unavailable** | N/A - Valuation Unavailable | ₹108.85 | **2.1231** | **1.7487** | **10.47%** | **0 errors** | **0.00 Cr** | `5b8a544f153f85a395d9af767ed8cd622ca38b23417c870d69133e13cdd06429` |

---

## 3. Fact vs. Analyst Assumption Boundary Audit

In strict compliance with user instructions, **company-specific facts absent from allowlisted sources cannot be relabeled as analyst assumptions**. Factual inputs without allowlisted evidence are withheld from calculations, and dependent valuations are marked `Unavailable`:

### A. Mahindra & Mahindra Limited (`MM`) — SOTP Valuation Unavailable
- **Missing Sourced Facts**: Auto segment EBITDA (₹11,700 Cr / ₹13,667 Cr), Farm Equipment segment EBITDA (₹6,716 Cr / ₹7,403 Cr), and Standalone Net Cash (+₹9,109 Cr) are statutory disclosure facts from investor presentations and standalone filings that are **not published on Screener.in consolidated financial statements**.
- **Action Taken**: Under the locked allowlist rule, these company-specific facts cannot be used in calculations or relabeled as analyst modeling assumptions.
- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, single source-of-truth cell `E42` is set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Secondary Comps P/E peer cross-check is presented in C18:D18 (₹4,274.65) with a cautionary note in C19.
- **Statements Supported**: Top-level 3-statement financials (Sales, Expenses, PAT, Assets, Liabilities) through FY25 are 100% sourced from Screener.in and tie out to `0.00 Cr`.

### B. Edelweiss Financial Services Limited (`EDELWEISS`) — SOTP Valuation Unavailable
- **Missing Sourced Facts**: Financial vertical subsidiary metrics (Alternatives PAT ₹375 Cr, EARCL Book Value ₹1,960 Cr, Mutual Fund AUM ₹1,50,000 Cr, Insurance Embedded Value ₹2,100 Cr) and Standalone Corporate HoldCo Net Debt (-₹3,067 Cr) are subsidiary statutory disclosures that are **not published on Screener.in**.
- **Action Taken**: Industrial FCFF DCF is methodologically prohibited for financial institutions where borrowings represent operating inventory. SOTP vertical metrics cannot be computed without unevidenced external facts.
- **Valuation Impact**: SOTP valuation cannot be computed from allowed sources. In `'DCF 1st way'`, single source-of-truth cell `E42` is set to `Unavailable`. Cover sheet `'Model ->'!C9` evaluates to `Unavailable`; cover upside/downside `C13` displays `N/A - Valuation Unavailable`. Secondary Comps P/E peer cross-check is presented in C18:D18 (₹90.65) with a cautionary note in C19.
- **Statements Supported**: All balance sheet lines and sub-schedules (Borrowings, Other Liabilities, Other Assets) are 100% sourced from Screener.in popup schedules and tie out to `0.00 Cr`.

### C. Clean Science and Technology Limited (`CLEAN`) — Manufacturing FCFF DCF Supported (₹408.55 / sh)
- **Fact Remediation**: The unverified 10,500 MTPA Unit 4 HALS capacity figure has been completely removed from forecast calculations, worksheet titles, and explicit cash flow labels.
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

---

## 5. Template V9 Defect Remediation Summary

| Defect Area | Description of Issue in V8 | Remediation Applied in V9 | Status |
|:---|:---|:---|:---:|
| **Wordmark Branding** | Top header in `Main Financials` displayed Marksans Pharma branding | Linked `Main Financials!B3` dynamically to `='Data Sheet'!B1` in 16pt bold `#1F497D` | `FIXED` |
| **Fiscal Year Dates** | Fiscal headers stored as strings with `00:00:00` timestamps | Populated with true `datetime.date(YYYY, 3, 31)` and explicit Excel number formats | `FIXED` |
| **Cell Comments** | Scattered openpyxl cell comments across multiple sheets | Purged 100% of openpyxl cell comments across all 64 sheets | `FIXED` |
| **Data Sheet Cash Flows** | Rows 82:85 repeated identical numbers across all 4 companies | Linked dynamically to `Cash Flow!E5:N5` through `E8:N8`, reconciling all historical statements | `FIXED` |
| **Share Count** | Data Sheet K70 repeated Marksans 453M share count | Set K70 to exact target shares (MM: 1.24B, LTM: 296M, CLEAN: 106M, EDELWEISS: 948M) | `FIXED` |
| **Intrinsic Growth NCA** | Gross Block Row 32 evaluated to 0, producing negative NCA and distorted ROIC | Set Row 32 to `=IF(SUM(E22:E31)>0, SUM(E22:E31), BS!J28)`; Net NCA is positive and ROIC realistic | `FIXED` |
| **Comps Formula Portability** | Modern `_xludf.IFS` function created compatibility risk in Excel 2016/2019 | Replaced with standard portable nested `IF` statements; cleared 166 stray borders | `FIXED` |
| **DCF Sensitivity Units** | Sensitivity table calculated EV instead of Rupees per share | Re-anchored to per-share `((PV_FCFF + PV_TV + Cash - Debt)/Shares)`; base cell J27 ties to E42 | `FIXED` |
| **DCF 2nd Way Dates** | Naive `+365` day addition caused fiscal year dates to drift | Pinned dates to March 31 via `=DATE(YEAR(C70)+1, 3, 31)` | `FIXED` |
| **Beta Formula** | Dynamic array formula `_xlfn.COVARIANCE.S` caused compatibility risk | Replaced with standard `=SLOPE(D9:D112, G9:G112)` | `FIXED` |
| **Ratio Analysis Sparklines** | openpyxl stripped Office 2010 sparkline `<x14:sparklineGroups>` XML | Post-build ZIP injector cleanly re-binds XML with `xmlns:xr2` namespace | `FIXED` |
| **Print Pagination** | Multiple multipage splits on `Main Financials`, `Comps`, and `DCF2` | Enforced `FitToPagesWide = 1`, `FitToPagesTall = False`, landscape; blank `Data ->` excluded | `FIXED` |

---

## 6. Verified Deliverables, SHA-256 Checksums & Artifact Links

### Mahindra & Mahindra Limited (`MM`)
- **16-Sheet Workbook**: [`companies/MM/template_v9/MM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v9/MM-FullTemplate.xlsx)
  - SHA-256: `57ab161c7c85c9258e1d6d57b5babe513c33cb7a84683d91273c9557c7e902d1` (116,658 bytes)
  - Valuation Status: **Unavailable** (Primary: **Unavailable**, Cover: **Unavailable**, Upside: `N/A - Valuation Unavailable`)
  - Note: Secondary Comps P/E Peer Cross-Check: ₹4,274.65 / share displayed for peer context with cautionary note.
- **Package Manifest**: [`companies/MM/template_v9/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v9/package.manifest.json)
- **Sheet Previews (15 Landscape PDFs in `companies/MM/template_v9/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model - ` | `previews/01_Model_-_.pdf` | 35,891 | `8be626c8d1cb6fc3f38ff3f82519360a7e89c6d5d8a50cfd2ff5bd04a3db28a9` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 250,602 | `7485031f2784785ac338dcee3dd9f802534a1ca04d72c31d302cdb577eb91087` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 257,329 | `8caf77abe9d9b248d9e2d49ed26ae956d960ecdecdd9bc0f839ea652325e7605` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 294,587 | `2539dfa708bf15dccd53277c8383c07729844585caeff6b6e505ecc6b3a356a8` |
| `Comps` | `previews/05_Comps.pdf` | 284,876 | `829975319c0008024a8bcb4bc74e7c6f20772a6a3084caf43e7bbe8a0bbd4e58` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 267,516 | `9d620b549453f8c7bb9eee3b982e539430a5e22696ac4141e339b0cf29a745c9` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 314,638 | `c432d2d1f5c3bedacf492eb876e8c7210add50e7da99216dd75c343fff8f2138` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 342,875 | `82ca434b945a68885f07ee3d66e266ac50ba3f64d73e7967f947772985ced567` |
| `compdata` | `previews/10_compdata.pdf` | 45,929 | `b72c1e5877df970bafe67e14ceb4472cb9d45099aaaf7a8b8a1cd5c19de1101f` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 78,300 | `25b93c02e9bb5b2a56fb8cae7e402312555af2df90c3379c2cfe457aa65956f0` |
| `BS` | `previews/12_BS.pdf` | 134,377 | `7d5c61a2a5bb0cbe8a73657ecbed5560b83a9727e116ae1c4947964577c2d7b2` |
| `Beta` | `previews/13_Beta.pdf` | 316,625 | `f5ac6c92d28c8433cce76bd4516235de9dbfc2d435af3c0cfef5f7e167f25a03` |
| `RM` | `previews/14_RM.pdf` | 35,830 | `82b7fa6a18edc6ec2fa27b54caccbe10b14346bfd5d7c03d53644078c708d347` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,076 | `968181a00d3ab0bc32244fb992a36fe871b3a38472e561c7efcf91dc7a40c4fe` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 53,467 | `26a8baa7421af8691f32da2d39642b774a5048df6f717cfbba7401bd1a434769` |

### LTIMindtree Limited (`LTM`)
- **16-Sheet Workbook**: [`companies/LTM/template_v9/LTM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v9/LTM-FullTemplate.xlsx)
  - SHA-256: `9189d9c663d16bf160f3b3a32925c18dd18571007b1fa1ef1fafa092b11c9b5d` (117,235 bytes)
  - Valuation Status: **Supported** (Primary: **₹2,828.60**, Cover: **₹2,828.60**, Upside: `-49.57%`)
  - Note: Fully supported full-year FCFF DCF baseline; quarterly presentation headcount/utilization metrics excluded.
- **Package Manifest**: [`companies/LTM/template_v9/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v9/package.manifest.json)
- **Sheet Previews (15 Landscape PDFs in `companies/LTM/template_v9/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model - ` | `previews/01_Model_-_.pdf` | 35,050 | `cd92f385430d6daa2f99f2122b0da8e23f7a753089c0ffdb4aab12c6412cbd54` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 247,832 | `32ef74d6e7e502ec2df73b7449dc41d9448feaf8ee4bd00802fce7eeb23b63ac` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 257,240 | `48cc47958f62343ea5455dcb6c3c2d2cf7f189bd3fad13fe0fc4a702927af436` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 292,536 | `4fad87ef84175e760a991540e5171453d2cd4be13d3f3c132ec8185b313e72bc` |
| `Comps` | `previews/05_Comps.pdf` | 283,996 | `a0a95d988d63993999e7a2a068163e02f6c6ea18c3385b0b3e993906d1d09f38` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 266,290 | `17cf51a074a54fa5a77078fb21f414e54a0473ae3a7c65aceb0bcb9b1575df4d` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 331,413 | `6e50350b535e3f5099e2b5cd40bcf1f9bf1d937daeb93d98976c497e868e0367` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 342,283 | `253ef91bbd652b1a3866b87c3906ef8990e36e376682dc72e40cb1663e10090c` |
| `compdata` | `previews/10_compdata.pdf` | 46,301 | `0a42f18b569ff900f90edf65a8a44c511a553e0871f28b901b387e22a753564e` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 77,559 | `5d8994e580ca9ce976603550acdc87bc223f83480b8dfe665fe6dc97faf94676` |
| `BS` | `previews/12_BS.pdf` | 142,022 | `e7f09b0360594609ca5989d905ac842c1c816724905ec43c4a84f2327a28fc18` |
| `Beta` | `previews/13_Beta.pdf` | 317,834 | `237a21bb7bfebe835def640635b5a8b81894957e7e50f30661a4276e73599483` |
| `RM` | `previews/14_RM.pdf` | 35,830 | `3b5191382ef265b6e276e964bc1ce801bb2877ca6063f2b5d32a9186bcfa3c77` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,062 | `65641f35570ec0d94133d524c05c4e779551232f8056010ba8362b33fb77c171` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 59,796 | `abbc8ebb0b7f2569433ce558c45721f5b925a12bacc7fa00b2a558757da3f0af` |

### Clean Science and Technology Limited (`CLEAN`)
- **16-Sheet Workbook**: [`companies/CLEAN/template_v9/CLEAN-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v9/CLEAN-FullTemplate.xlsx)
  - SHA-256: `c37fbb8935112e534f4f614bda9ec98b90ece60bbd73ed7258013ccd93ea033c` (115,948 bytes)
  - Valuation Status: **Supported** (Primary: **₹408.55**, Cover: **₹408.55**, Upside: `-60.77%`)
  - Note: Fully supported manufacturing FCFF DCF baseline; unverified 10,500 MTPA capacity figure removed.
- **Package Manifest**: [`companies/CLEAN/template_v9/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v9/package.manifest.json)
- **Sheet Previews (15 Landscape PDFs in `companies/CLEAN/template_v9/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model - ` | `previews/01_Model_-_.pdf` | 35,309 | `b31421e8bbdd9b7ee0ee6733ec124757f89623ce28ac5c40a5edd3b705dd5eec` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 251,058 | `56e36becb7a03098517263dbc986ecb1d405580588968ea193932881b29440b8` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 258,174 | `524cc1e63a60b93929096aa9d6679f1fd6aad7b8574fec625cb9f68a230ac153` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 292,629 | `ea497d9974fdc346a45b959d34af4c04bb6608c11eb8b715b5700baa0b4bd6eb` |
| `Comps` | `previews/05_Comps.pdf` | 282,873 | `a27a48633da8875b3e5e7ed1281337969b01f54c332d34701d585cbc9a58598a` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 268,057 | `ca2723b075a6e45efb40de7d6a0e1ff3f3ef04766daacfcffd2da9aa1ef15439` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 330,173 | `83caee8c612959ef129ea4df523bb2f9494c5fb01b94c68a4a0fe42ea6a043d9` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 343,270 | `5851abf8afc8ae6beb6fdc540032ccf71372a7dc469390dec18d21d84b27eafd` |
| `compdata` | `previews/10_compdata.pdf` | 46,370 | `94cb32996db8a5718addb468fb5df87aed26de4d2e1acc05ddcdce1e27250853` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 77,371 | `ca77312a401816da14228602782c80785ce38cda30caf59d675ee1332278ea9d` |
| `BS` | `previews/12_BS.pdf` | 139,997 | `9742ed4ad963c072754d1a9456ebbcd81b61aec25188b3f28cd70cdf036136d5` |
| `Beta` | `previews/13_Beta.pdf` | 320,699 | `ca597531c4ad36f10f25133e3d85bd86a6cfe05f28e17f1f3fd3b0eb75d8824c` |
| `RM` | `previews/14_RM.pdf` | 35,830 | `ffa41b978ab5227a82d4cec3289ebb9fddff378dec12fc2fa33d6356becef919` |
| `peerdata` | `previews/15_peerdata.pdf` | 35,606 | `edf4707b6967b4774f09368b9174d5533ae6991668f8da3ee232920b25323dc6` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 57,448 | `ed99c1d8f65801d5ff132ebc8739c325f0bbe37b121067aa8bc86aa4619e469a` |

### Edelweiss Financial Services Limited (`EDELWEISS`)
- **16-Sheet Workbook**: [`companies/EDELWEISS/template_v9/EDELWEISS-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v9/EDELWEISS-FullTemplate.xlsx)
  - SHA-256: `5b8a544f153f85a395d9af767ed8cd622ca38b23417c870d69133e13cdd06429` (115,537 bytes)
  - Valuation Status: **Unavailable** (Primary: **Unavailable**, Cover: **Unavailable**, Upside: `N/A - Valuation Unavailable`)
  - Note: Industrial FCFF prohibited for financials; vertical segment inputs absent from allowlisted sources. Secondary Comps P/E Cross-Check: ₹90.65 / share displayed for peer context.
- **Package Manifest**: [`companies/EDELWEISS/template_v9/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v9/package.manifest.json)
- **Sheet Previews (15 Landscape PDFs in `companies/EDELWEISS/template_v9/previews/`)**:

| Sheet Name | PDF File | Size (Bytes) | SHA-256 Checksum |
|:---|:---|:---:|:---|
| `Model - ` | `previews/01_Model_-_.pdf` | 35,848 | `25989661777684a599202734f9929ac1ab3c846c8db801fcb8c3b2e875abf8d8` |
| `Main Financials` | `previews/02_Main_Financials.pdf` | 251,593 | `36f9d5b2e56883b2d33b85a4e3b8d76816fffd637e48a8b50b95e7049d2c3424` |
| `Ratio Analysis` | `previews/03_Ratio_Analysis.pdf` | 260,154 | `e67540b2c137a60ce7dbe292802858bc66d1875093a762c427988baa51bc0995` |
| `Intrinsic Growth` | `previews/04_Intrinsic_Growth.pdf` | 296,610 | `f3f93c7804b3d96452e126c09ff3b13fdab3ea3082a2cd788ac052b8a5ed16db` |
| `Comps` | `previews/05_Comps.pdf` | 281,193 | `f5d885a35328100ef34eb7767db8fc844f718f6a2495aadf58761846dca6903d` |
| `WACC-DCF` | `previews/06_WACC-DCF.pdf` | 270,751 | `e406406446cab97e077e1dbc1419ed6b617dbee7e2bab63d098a4159bd9e8b1a` |
| `DCF 1st way` | `previews/07_DCF_1st_way.pdf` | 313,679 | `f095b9539fdac189e749930211f175a3a0fef2bd58a2eeeae10e992723bb0b3a` |
| `DCF 2nd way` | `previews/08_DCF_2nd_way.pdf` | 344,842 | `201e0e2c5ede63498f51bbd30b4cf04c2e6bc6672fc2ef92cf112acdf911ee42` |
| `compdata` | `previews/10_compdata.pdf` | 46,229 | `8d707bb7355285eeeed005866e5b81c70680918eff0c6c6f6464bc691b783071` |
| `Data Sheet` | `previews/11_Data_Sheet.pdf` | 77,579 | `f998dac03076c929319ee3147903862adca8fe02121e1c72f4506719d06d876a` |
| `BS` | `previews/12_BS.pdf` | 142,148 | `85d77cc8ab9b0d0466ffc4ee3d1072d6961fdd53b29400c5f3c120bcea787da0` |
| `Beta` | `previews/13_Beta.pdf` | 319,946 | `0def1d9d672b0ea8f502dca75ded579ba79f94defbc9d89073141e5b8bdc34de` |
| `RM` | `previews/14_RM.pdf` | 35,830 | `145566897ee0f8b0019b08df9269722b1b5249dc0b0a50ac57ed3e28e0494858` |
| `peerdata` | `previews/15_peerdata.pdf` | 36,290 | `964c01ca141cf15c26c765d4bb006ee68daf315cd6ec94789fead9ad36457547` |
| `Cash Flow` | `previews/16_Cash_Flow.pdf` | 60,011 | `80d79f02bc623ce5255578bfc9c42820c5a9bb57ccd6eb5c3b07c640f5cda5b8` |

---

## 7. Verification Suite & Release Gate Record

| Verification Check | Standard / Command | Result | Gate Status |
|:---|:---|:---:|:---:|
| **Master Source Preservation** | `python verify_sources.py` | 2/2 Excel masters verified byte-identical (`4dc79842...` & `5a1f9af7...`) | **PASSED** |
| **Disallowed Domain Rejection** | `pytest test_package_safety.py` | Zero disallowed domains across active registries/manifests | **PASSED** |
| **Full Regression Suite** | `pytest` | **94/94 tests passing** (100% pass rate) | **PASSED** |
| **Excel COM Recalculation** | `python work/recalculate_template_v9_models.py` | Exactly 0 formula errors across all 64 sheets | **PASSED** |
| **Balance Sheet Integrity** | `BS!Row 40 = Row 38 - Row 16` | Exactly `0.00 Cr` across all 48 company-periods (FY26E diff `+0.0000 Cr`) | **PASSED** |
| **Dynamic Sensitivities** | LTM & CLEAN perturbation / restoration | Live responsiveness with exact restoration to baseline | **PASSED** |
| **Audit Ledger Completeness** | `docs/TEMPLATE_V9_SHEET_AUDIT_LEDGER.md` | 64 entries (4 targets × 16 sheets) verified | **PASSED** |
| **Human Signoff Gate** | User review of Template V9 deliverables | `ready: false`, `human_pack_approval: pending` | **PENDING** |
