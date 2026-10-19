from pathlib import Path

pr3_path = Path("work/pr3_body.md")
text = pr3_path.read_text(encoding="utf-8")

target_str = "### 16. Source Evidence & Valuation Alignment: Evidenced 6.51% Risk-Free Rate Proxy, Assumption Classification & PBT Valuation Trace"
idx = text.find(target_str)
assert idx != -1, "Target not found in pr3_body.md"

replacement = """### 16. Source Evidence & Valuation Alignment: Evidenced 6.51% Risk-Free Rate Proxy, Assumption Classification & Fact Boundary

- **Risk-Free Rate Evidenced Proxy**:
  - Cell `'WACC-DCF'!J22`: Updated to **6.51%** (0.0651), sourced from allowlisted provider **Trading Economics** (`https://tradingeconomics.com/india/government-bond-yield`), supported by its 14-Oct-2025 published reading reporting the 10-year benchmark yield just below 6.51% as the dated prior-session proxy preceding the 15-Oct-2025 cutoff.
- **Strict Classification of Modeling Assumptions**:
  - **Equity Risk Premium (ERP, `'WACC-DCF'!J23` = 5.50%)**: Classified as **Analyst Macroeconomic Assumption**; not attributed to Trading Economics.
  - **Marginal Corporate Tax Rate (`'WACC-DCF'!E23` and `'Data Sheet'!K29` = 25.17%)**: Classified as **Analyst Modeling Assumption** (22% base + 10% surcharge + 4% cess).
- **Fact vs Assumption Governance**:
  - Company-specific factual inputs absent from allowlisted sources cannot be relabeled as analyst assumptions.
  - **MM & EDELWEISS**: Valuation marked **Unavailable (`N/A`)** because required factual segment/subsidiary inputs are absent from allowed sources.
  - **LTM & CLEAN**: Unlevered FCFF DCF valuations are fully supported by allowed sources (LTM: **₹2,620.24 / share**; CLEAN: **₹393.03 / share**).

### 17. Final Template V16 Accounting Closeout & Deliverables Handoff

- **Final Deliverables in `companies/<TARGET>/template_v16/` & `deliverables/V16_user_handoff/`**:
  - `MM`: `companies/MM/template_v16/MM-FullTemplate.xlsx` (112,365 bytes, SHA-256: `2b5536e08a17d4b0f2a53afe4c39a84345a6a179ce6c7bd75a636e6bc452daf2`) — Forecast: **BLOCKED** (+₹7,677.49 Cr exposed gap); Valuation: **Unavailable (`N/A`)**.
  - `LTM`: `companies/LTM/template_v16/LTM-FullTemplate.xlsx` (114,500 bytes, SHA-256: `edbd6ba745a6a25de7f5308d3ef68d6d39f6440ada2e524244319895fd8cb769`) — Forecast: **UNRECONCILED** (+₹1.00 Cr source integer rounding delta); Valuation: **₹2,620.24 / share** (Supported).
  - `CLEAN`: `companies/CLEAN/template_v16/CLEAN-FullTemplate.xlsx` (112,841 bytes, SHA-256: `9bc682eed4885eb7fe32feb5bdc06f6fb5d8f096351d57dd1a76d5b4b08f4616`) — Forecast: **UNRECONCILED** (−₹1.13 Cr source rounding delta); Valuation: **₹393.03 / share** (Supported).
  - `EDELWEISS`: `companies/EDELWEISS/template_v16/EDELWEISS-FullTemplate.xlsx` (111,803 bytes, SHA-256: `899b1e9e3ca9c39c84b75364111736917e1fc72e15ea80598f494b6ea68ed2ed`) — Forecast: **UNRECONCILED** (+₹1.00 Cr source integer rounding delta); Valuation: **Unavailable (`N/A`)**.
- **Consolidated User Handoff Package**:
  - Handoff directory: `deliverables/V16_user_handoff/`
  - Handoff ZIP: `deliverables/V16_user_handoff.zip` (9,682,773 bytes; 67 files: 4 workbooks, 60 landscape PDFs, README_FIRST.md, USER_APPROVAL_PACKET.md, TEMPLATE_V16_SHEET_AUDIT_LEDGER.md).
- **Review Gate & Governance**:
  - Gate status: `ready: false`, pending gate: `["human_pack_approval"]`.
  - PR #3 remains open and unmerged pending human review.
"""

new_text = text[:idx] + replacement
pr3_path.write_text(new_text, encoding="utf-8")
print("Updated work/pr3_body.md successfully.")
