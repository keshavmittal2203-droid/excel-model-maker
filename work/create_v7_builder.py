with open('work/build_template_v6_models.py', encoding='utf-8') as f:
    text = f.read()

# Replace V6 references with V7
text_v7 = text.replace('template_v6', 'template_v7')
text_v7 = text_v7.replace('Template V6', 'Template V7')
text_v7 = text_v7.replace('build_template_v6', 'build_template_v7')

# Add explicit assumption disclosures to MM and EDELWEISS on Model -> cover
old_cov = '''    cov["C16"].value = f"Valuation Basis: {cfg['primary_method_name']}. Approved sources: Screener, Yahoo Finance, NSE, Trading Economics."'''
new_cov = '''    val_status_tag = "Assumption-Dependent (inputs unavailable on Screener.in)" if target_key in ("MM", "EDELWEISS") else "Fully Supported from Allowlisted Sources"
    cov["C16"].value = f"Valuation Basis: {cfg['primary_method_name']}. Status: {val_status_tag}. Approved sources: Screener, Yahoo Finance, NSE, Trading Economics."'''
assert old_cov in text_v7, 'Could not find cov[C16] in text'
text_v7 = text_v7.replace(old_cov, new_cov)

# In DCF 1st way, add assumption comments to E42 for MM
old_mm_e42 = '''        dcf1["E42"].value = "=G60" # Single source of truth'''
new_mm_e42 = '''        dcf1["E42"].value = "=G60" # Single source of truth
        add_comment(dcf1["E42"], "SOTP valuation is assumption-dependent: segment Core EBITDA (Auto Rs 11,700 Cr, Farm Rs 6,716 Cr) and standalone net cash (+Rs 9,109 Cr) are unverified filing inputs / analyst modeling assumptions not published on Screener.in.")'''
assert old_mm_e42 in text_v7, 'Could not find MM E42 in text'
text_v7 = text_v7.replace(old_mm_e42, new_mm_e42)

# In DCF 1st way, add assumption comments to E42 for EDELWEISS
old_ed_e42 = '''        dcf1["E42"].value = "=G60"\n'''
new_ed_e42 = '''        dcf1["E42"].value = "=G60"
        add_comment(dcf1["E42"], "SOTP valuation is assumption-dependent: vertical financial metrics (Alternatives PAT Rs 375 Cr, EARCL BV Rs 1,960 Cr, Mutual Fund AUM Rs 1,50,000 Cr, Insurance IEV Rs 2,100 Cr) and standalone debt (-Rs 3,067 Cr) are unverified filing inputs / analyst modeling assumptions not published on Screener.in.")\n'''
assert old_ed_e42 in text_v7, 'Could not find EDELWEISS E42 in text'
text_v7 = text_v7.replace(old_ed_e42, new_ed_e42)

# In BS sheet, add comment documenting Screener popup schedules vs unverified sub-schedules
old_bs = '''    bs_ws = wb["BS"]'''
new_bs = '''    bs_ws = wb["BS"]
    if target_key == "MM":
        add_comment(bs_ws["A1"], "Mahindra & Mahindra Consolidated Balance Sheet: Top-level balance sheet lines are verified from Screener.in. Consolidated sub-schedules (Trade Payables, NCI, Other Liabilities, Cash, Receivables, Loans) are unverified filing inputs / analyst modeling assumptions not published on Screener.in.")
    elif target_key in ("LTM", "EDELWEISS"):
        add_comment(bs_ws["A1"], f"{cfg['company_name']} Consolidated Balance Sheet: All balance sheet lines and sub-schedules (Borrowings, Other Liabilities, Other Assets) are verified directly against Screener.in popup schedules.")
    elif target_key == "CLEAN":
        add_comment(bs_ws["A1"], "Clean Science Consolidated Balance Sheet: FY18-FY25 lines and sub-schedules verified on Screener.in consolidated popup schedules. Pre-FY18 standalone figures verified on Screener.in standalone page.")'''
assert old_bs in text_v7, 'Could not find bs_ws = wb["BS"] in text'
text_v7 = text_v7.replace(old_bs, new_bs)

with open('work/build_template_v7_models.py', 'w', encoding='utf-8') as f:
    f.write(text_v7)

print('Successfully created work/build_template_v7_models.py')
