import openpyxl

path = 'C:/Users/kesha/Documents/Codex/2026-10-18/excel-model-maker-v9-remediation/companies/LTM/template_v15_excel_com/LTM-FullTemplate-v15-ExcelCOM.xlsx'
wb = openpyxl.load_workbook(path, data_only=True)
bs = wb['BS']
cf = wb['Cash Flow']
ds = wb['Data Sheet']

# Let's compute Assets(N) - Assets(M):
# Assets: N28 - N29 + N30 + N31 + N33 + N34 + N35 + N36 + N37
# Liabilities: N4 + N5 + N7 + N8 + N10 + N12 + N13 + N14 + N15

assets_M = (bs['M28'].value - bs['M29'].value + bs['M30'].value + bs['M31'].value + 
            bs['M33'].value + bs['M34'].value + bs['M35'].value + bs['M36'].value + bs['M37'].value)
liab_M = (bs['M4'].value + bs['M5'].value + bs['M7'].value + bs['M8'].value + bs['M10'].value + 
          bs['M12'].value + bs['M13'].value + (bs['M14'].value or 0) + bs['M15'].value)
print(f"FY25 Assets (M): {assets_M}, Liab (M): {liab_M}, Diff M: {assets_M - liab_M}")

# Now let's calculate the formula values of N:
N28 = bs['N28'].value
N29 = bs['N29'].value
N30 = bs['N30'].value
N31 = bs['N31'].value
N33 = bs['N33'].value
N34 = bs['N34'].value
N35 = bs['N35'].value
N36 = bs['N36'].value
N37 = bs['N37'].value

assets_N = N28 - N29 + N30 + N31 + N33 + N34 + N35 + N36 + N37
reported_assets_N = bs['N38'].value
print(f"Assets N calculated: {assets_N}, Reported N38: {reported_assets_N}")

N4 = bs['N4'].value
N5 = bs['N5'].value
N7 = bs['N7'].value
N8 = bs['N8'].value
N10 = bs['N10'].value
N12 = bs['N12'].value
N13 = bs['N13'].value
N14 = bs['N14'].value or 0
N15 = bs['N15'].value

liab_N = N4 + N5 + N7 + N8 + N10 + N12 + N13 + N14 + N15
reported_liab_N = bs['N16'].value
print(f"Liab N calculated: {liab_N}, Reported N16: {reported_liab_N}")

print(f"Diff N: {assets_N - liab_N}")

# Now let's check N35 (Cash):
print(f"M35: {bs['M35'].value}")
print(f"N35: {bs['N35'].value}")
print(f"CF!N31: {cf['N31'].value}")
print(f"CF CFO (N5): {cf['N5'].value}")
print(f"CF CFI (N14): {cf['N14'].value}")
print(f"CF CFF (N23): {cf['N23'].value}")
print(f"True CF Net (N5+N14+N23): {cf['N5'].value + cf['N14'].value + cf['N23'].value}")

# If N35 was M35 + (N5 + N14 + N23):
true_N35 = bs['M35'].value + (cf['N5'].value + cf['N14'].value + cf['N23'].value)
print(f"If N35 = M35 + true_CF_net: {true_N35}")
true_assets_N = N28 - N29 + N30 + N31 + N33 + N34 + true_N35 + N36 + N37
print(f"True Assets N - Liab N: {true_assets_N - liab_N}")
