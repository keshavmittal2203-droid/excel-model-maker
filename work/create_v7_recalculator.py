with open('work/recalculate_template_v6_models.py', encoding='utf-8') as f:
    text = f.read()

# Replace V6 references with V7
text_v7 = text.replace('template_v6', 'template_v7')
text_v7 = text_v7.replace('Template V6', 'Template V7')

with open('work/recalculate_template_v7_models.py', 'w', encoding='utf-8') as f:
    f.write(text_v7)

print('Successfully created work/recalculate_template_v7_models.py')
