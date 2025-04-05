import openpyxl

workbook = openpyxl.Workbook()
del workbook['Sheet']
workbook.create_sheet('produtos')

sheet_produtos = workbook['produtos']

sheet_produtos.append(['computador','ano','preço'])

sheet_produtos.append(['computador 1','2001','500'])
sheet_produtos.append(['computador 2','2002','1000'])
sheet_produtos.append(['computador 3','2003','1500'])
sheet_produtos.append(['computador 4','2004','2500'])
sheet_produtos.append(['computador 5','2005','3000'])
workbook.save('computadores.xlsx')