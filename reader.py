import pandas as pd
import os
import win32com.client as win32

invoice_form = pd.read_excel('Forma.xlsx', header=None)
schools = pd.read_excel('Sąrašas.xlsx', sheet_name=None, header=None)
month_numeric = '04'
month_text = 'balandžio mėn.'
invoice_No = 1

for school_name, students in schools.items():
    for i in range(2, len(students)):
        full_student_name = students.iloc[i, 1] + ' ' + students.iloc[i, 2]
        Full_parent_name = students.iloc[i, 6]
        price_numeric = students.iloc[0, 3]
        price_text = students.iloc[0, 4]

        invoice = invoice_form.copy()

        invoice.iloc[1, 1] = 'Serija ILO Nr. 2025' + month_numeric + '01-' + str(invoice_No)
        invoice.iloc[2, 1] = '2025-' + month_numeric + '-01'
        invoice.iloc[7, 2] = Full_parent_name
        invoice.iloc[15, 2] = str(price_numeric) + '.00'
        invoice.iloc[16, 1] = 'edukaciniai užsiėmimai, už ' + month_text
        invoice.iloc[17, 1] = full_student_name
        invoice.iloc[19, 0] = 'Pastabos: Mokėjimo paskirtyje įrašyti: ' + full_student_name
        invoice.iloc[20, 0] = 'Suma žodžiais: ' + price_text + ' eurų 00 ct'
        invoice.iloc[21, 0] = 'Sąskaitą apmokėti iki 2025.' + month_numeric + '.15'

        temp_excel = os.path.join('sąskaitos_exl', f"{full_student_name}.xlsx")
        invoice.to_excel(temp_excel, index=False, header=False)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.abspath(temp_excel))
        ws = wb.Sheets(1)
        ws.Columns("A").ColumnWidth = 10
        ws.Columns("B").ColumnWidth = 40

        xlEdgeLeft = 7
        xlEdgeTop = 8
        xlEdgeBottom = 9
        xlEdgeRight = 10
        xlInsideHorizontal = 12
        xlLineStyleNone = -4142

        for col in ['A', 'B', 'C']:
            cell = ws.Range(f"{col}15")
            for border_id in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]:
                border = cell.Borders(border_id)
                border.LineStyle = 1
                border.Weight = 2

        for col in ['A', 'B', 'C']:
            cell_range = ws.Range(f"{col}16:{col}18")
            for border_id in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]:
                border = cell_range.Borders(border_id)
                border.LineStyle = 1
                border.Weight = 2
            cell_range.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

        pdf_path = os.path.join('sąskaitos_pdf', f"{full_student_name}.pdf")
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        wb.Close(False)
        excel.Quit()

        invoice_No += 1
