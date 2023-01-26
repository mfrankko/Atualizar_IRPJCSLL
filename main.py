import openpyxl as xl
import tkinter as tk
from tkinter import ttk

def atualizar_apuracao(empresa, periodo):
    balancetes = 'J:\\Tributario e fiscal\\Impostos Diretos\\Apurações IRPJ-CSLL\\Balancetes\\Balancete Macro - ' + str(
        periodo) + '.xlsm'
    empresa_apuracao = 'J:\\Tributario e fiscal\\Impostos Diretos\\Apurações IRPJ-CSLL\\' + str(empresa) + '.xlsx'
    wb = xl.load_workbook(filename=balancetes, data_only=True)
    ws1 = wb[empresa]
    wb2 = xl.load_workbook(filename=empresa_apuracao)
    ws2 = wb2[str(periodo)]
    for row in ws1:
        for cell in row:
            ws2[cell.coordinate].value = cell.value
    ws2.delete_rows(1,20)
    ws2.delete_cols(1)

    for cell in ws2[('A')]:
        cell.number_format = '0'
    wb2.save(empresa_apuracao)
    print('Apuração da ' + str(empresa) + ' atualizada com sucesso!')

root = tk.Tk()
root.title("GUI Form")

def get_input():
    global company, period
    company = company_var.get()
    period = period_var.get()
    return company, period

company_var = tk.StringVar()
period_var = tk.StringVar()

company_label = ttk.Label(root, text="Empresa:")
company_label.pack()

company_dropdown = ttk.Combobox(root, textvariable=company_var)
company_dropdown['values'] = ["BR01", "BR02", "BR03", "BR05", "BR06", "BR07", "BR08", "BR09", "BR10", "BR11"]
company_dropdown.pack()

period_label = ttk.Label(root, text="Período:")
period_label.pack()

period_dropdown = ttk.Combobox(root, textvariable=period_var)
period_dropdown['values'] = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
period_dropdown.pack()

submit_button = ttk.Button(root, text="Submit", command=lambda: [get_input(), root.quit()])
submit_button.pack()

root.mainloop()

atualizar_apuracao(company, period)