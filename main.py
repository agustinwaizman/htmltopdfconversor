import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import pdfkit
import jinja2
import os
import datetime

info = {}

def crea_pdf(info):
    template_loader = jinja2.FileSystemLoader('./')
    template_env = jinja2.Environment(loader=template_loader)

    html_template = 'temp.html'
    template = template_env.get_template(html_template)
    output_text = template.render(info)

    config = pdfkit.configuration(wkhtmltopdf='/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')
    output_pdf = f"comprobante_de_pago_{info['piso_dpto']}.pdf"
    pdfkit.from_string(output_text, output_pdf, configuration=config, css='style.css')


def excel_to_pdf(excel_file):

    if excel_file:
        wb = openpyxl.load_workbook(excel_file)
        sheets = wb.sheetnames
        for sheet in sheets:
            current_sheet = wb.get_sheet_by_name(sheet)
            fecha = current_sheet['A2'].value
            nro_de_pago = current_sheet['B2'].value
            locatario = current_sheet['C2'].value
            cond_pago = current_sheet['D2'].value
            concepto = current_sheet['E2'].value
            edif = current_sheet['F2'].value
            piso_dpto = current_sheet['G2'].value
            tot_a_pagar = current_sheet['H2'].value
            mora = current_sheet['I2'].value
            pagado = current_sheet['J2'].value
            info = {
                'fecha': fecha,
                'nro_de_pago': str(nro_de_pago),
                'locatario':str(locatario),
                'cond_pago':str(cond_pago),
                'concepto':str(concepto),
                'edif':str(edif),
                'piso_dpto':str(piso_dpto),
                'tot_a_pagar':str(tot_a_pagar),
                'mora':str(mora),
                'pagado':str(pagado),
            }
            crea_pdf(info)


def select_excel():
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    label1.config(text=os.path.basename(excel_file))
    return excel_file

def convert():
    excel_to_pdf(excel_file)
    messagebox.showinfo("Éxito", "Edición del PDF completada con éxito.")

root = tk.Tk()
root.title("Conversor de Archivos")

label1 = tk.Label(root, text="No se ha seleccionado ningún archivo Excel")
label1.pack()

button1 = tk.Button(root, text="Seleccionar archivo Excel", command=select_excel)
button1.pack()

convert_button = tk.Button(root, text="Convertir", command=convert)
convert_button.pack()

root.mainloop()