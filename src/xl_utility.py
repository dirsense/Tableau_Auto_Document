import xlwings as xw
import PySimpleGUI as sg

def run_excel_macro(xlsmpath: str, macroname: str):
    try:
        excel = xw.App(visible=False)
        wb = excel.books.open(xlsmpath)

        macro = wb.macro(macroname)
        macro()

        wb.save()
        wb.close()

    except Exception as e:
        sg.popup('An error occurred in the Excel macro execution process.\n' + e.args[0], title='Error')

    finally:
        excel.quit()

def file_is_open(filepath: str):
    try:
        f = open(filepath, 'a')
        f.close()
    except:
        return True
    else:
        return False