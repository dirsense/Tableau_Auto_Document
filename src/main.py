import PySimpleGUI as sg
import xl_utility as xu
import output_csv
import glob, os, shutil

sg.theme('DarkGrey8')

def generate_document(values):
    op1 = values['-genop1-']
    twbpath = values['-twbpath-']
    twbname = str(os.path.splitext(os.path.basename(twbpath))[0]).replace("'", "")
    finalpath = 'Output/TAD_' + twbname + '.xlsm'
    tmppath = 'Template/TAD_Template.xlsm'
    tadexist = os.path.isfile(finalpath)

    if not os.path.isfile(twbpath):
        sg.popup('The specified Tableau Workbook file does not exist\n' + twbpath, title='Validation Error')
        return
    if tadexist and xu.file_is_open(finalpath):
        sg.popup('A TAD file with the name to be saved is opened, close it and try again\n' + finalpath, title='Validation Error')
        return
    if op1 and not os.path.isfile(tmppath):
        sg.popup('The original template file does not exist\n' + tmppath, title='Validation Error')
        return

    try:
        output_csv.run(twbpath)
    except Exception as e:
        sg.popup('An error occurred during the CSV output process\n' + e.args[0], title='Error')
        return

    if op1:
        shutil.copyfile(tmppath, finalpath)
        xu.run_excel_macro(finalpath, 'CSV_NewPostingProcess')
        sg.popup('Data Update & Create Book is Complete!\n' + finalpath, title='Complete')
    else:
        xu.run_excel_macro(finalpath, 'CSV_UpdatePostingProcess')
        sg.popup('Data Update is Complete!\n' + finalpath, title='Complete')

def get_current_paths() -> tuple[str, str]:
    cur_paths = glob.glob('*.twb') + glob.glob('*.twbx')
    cur_value = '' if len(cur_paths) == 0 else cur_paths[0]
    return cur_value, cur_paths

def main():
    cur_value, cur_paths = get_current_paths()

    layout = [
        [sg.Column(vertical_alignment='top', layout=[
            [sg.Text('TWB or TWBX Path')],
            [sg.Button('Generate', pad=((5, 5), (5, 3)))],
        ], pad=(0, 0)),
        sg.Column(vertical_alignment='top', layout=[
            [sg.Combo(cur_paths, default_value=cur_value, key='-twbpath-', size=(50, 1))],
            [sg.Button('Reload Current Path', pad=((5, 5), (5, 3))),
                sg.FileBrowse('...', target='-twbpath-', file_types=(('Tableau Workbook', '*.twb*'), ),
                initial_folder='.', pad=((5, 5), (5, 3)))]
        ], pad=(0, 0)),
        ],
        [sg.Column(vertical_alignment='top', layout=[
            [sg.Radio('Data Update & Create Book (Recommend)', key='-genop1-', group_id='genop', default=True, pad=((1, 0), (3, 0)))],
            [sg.Radio('Data Update Only', key='-genop2-', group_id='genop', pad=((1, 0), (0, 0)))]
        ], pad=(0, 0))]
    ]

    window = sg.Window('Tableau Auto Document', layout)

    while True:
        event, values = window.read()

        if event is None:
            break

        elif event == 'Generate':
            generate_document(values)
        
        elif event == 'Reload Current Path':
            cur_value, cur_paths = get_current_paths()
            window['-twbpath-'].update(value=cur_value, values=cur_paths)

    window.close()

if __name__ == '__main__':
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass
    main()
