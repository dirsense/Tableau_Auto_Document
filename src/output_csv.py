from tab_util import TabUtil as tu
import tab_fields, tab_filters, tab_actions, tab_datasources
import tab_displays, tab_controls, tab_parameters, tab_viewsheets, tab_zones
import os

def run(filepath: str):
    dir = 'Output/csv'
    for files in os.listdir(dir):
        path = os.path.join(dir, files)
        try:
            os.remove(path)
        except:
            pass

    tu.initialize(filepath)

    tab_fields.outcsv()
    tab_filters.outcsv()
    tab_actions.outcsv()
    tab_datasources.outcsv()
    tab_displays.outcsv()
    tab_controls.outcsv()
    tab_parameters.outcsv()
    tab_viewsheets.outcsv()
    tab_zones.outcsv()

def main():
    run('test.twbx')
    print('Done')

if __name__ == '__main__':
    main()