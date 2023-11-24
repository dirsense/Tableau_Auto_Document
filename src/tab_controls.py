from col_controls import ColControls as cc
from tab_util import TabUtil as tu
import xml.etree.ElementTree as ET
import re

def outcsv():
    def common_process(display_shname: str, e: ET.Element, dash: bool):
        param = e.attrib['param']
        emode = '' if not 'mode' in e.attrib else e.attrib['mode']
        etype = e.attrib['type' + ('-v2' if dash else '')]
        dcap = ''
        fcap = ''

        if 'param' in etype:
            dcap = 'Parameters'
            did = '[' + re.findall(tu.inner_square_pattern, param)[1] + ']'
            fcap = tu.param_dict[did]
        else:
            dsfs = tu.get_ds_fields_by_coron_str(param)[0]
            dcap = dsfs[0]
            fcap = dsfs[1]

        # Do not allow double registration (zone tags will be doubled if nothing is done because the <devicelayouts> tag has the exact same combination).
        opekey = display_shname + '-' + dcap + '-' + fcap
        if opekey in opedict:
            return
        opedict[opekey] = 0

        mode = ''
        modeshort = ''
        if 'show-null-ctrls' in e.attrib:
            # For major filters, it comes here. In the case of major, it is difficult to find a way to determine the filter type in detail, so we'll fix the "slider".
            mode = 'Slider'
            modeshort = 'Slider'
        elif emode == '' and not dash:
            # If worksheet display in dimension and mode is empty, it should be for multiple values (list) (maybe because it is the default value, no mode attribute is attached).
            mode = 'Multiple Values (list)'
            modeshort = '* List'
        elif emode == 'list':
            mode = 'Single Value (list)'
            modeshort = '1 List'
        elif emode == 'compact':
            mode = 'Single Value (dropdown)'
            modeshort = '1 DD'
        elif emode == 'slider':
            mode = 'Slider' if 'param' in etype else 'Single Value (slider)'
            modeshort = 'Slider'
        elif emode == 'type_in':
            mode = 'Input'
            modeshort = 'Input'
        elif emode == 'radiolist':
            mode = 'Single Value (list)'
            modeshort = '1 List'
        elif emode == 'dropdown':
            mode = 'Single Value (dropdown)'
            modeshort = '1 DD'
        elif emode == 'checklist':
            mode = 'Multiple Values (list)'
            modeshort = '* List'
        elif emode == 'checkdropdown':
            mode = 'Multiple Values (dropdown)'
            modeshort = '* DD'
        elif emode == 'typeinlist':
            mode = 'Multiple Values (custom list)'
            modeshort = '* List'
        elif emode == 'pattern':
            mode = 'Wildcard Match'
            modeshort = 'Input'
        
        src_shname = display_shname
        if dash and not 'param' in etype:
            src_shname = e.attrib['name']

        opeary = [''] * len(cc)
        opeary[cc.DisplaySheetName] = display_shname
        opeary[cc.SourceSheetName] = src_shname
        opeary[cc.DataSourceName] = str(dcap).strip('[]')
        opeary[cc.FieldName] = str(fcap).strip('[]')
        opeary[cc.ControlName] = mode
        opeary[cc.ControlShortName] = modeshort

        print(','.join(opeary), file=csvio)

    csvpath = 'Output/csv/controls.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cc.get_colstr(), file=csvio)
    opedict = {}

    for d in tu.root.iter('dashboard'):
        dashname = d.attrib['name']

        for z in d.iter('zone'):
            if 'type-v2' in z.attrib and (z.attrib['type-v2'] == 'filter' or z.attrib['type-v2'] == 'paramctrl'):
                common_process(dashname, z, True)
    
    for w in tu.root.iter('window'):
        if 'class' in w.attrib and w.attrib['class'] == 'worksheet':
            wsname = w.attrib['name']

            for c in w.iter('card'):
                if 'type' in c.attrib and (c.attrib['type'] == 'filter' or c.attrib['type'] == 'parameter'):
                    common_process(wsname, c, False)

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')


