from displays_type import DisplaysType
from col_displays import ColDisplays as cd
from tab_util import TabUtil as tu
from col_fields import ColFields as cf

def outcsv():
    def input_dlist(shname: str, youto: str, dsname: str, fdname: str, direct_formula: str, role: str, dataType: str):
        dlist = [''] * len(cd)
        dlist[cd.WorkSheetName] = shname
        dlist[cd.use] = youto
        dlist[cd.DataSourceName] = dsname.strip('[]')
        if fdname[len(fdname)-1:] == ']' and fdname[:1] == '[':
            dlist[cd.FieldName] = fdname.strip('[]')
        else:
            dlist[cd.FieldName] = fdname

        if role != '':
            dlist[cd.role] = 'D' if role == 'dimension' else 'M'
        if dataType != '':
            dlist[cd.type] = cf.convert_short_datatype(dataType)
        if direct_formula != '':
            dlist[cd.FieldName] = '"' + fdname.replace('"', '""') + '"'
        if '//' in direct_formula[:4]:
            dlist[cd.AdHocCalculatedField] = direct_formula

        print(','.join(dlist), file=csvio)

    csvpath = 'Output/csv/displays.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cd.get_colstr(), file=csvio)

    for w in tu.root.iter('worksheet'):
        shname = w.attrib['name']

        for dtype in DisplaysType:
            dcap_fcap_list = []
            tag = DisplaysType.get_tagstr(dtype)

            if dtype == DisplaysType.Rows or dtype == DisplaysType.Cols:
                for m in w.iter(tag):
                    if m.text is not None:
                        dcap_fcap_list = tu.get_display_fields_by_coron_str(m.text)
                    break
            else:
                for e in w.iter('encodings'):
                    for m in e.iter(tag):
                        dcap_fcap_list.extend(tu.get_display_fields_by_coron_str(str(m.attrib['column'])))
            
            for dfl in dcap_fcap_list:
                input_dlist(shname, DisplaysType.get_capstr(dtype), dfl[0], dfl[1], dfl[4], dfl[5], dfl[6])

        for gf in w.iter('groupfilter'):
            if 'level' in gf.attrib and gf.attrib['level'] == '[:Measure Names]':
                if 'member' in gf.attrib:
                    dfl = tu.get_display_fields_by_coron_str(str(gf.attrib['member']))[0]
                    input_dlist(shname, 'MV', dfl[0], dfl[1], '', '', '')

    csvio.close()

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')