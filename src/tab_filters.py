from col_filters import ColFilters as cf
from tab_util import TabUtil as tu
from usrkey import UsrKey as uk
import itertools, re
import xml.etree.ElementTree as ET

def outcsv():
    csvpath = 'Output/csv/filters.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cf.get_colstr(), file=csvio)

    # Create a dictionary of worksheets to apply to.
    fil_tgtsh_dict = {}
    for w in tu.root.iter('worksheet'):
        for f in w.iter('filter'):
            if 'filter-group' in f.attrib and f.attrib['filter-group'] == '3':
                fg3key = f.attrib['class'] + '/' + f.attrib['column']
                fil_tgtsh_dict.setdefault(fg3key, []).append(w.attrib['name'])

    # The worksheets to which the filter is applied should be added to this set (to avoid duplicate registration).
    fg3key_added = set()
    one_sheets = []

    for t in itertools.chain(tu.root.iter('datasource'), tu.root.iter('shared-view'), tu.root.iter('worksheet')):
        shname = t.attrib['name']

        for f in t.iter('filter'):
            filary, fg3hit = input_filary(t, f, shname, fil_tgtsh_dict, fg3key_added)

            if len(filary) == len(cf):
                if fg3hit is None and t.tag == 'worksheet':
                    one_sheets.append(filary)
                else:
                    print(','.join(filary), file=csvio)

    for o in one_sheets:
        print(','.join(o), file=csvio)

    csvio.close()

def input_filary(t: ET.Element, f: ET.Element, shname: str, fil_tgtsh_dict: dict, fg3key_added: set):
    fcol = f.attrib['column']
    if t.tag == 'datasource':
        fcol = '[' + t.attrib['name'] + '].' + fcol
    dsname, fieldname, attrL, attrR = tu.get_ds_fields_by_coron_str(fcol)[0]
    dsname = str(dsname).strip('[]')

    # Actions are output elsewhere, so they are not listed in the Filter's CSV.
    if attrL[:6] == 'Action' or fieldname[:7] == '[Action':
        return [], None

    fg3hit = None
    filary = [''] * len(cf)
    filary[cf.DataSourceName] = dsname
    filary[cf.FieldName] = str(fieldname).strip('[]')
    filary[cf.ElementL] = attrL
    filary[cf.ElementR] = attrR

    if t.tag == 'datasource':
        filary[cf.ApplyTo] = 'DS (' + str(tu.ds_dict['[' + t.attrib['name'] + ']']).strip('[]') + ')'

    elif t.tag == 'shared-view':
        if f.find('target') is None:
            filary[cf.ApplyTo] = 'All Using This Data Source (' + dsname + ')'
        else:
            filary[cf.ApplyTo] = 'All Using Related Data Sources (' + dsname + ')'
    else:
        fg3key = f.attrib['class'] + '/' + f.attrib['column']
        if fg3key in fg3key_added:
            return [], None
        if fg3key in fil_tgtsh_dict and shname in fil_tgtsh_dict[fg3key]:
            filary[cf.ApplyTo] = ';'.join(fil_tgtsh_dict[fg3key])
            fg3key_added.add(fg3key)
            fg3hit = True
        else:
            filary[cf.ApplyTo] = shname

    if f.attrib['class'] == 'quantitative':
        filary[cf.Type] = 'M'
        incval = f.attrib['included-values']
        zenpan = ''
        if 'range' in incval:
            if f.find('min') is None:
                zenpan = 'At most'
            elif f.find('max') is None:
                zenpan = 'At least'
            else:
                zenpan = 'Range'
            
            if 'or-null' in incval:
                filary[cf.OR_NULL] = 'TRUE'
        else:
            zenpan = incval.replace('non-', '≠')

        filary[cf.General] = zenpan

    else:
        filary[cf.Type] = 'C' if 'context' in f.attrib else 'D'

        ## General ~ Top
        zenpan = 'List'
        zenpan_mem = []
        zenpan_jogailist = ''
        zenpan_jogai = ''

        wild_text = ''
        wild_match = ''
        wild_jogai = ''

        jouken = ''
        joui = ''
        joui_count = ''
        joui_kai = ''

        for gf in f.iter('groupfilter'):
            gfunc = gf.attrib['function']
            uienum = ''
            if uk.UiEnum.value in gf.attrib:
                uienum = gf.attrib[uk.UiEnum.value]

                if uk.UiMarker.value in gf.attrib and gf.attrib[uk.UiMarker.value] == 'enumerate':
                    if uienum == 'all':
                        zenpan = 'All'
                    elif uienum == 'exclusive':
                        zenpan_jogai = 'TRUE'

                    if gfunc == 'except':
                        zenpan_jogailist = 'TRUE'

            if 'member' in gf.attrib:
                if uk.UiCustom.value in gf.attrib:
                    zenpan = 'Custom'
                if gfunc == 'empty-level':
                    pass
                else:
                    mem = str(gf.attrib['member']).strip('"')
                    if mem == '%null%':
                        mem = 'NULL'
                    elif mem == '':
                        mem = '(Empty string)'
                    elif mem[0] == '[':
                        _, fieldname, _, _ = tu.get_ds_fields_by_coron_str(mem)[0]
                        mem = fieldname
                    zenpan_mem.append(mem)

            elif uk.UiPatText.value in gf.attrib:
                wild_text = gf.attrib[uk.UiPatText.value]
                uipattype = gf.attrib[uk.UiPatType.value]
                if uipattype == 'contains':
                    wild_match = 'Contains'
                elif uipattype == 'starts-with':
                    wild_match = 'Starts with'
                elif uipattype == 'ends-with':
                    wild_match = 'Ends with'
                elif uipattype == 'exact-match':
                    wild_match = 'Exactly matches'
                wild_jogai = 'TRUE' if uienum == 'exclusive' else ''

            elif uk.UiMarker.value in gf.attrib and gf.attrib[uk.UiMarker.value] == 'filter-by' and 'expression' in gf.attrib:
                jouken = gf.attrib['expression']
                rfall = re.findall('\[.+?\]', fcol)
                if len(rfall) > 0:
                    jouken = tu.replace_calc_field_id_to_caption(jouken, rfall[0])
            
            elif gfunc == 'order' and 'expression' in gf.attrib:
                joui = gf.attrib['expression']
                rfall = re.findall('\[.+?\]', fcol)
                if len(rfall) > 0:
                    joui = tu.replace_calc_field_id_to_caption(joui, rfall[0])

            elif 'count' in gf.attrib:
                joui_count = gf.attrib['count']
                joui_kai = 'Top' if gf.attrib['end'] == 'top' else 'Bottom'

        filary[cf.General] = zenpan
        filary[cf.General_List] = ';'.join(zenpan_mem)
        filary[cf.General_ExcludeList] = zenpan_jogailist
        filary[cf.General_ExcludeCheck] = zenpan_jogai

        filary[cf.Wildcard] = wild_text
        filary[cf.Wildcard_Match] = wild_match
        filary[cf.Wildcard_Exclude] = wild_jogai

        filary[cf.Condition] = jouken

        filary[cf.Top] = joui
        filary[cf.Top_Count] = joui_count
        filary[cf.Top_Bottom] = joui_kai

    if t.tag == 'datasource':
        filary[cf.Type] = 'DS'

    return filary, fg3hit

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')