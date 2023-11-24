from col_parameters import ColParameters as cp
from col_fields import ColFields
from tab_util import TabUtil as tu

def outcsv():
    def get_domain_type_jpn(domtype: str):
        if domtype == 'any':
            return 'All'
        elif domtype == 'list':
            return 'List'
        elif domtype == 'range':
            return 'Range'
        else:
            return domtype

    csvpath = 'Output/csv/parameters.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cp.get_colstr(), file=csvio)

    for ds in tu.root.iter('datasource'):
        if ds.attrib['name'] == 'Parameters':
            for col in ds.iter('column'):
                pary = [''] * len(cp)
                pary[cp.ParameterName] = col.attrib['caption']
                pary[cp.Type] = ColFields.convert_short_datatype(col.attrib['datatype'])
                pary[cp.AllowableValues] = get_domain_type_jpn(col.attrib['param-domain-type'])
                if '_.fcp.ParameterDefaultValues.true...default-value-field' in col.attrib:
                    defvalue = col.attrib['_.fcp.ParameterDefaultValues.true...default-value-field']
                    caps = tu.get_ds_fields_by_coron_str(defvalue)[0]
                    deffield = str(caps[1]).strip('[]') + ' (' + str(caps[0]).strip('[]') + ')'
                    pary[cp.WhenWorkbookOpens] = deffield
                else:
                    pary[cp.WhenWorkbookOpens] = 'Current value'
                print(','.join(pary), file=csvio)
            break

    csvio.close()

if __name__ == '__main__':
    tu.initialize('test.twb')
    outcsv()
    print('Done')