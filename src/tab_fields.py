from tableaudocumentapi.field import Field
from tab_util import TabUtil as tu
from col_fields import ColFields as cf

def outcsv():
    csvpath = 'Output/csv/fields.csv'
    csv_fields = open(csvpath, 'w', encoding='utf-8')
    print(cf.get_colstr(), file=csv_fields)

    for datasource in tu.twbx.datasources:
        if datasource.name == 'Parameters':
            continue
        for field in datasource.fields.values():
            f: Field = field
            if str(f.id).startswith(('[パラメーター', '[Parameter', '[Parámetro', '[Paramètre', '[Parâmetro', '[参数', '[參數', '[매개 변수')):
                continue
            if f.id == '[:Measure Names]':
                continue

            fldary = [''] * len(cf)
            fldary[cf.DataSourceName] = str(datasource.name) if str(datasource.caption) == '' else str(datasource.caption)
            fldary[cf.role] = 'D' if str(f.role) == 'dimension' else 'M'
            fldary[cf.type] = cf.convert_short_datatype(str(f.datatype))

            if f.caption is None:
                fcp = str(f.id).strip('[]')
            else:
                fcp = str(f.caption)
            fldary[cf.FieldName] = fcp

            if not f.calculation is None:
                dsid = '[' + datasource.name + ']'
                fldary[cf.Formula] = tu.replace_calc_field_id_to_caption(str(f.calculation), dsid)

            print(','.join(fldary), file=csv_fields)

    csv_fields.close()

if __name__ == '__main__':
    tu.initialize('10 hacks to make your dashboard GREAT!.twbx')
    outcsv()
    print('Done')