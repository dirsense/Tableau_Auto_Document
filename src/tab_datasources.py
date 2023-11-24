from tab_util import TabUtil as tu
from col_datasources import ColDatasources as cd

def outcsv():
    def get_rsdict():
        g_rsdict = {}
        for rs in tu.root.iter('datasource-relationship'):
            srcid = '[' + rs.attrib['source'] + ']'
            dstid = '[' + rs.attrib['target'] + ']'
            srccp = str(tu.ds_dict[srcid]).strip('[]')
            dstcp = str(tu.ds_dict[dstid]).strip('[]')
            klist = []
            vlist = []
            for mk in rs.iter('map'):
                k = str(tu.get_ds_fields_by_coron_str(mk.attrib['key'])[0][1]).strip('[]')
                v = str(tu.get_ds_fields_by_coron_str(mk.attrib['value'])[0][1]).strip('[]')
                klist.append(k)
                vlist.append(v)
            if len(klist) > 0:
                g_rsdict[srccp + '.' + dstcp] = (klist, vlist)

        return g_rsdict

    def get_mapset(ds1: tuple[str, str], ds2: tuple[str, str]) -> set:
        def in_xcap(xcaps: list) -> bool:
            xflg = False
            for xcap in xcaps:
                if (xcap[0] == ('[' + ds1[1] + ']') and xcap[1] == ccap) or (xcap[0] == ('[' + ds2[1] + ']') and xcap[1] == ccap):
                    xflg = True
                    break
            return xflg

        gmapset = set()
        for dd in w.iter('datasource-dependencies'):
            if ds1[0] == dd.attrib['datasource']:
                instance_ids = set()
                for c in dd.iter('column-instance'):
                    instance_ids.add(c.attrib['column'])

                for c in dd.iter('column'):
                    if c.attrib['type'] != 'quantitative':
                        if 'caption' in c.attrib:
                            ccap = '[' + c.attrib['caption'] + ']'
                        else:
                            ccap = c.attrib['name']
                        if ccap in tu.field_dict['[' + ds2[0] + ']'].values():
                            if (c.attrib['name'] in instance_ids or in_xcap(ixcaps)) and not in_xcap(excaps):
                                gmapset.add(ccap.strip('[]'))
                break
        return gmapset

    csvpath = 'Output/csv/datasources.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cd.get_colstr(), file=csvio)

    rs_dict = get_rsdict()

    for w in tu.root.iter('worksheet'):
        dsary = [''] * len(cd)
        primary = ()
        secondaries = []
        ixcaps = set()
        excaps = set()
        wsname = w.attrib['name']

        if not w.find('table') is None and not w.find('table').find('join-lod-include-overrides') is None:
            for ix in w.find('table').find('join-lod-include-overrides').findall('column'):
                ixlist = tu.get_display_fields_by_coron_str(ix.text)[0]
                if len(ixlist) > 0:
                    ixcaps.add((ixlist[0], ixlist[1]))

        if not w.find('table') is None and not w.find('table').find('join-lod-exclude-overrides') is None:
            for ex in w.find('table').find('join-lod-exclude-overrides').findall('column'):
                exlist = tu.get_display_fields_by_coron_str(ex.text)[0]
                if len(exlist) > 0:
                    excaps.add((exlist[0], exlist[1]))

        for i, ds in enumerate(w.iter('datasource')):
            dscap = ds.attrib['name']
            if 'caption' in ds.attrib:
                dscap = ds.attrib['caption']
            if ds.attrib['name'] == 'Parameters':
                continue
            if i == 0:
                primary = (ds.attrib['name'], dscap)
            else:
                secondaries.append((ds.attrib['name'], dscap))

        if len(primary) == 0:
            continue
        if len(secondaries) == 0:
            dsary[cd.WorkSheetName] = wsname
            dsary[cd.PrimaryDS] = primary[1]
            print(','.join(dsary), file=csvio)
        else:
            for j in range(len(secondaries)):
                dsary[cd.WorkSheetName] = wsname
                dsary[cd.PrimaryDS] = primary[1]
                secondary = secondaries[j]
                dsary[cd.SecondaryDS] = secondary[1]
                pskey = primary[1] + '.' + secondary[1]
                if pskey in rs_dict:
                    pkeys = ';'.join(rs_dict[pskey][0])
                    skeys = ';'.join(rs_dict[pskey][1])
                    dsary[cd.PrimaryKey] = pkeys
                    dsary[cd.SecondaryKey] = skeys
                else:
                    mkeys = ';'.join(get_mapset(primary, secondary).union(get_mapset(secondary, primary)))
                    dsary[cd.PrimaryKey] = mkeys
                    dsary[cd.SecondaryKey] = mkeys

                print(','.join(dsary), file=csvio)

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')
        
