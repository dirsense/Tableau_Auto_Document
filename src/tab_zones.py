from col_zones import ColZones as cz
from tab_util import TabUtil as tu

def outcsv():
    nameset = set()
    csvpath = 'Output/csv/zones.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(cz.get_colstr(), file=csvio)

    for dash in tu.root.iter('dashboard'):
        for z in dash.iter('zone'):
            if not 'type-v2' in z.attrib and 'name' in z.attrib:
                exclude_flg = False
                for za in z.attrib:
                    if '_.' in za[:2]:
                        exclude_flg = True
                        break
                if exclude_flg:
                    continue

                namekey = dash.attrib['name'] + '.' + z.attrib['name']
                if not namekey in nameset:
                    nameset.add(namekey)
                    zary = [''] * len(cz)
                    zary[cz.DashboardName] = dash.attrib['name']
                    zary[cz.WorkSheetName] = z.attrib['name']
                    zary[cz.W] = z.attrib['w']
                    zary[cz.H] = z.attrib['h']
                    zary[cz.X] = z.attrib['x']
                    zary[cz.Y] = z.attrib['y']
                    print(','.join(zary), file=csvio)
    csvio.close()

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')