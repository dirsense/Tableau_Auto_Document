from tab_util import TabUtil as tu

def outcsv():
    csvpath = 'Output/csv/viewsheets.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print('Display Sheet,Configuration Sheet', file=csvio)

    dashshset = set()

    for k, v in tu.usedsheets_in_dashboards.items():
        dashshset |= v
        line = k + ',' + ';'.join(v)
        print(line, file=csvio)

    # Only list sheets that are not used on any dashboard
    for w in tu.twbx.worksheets:
        if not w in dashshset:
            print(w + ',', file=csvio)

    csvio.close()

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')