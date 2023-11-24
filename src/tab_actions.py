from col_actions import ColActions as ca
from tab_util import TabUtil as tu
import urllib.parse, re, itertools

def outcsv():
    csvpath = 'Output/csv/actions.csv'
    csvio = open(csvpath, 'w', encoding='utf-8')
    print(ca.get_colstr(), file=csvio)

    action_chain = itertools.chain(tu.root.iter('action'), tu.root.iter('nav-action'), tu.root.iter('edit-parameter-action'), tu.root.iter('edit-group-action'))

    for r in action_chain:
        actary = [''] * len(ca)

        actary[ca.ActionName] = r.attrib['caption']

        srctype, srcset = tu.get_sourceinfo_by_action(r)
        actary[ca.SourceType] = srctype
        actary[ca.SourceSheet] = ';'.join(srcset)

        actary[ca.ActionTarget] = tu.get_targettype_by_action(r)
        actary[ca.SingleSelect] = tu.get_singleclick_by_action(r)

        # Filter Action
        if not r.find('command') is None and r.find('command').attrib['command'] == 'tsc:tsl-filter':
            actary[ca.ActionType] = 'Filter'
            actary[ca.TargetSheet] = ';'.join(tu.get_targetsheet_by_action(r))

            if r.find('link') is None:
                actary[ca.SourceField] = 'All'
                actary[ca.TargetField] = 'All'
            else:
                links = urllib.parse.unquote(r.find('link').attrib['expression'])
                link = links.split('~na>&')
                srcary = []
                tgtary = []
                for l in link:
                    l_fields = re.findall(tu.inner_square_pattern, l)
                    if len(l_fields) == 4:
                        srcary.append(l_fields[3])
                        tgtary.append(l_fields[1])

                actary[ca.SourceField] = ';'.join(srcary)
                actary[ca.TargetField] = ';'.join(tgtary)

            if r.find('activation') is None:
                actary[ca.ClearingTheSelectionWill] = 'None'
            else:
                if 'auto-clear' in r.find('activation').attrib:
                    filter_result = 'Show all values'
                    for p in r.find('command').iter('param'):
                        if p.attrib['name'] == 'on-empty':
                            filter_result = 'Exclude all values'
                            break
                    actary[ca.ClearingTheSelectionWill] = filter_result
                else:
                    actary[ca.ClearingTheSelectionWill] = 'Keep filtered values'

        # Highlight Action
        elif not r.find('command') is None and r.find('command').attrib['command'] == 'tsc:brush':
            actary[ca.ActionType] = 'Highlight'
            actary[ca.TargetSheet] = ';'.join(tu.get_targetsheet_by_action(r))

            for p in r.find('command').iter('param'):
                if p.attrib['name'] == 'special-fields':
                    actary[ca.TargetField] = 'All' if p.attrib['value'] == 'all' else 'Dates and Times'
                elif p.attrib['name'] == 'field-captions':
                    actary[ca.TargetField] = p.attrib['value'].replace(',', ';')

        # Go to URL Action
        elif not r.find('link') is None and 'caption' in r.find('link').attrib and r.find('link').attrib['caption'] == '':
            actary[ca.ActionType] = 'GoToURL'
            linktag = r.find('link')
            actary[ca.TargetSheet] = 'New Tab if No Web Page Object Exists' if linktag.find('url-action-type') is None else 'New Browser Tab'
            if 'expression' in linktag.attrib:
                actary[ca.TargetField] = linktag.attrib['expression']

        # Go to Sheet Action
        elif r.tag == 'nav-action':
            actary[ca.ActionType] = 'GoToSheet'
            for p in r.iter('param'):
                if p.attrib['name'] == 'sheet':
                    actary[ca.TargetField] = p.attrib['value']
                    break

        # Parameter Action
        elif r.tag == 'edit-parameter-action':
            actary[ca.ActionType] = 'Parameter'
            for p in r.iter('param'):
                if 'value' in p.attrib:
                    df = tu.get_ds_fields_by_coron_str(p.attrib['value'])
                    if len(df) == 0:
                        pv = 'None'
                    else:
                        pv = str(df[0][1]).strip('[]')
                    if p.attrib['name'] == 'source-field':
                        actary[ca.TargetField] = pv
                    elif p.attrib['name'] == 'target-parameter':
                        pids = re.findall('\[.+?\]', p.attrib['value'])
                        if len(pids) == 2 and pids[1] in tu.param_dict:
                            actary[ca.TargetSheet] = tu.param_dict[pids[1]]

        # Set Action
        elif r.tag == 'edit-group-action':
            actary[ca.ActionType] = 'Set'
            tset = 'None'
            for p in r.iter('param'):
                if p.attrib['name'] == 'target-group':
                    tset = str(tu.get_ds_fields_by_coron_str(p.attrib['value'])[0][1]).strip('[]')
                    break
            actary[ca.TargetSheet] = tset

        if len(actary) > 0:
            if actary[ca.TargetSheet] == 'None':
                actary[ca.TargetSheet] = ''
            print(','.join(actary), file=csvio)

    csvio.close()

if __name__ == '__main__':
    tu.initialize('test.twbx')
    outcsv()
    print('Done')