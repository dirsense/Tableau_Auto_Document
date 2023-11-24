from tableaudocumentapi import Workbook
from tableaudocumentapi.field import Field
from myfield import MyField
import tableaudocumentapi.xfile as tabx
import xml.etree.ElementTree as ET
import re

class TabUtil:
    workbook_path = ''
    twbx = None
    root: ET.Element = None
    ds_dict = {}
    field_dict = {}
    param_dict = {}
    direct_dict = {}
    internal_dict = {}
    inner_square_pattern = '(?<=\[).+?(?=\])'
    usedsheets_in_datasources = {}
    usedsheets_in_dashboards = {}

    @classmethod
    def initialize(cls, twbpath: str):
        cls.workbook_path = ''
        cls.twbx = None
        cls.root: ET.Element = None
        cls.ds_dict = {}
        cls.field_dict = {}
        cls.param_dict = {}
        cls.direct_dict = {}
        cls.usedsheets_in_datasources = {}
        cls.usedsheets_in_dashboards = {}

        cls.create_id_dict(twbpath)
        cls.create_usedsheets_in_datasources()
        cls.create_usedsheets_in_dashboards()

    # Create a sheet set used for each data source.
    @classmethod
    def create_usedsheets_in_datasources(cls):
        for dsid in cls.ds_dict.keys():
            dkey = str(dsid).strip('[]')
            shset = cls.create_usedsheets_in_datasource(dkey)
            if len(shset) > 0:
                cls.usedsheets_in_datasources[dkey] = shset

    @classmethod
    def create_usedsheets_in_datasource(cls, datasource_id: str):
        viewset = set()
        for w in cls.root.iter('worksheet'):
            shname = w.attrib['name']
            for d in w.iter('datasource-dependencies'):
                if datasource_id == d.attrib['datasource']:
                    for ci in d.iter('column-instance'):
                        viewset.add(shname)
                    break
        return viewset

    # Create a sheet set used for each dashboard.
    @classmethod
    def create_usedsheets_in_dashboards(cls):
        for dash in cls.root.iter('dashboard'):
            dashname = dash.attrib['name']
            shset = cls.create_usedsheets_in_dashboard(dashname)
            if len(shset) > 0:
                cls.usedsheets_in_dashboards[dashname] = shset

    @classmethod
    def create_usedsheets_in_dashboard(cls, dashboard_name: str):
        viewset = set()
        for w in cls.root.iter('window'):
            if 'class' in w.attrib and w.attrib['class'] == 'dashboard' and w.attrib['name'] == dashboard_name:
                for vp in w.iter('viewpoint'):
                    viewset.add(vp.attrib['name'])
                break
        return viewset

    @classmethod
    def get_sourceinfo_by_action(cls, atag: ET.Element):
        asrc = atag.find('source')
        atype = asrc.attrib['type']
        stype = atype

        incset = set()
        excset = set()
        for exc in asrc.iter('exclude-sheet'):
            excset.add(exc.attrib['name'])
        
        if atype == 'all':
            incset = set(cls.twbx.worksheets) ^ excset
        elif atype == 'datasource':
            dsid = asrc.attrib['datasource']
            if dsid in cls.usedsheets_in_datasources:
                incset = cls.usedsheets_in_datasources[dsid] ^ excset
        elif atype == 'sheet':
            if 'dashboard' in asrc.attrib:
                stype = 'dashboard'
                if 'worksheet' in asrc.attrib:
                    incset.add(asrc.attrib['worksheet'])
                else:
                    dashname = asrc.attrib['dashboard']
                    if dashname in cls.usedsheets_in_dashboards:
                        incset = cls.usedsheets_in_dashboards[dashname] ^ excset
            else:
                incset.add(asrc.attrib['worksheet'])

        return stype, incset

    @classmethod
    def get_targettype_by_action(cls, atag: ET.Element):
        if atag.find('activation') is None:
            return 'Menu'
        else:
            return 'Select' if atag.find('activation').attrib['type'] == 'on-select' else 'Hover'
    
    @classmethod
    def get_singleclick_by_action(cls, atag: ET.Element):
        single_click = ''
        if atag.find('command') is None:
            return ''
        for p in atag.find('command').iter('param'):
            if p.attrib['name'] == 'single-select':
                single_click = 'TRUE'
                break
        return single_click
    
    @classmethod
    def get_targetsheet_by_action(cls, atag: ET.Element):
        acmd = atag.find('command')
        tval = ''
        incset = set()
        excset = set()

        for prm in acmd.iter('param'):
            if prm.attrib['name'] == 'target':
                tval = prm.attrib['value']
            if prm.attrib['name'] == 'exclude':
                excset |= set(prm.attrib['value'].split(','))

        if tval in cls.usedsheets_in_dashboards:
            incset = cls.usedsheets_in_dashboards[tval] ^ excset
        else:
            incset.add(tval)
        
        return incset

    @classmethod
    def create_id_dict(cls, workbook_path: str):
        cls.twbx = Workbook(workbook_path)
        cls.root: ET.Element = tabx.xml_open(workbook_path).getroot()
        fkey = ''
        fkeys = set()

        for dtag in cls.root.find('datasources').iter('datasource'):
            if dtag.attrib['name'] == 'Parameters':
                for c in dtag.findall('column'):
                    cls.param_dict[c.attrib['name']] = '[' + c.attrib['caption'] + ']'
            else:
                dsid = '[' + dtag.attrib['name'] + ']'
                if 'caption' in dtag.attrib:
                    cls.ds_dict[dsid] = '[' + dtag.attrib['caption'] + ']'
                else:
                    cls.ds_dict[dsid] = dsid
                f_dict = {}
                for c in dtag.findall('column'):
                    fkey = c.attrib['name']
                    if 'caption' in c.attrib:
                        fcap = '[' + c.attrib['caption'] + ']'
                    else:
                        fcap = fkey
                    f_dict[fkey] = fcap
                    fkeys.add(fkey)

                cls.field_dict[dsid] = f_dict

        # Create a dictionary of fields that cannot be retrieved by standard.
        for col in cls.root.iter('_.fcp.ObjectModelTableType.true...column'):
            if 'caption' in col.attrib and 'name' in col.attrib:
                interid = col.attrib['name']
                if '[__tableau_internal_object_id__].' in interid:
                    interkey = interid.replace('[__tableau_internal_object_id__].', '')
                    if not interkey in cls.internal_dict:
                        cls.internal_dict[interkey] = col.attrib['caption']

        # Create a separate dictionary for ad-hoc formulas since they are not listed in field_dict.
        for col in cls.root.iter('column'):
            if 'name' in col.attrib and 'caption' in col.attrib and 'role' in col.attrib and 'datatype' in col.attrib:
                fkey = col.attrib['name']
                if not fkey in fkeys and not fkey in cls.direct_dict:
                    formula = ''
                    cfc = col.find('calclation')
                    if not cfc is None and 'formula' in cfc.attrib:
                        formula = cfc.attrib['formula']
                    cls.direct_dict[fkey] = MyField(col.attrib['caption'], col.attrib['datatype'], col.attrib['role'], formula)

    # If the field has a constant directly in the expression, it cannot be taken in the normal tag, so loop through the column tags to get the ones with matching IDs.
    @classmethod
    def find_direct_field(cls, id: str):
        for col in cls.root.iter('column'):
            if 'name' in col.attrib and col.attrib['name'] == id and 'caption' in col.attrib:
                return col.attrib['caption']
        return ''
    
    # [none:FieldName:nk]
    # Divide a string like this into three parts inside parentheses based on a colon (FieldName, none, nk).
    @classmethod
    def get_coron_tuple(cls, text: str):
        if 'Measure Names' in text:
            return ('Measure Names', '', '')
        elif 'Multiple Values' in text:
            return ('Multiple Values', '', '')

        # If there are three colons, make two.
        coron2text = text
        if text.count(':') == 3:
            exclude_part = text[1:text.find(':') + 1]
            coron2text = text.replace(exclude_part, '')

        colonpos1 = coron2text.find(':')
        if colonpos1 == -1:
            return (coron2text, '', '')
        colonpos2 = coron2text[colonpos1 + 1:].find(':')
        fieldname = coron2text[colonpos1 + 1 : colonpos1 + colonpos2 + 1]
        attrL = coron2text[:colonpos1]
        attrR = coron2text[colonpos1 + colonpos2 + 2:]
        return (fieldname, attrL, attrR)

    # '[sqlproxy.154uepn1feytms1acmtm41ahvc7d].[none:FieldName:nk]'
    # Given a string like this, process it by considering the first bracket as the data source ID and the second bracket as the field ID.
    # Return the data source and field caption values and field attribute values.
    @classmethod
    def get_ds_fields_by_coron_str(cls, text: str):
        dsfs = []
        cur_dsid = dcap = fcap = ''
        # If internal~ is present, there should be three pairs of square brackets, so remove them.
        target_text = text.replace('[__tableau_internal_object_id__].', '')

        for count, rf in enumerate(re.findall(cls.inner_square_pattern, target_text)):
            if count % 2 == 0:
                cur_dsid = '[' + rf + ']'
                if cur_dsid in cls.ds_dict:
                    dcap = cls.ds_dict[cur_dsid]
                else:
                    dcap = cur_dsid
            else:
                ftuple = cls.get_coron_tuple(rf)
                fkey = '[' + ftuple[0] + ']'

                if cur_dsid == '[Parameters]':
                    # Sometimes a target parameter is set to "none" in the workbook, but refers to a garbage parameter that does not exist in XML.
                    # In such a case, the existence of the parameter in cls.param_dict is used to determine if the parameter exists. If not, return empty and exit.
                    if not fkey in cls.param_dict:
                        return []
                    fcap = cls.param_dict[fkey]

                elif fkey in cls.field_dict[cur_dsid]:
                    fcap = cls.field_dict[cur_dsid][fkey]
                else:
                    if 'Calculation' in fkey:
                        dircol = cls.find_direct_field(fkey)
                        fcap = fkey if dircol == '' else dircol
                    else:
                        fcap = fkey
                dsfs.append((dcap, fcap, ftuple[1], ftuple[2]))

        return dsfs

    @classmethod
    def get_display_fields_by_coron_str(cls, text: str):
        dsfs = []
        cur_dsid = dcap = fcap = ''
        # If internal~ is present, there should be three pairs of square brackets, so remove them.
        target_text = text.replace('[__tableau_internal_object_id__].', '')

        for count, rf in enumerate(re.findall(cls.inner_square_pattern, target_text)):
            if count % 2 == 0:
                cur_dsid = '[' + rf + ']'
                if cur_dsid in cls.ds_dict:
                    dcap = cls.ds_dict[cur_dsid]
                else:
                    dcap = cur_dsid
            else:
                myf = MyField()
                ftuple = cls.get_coron_tuple(rf)
                fkey = '[' + ftuple[0] + ']'

                if cur_dsid == '[Parameters]':
                    if not fkey in cls.param_dict:
                        return []
                    fcap = cls.param_dict[fkey]
                
                elif fkey in cls.field_dict[cur_dsid]:
                    fcap = cls.field_dict[cur_dsid][fkey]
                
                elif fkey in cls.direct_dict:
                    myf: MyField = cls.direct_dict[fkey]
                    fcap = myf.caption
                    if myf.formula != '':
                        myf.formula = cls.replace_calc_field_id_to_caption(myf.formula, cur_dsid)

                elif fkey in cls.internal_dict:
                    fcap = 'COUNT (' + cls.internal_dict[fkey] + ')'
                    myf.role = 'M'
                    myf.datatype = 'integer'
                else:
                    fcap = fkey
                dsfs.append((dcap, fcap, ftuple[1], ftuple[2], myf.formula, myf.role, myf.datatype))

        return dsfs

    @classmethod
    def replace_calc_field_id_to_caption(cls, formula: str, primary_dsid: str):
        blend_square_ptn = '\[[^\]]+\]\.\[[^\]]+\]'
        square_ptn = '\[.+?\]'

        blend_tmp_names = []
        blend_target_ids = []
        blend_captions = []

        ### ------------------------------------------------------------------
        # Replace blend or parameter fields with temporary strings first
        count = 0
        for rf in re.findall(blend_square_ptn, formula):
            if rf in blend_target_ids:
                continue

            ids = re.findall(square_ptn, rf)
            dsid = ids[0]
            fid = ids[1]
            caption = ''
            if dsid in cls.ds_dict:
                dcap = cls.ds_dict[dsid]
                if fid in cls.field_dict[dsid]:
                    caption = dcap + '.' + cls.field_dict[dsid][fid]
            elif dsid == '[Parameters]':
                if fid in cls.param_dict:
                    caption = cls.param_dict[fid]

            elif dsid == '[__tableau_internal_object_id__]':
                if fid in cls.internal_dict:
                    caption = '[' + cls.internal_dict[fid] + ']'

            if caption != '':
                blend_target_ids.append(rf)
                blend_tmp_names.append('blend-replace-' + str(count))                
                blend_captions.append(caption)
                count += 1
            else:
                # This is a branch that is unnecessary because it should not come here, but I'm leaving it here for checking.
                print('The blend field in the formula has not been replaced')

        for count, id in enumerate(blend_target_ids):
            formula = formula.replace(id, blend_tmp_names[count])
        ### ------------------------------------------------------------------

        # Replace Primary fields other than Blend or Parameter with ID => Caption.
        fml = formula
        for id in re.findall(square_ptn, fml):
            fdict = cls.field_dict[primary_dsid]
            if id in fdict:
                formula = formula.replace(id, fdict[id])

        # Replace temporary string => caption only when blend or Parameter field is present.
        for count, tmp in enumerate(blend_tmp_names):
            formula = formula.replace(tmp, blend_captions[count])

        return '"' + formula.replace('"', '""') + '"'
