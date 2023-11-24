from enum import Enum

class UsrKey(Enum):
    UiEnum = '{http://www.tableausoftware.com/xml/user}ui-enumeration'
    UiCustom = '{http://www.tableausoftware.com/xml/user}ui-manual-selection'
    UiPatText = '{http://www.tableausoftware.com/xml/user}ui-pattern_text'
    UiPatType = '{http://www.tableausoftware.com/xml/user}ui-pattern_type'
    UiMarker = '{http://www.tableausoftware.com/xml/user}ui-marker'
    UiDomain = '{http://www.tableausoftware.com/xml/user}ui-domain'
    UiTopByField = '{http://www.tableausoftware.com/xml/user}ui-top-by-field'