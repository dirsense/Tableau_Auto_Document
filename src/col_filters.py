from enum import IntEnum, auto

class ColFilters(IntEnum):
    ApplyTo = 0
    DataSourceName = auto()
    FieldName = auto()
    ElementL = auto()
    ElementR = auto()
    Type = auto()
    General = auto()
    General_List = auto()
    General_ExcludeList = auto()
    General_ExcludeCheck = auto()
    Wildcard = auto()
    Wildcard_Match = auto()
    Wildcard_Exclude = auto()
    Condition = auto()
    Top = auto()
    Top_Count = auto()
    Top_Bottom = auto()
    OR_NULL = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColFilters:
            colary.append(col.name)
        
        return ','.join(colary)