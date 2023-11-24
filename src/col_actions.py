from enum import IntEnum, auto

class ColActions(IntEnum):
    SourceType = 0
    ActionType = auto()
    ActionName = auto()
    SourceSheet = auto()
    TargetSheet = auto()
    SourceField = auto()
    TargetField = auto()
    ActionTarget = auto()
    ClearingTheSelectionWill = auto()
    SingleSelect = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColActions:
            colary.append(col.name)

        return ','.join(colary)