from enum import IntEnum, auto

class ColParameters(IntEnum):
    ParameterName = 0
    Type = auto()
    AllowableValues = auto()
    WhenWorkbookOpens = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColParameters:
            colary.append(col.name)

        return ','.join(colary)