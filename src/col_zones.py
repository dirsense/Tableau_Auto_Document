from enum import IntEnum, auto

class ColZones(IntEnum):
    DashboardName = 0
    WorkSheetName = auto()
    W = auto()
    H = auto()
    X = auto()
    Y = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColZones:
            colary.append(col.name)

        return ','.join(colary)