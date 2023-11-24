from enum import IntEnum, auto

class ColDatasources(IntEnum):
    WorkSheetName = 0
    PrimaryDS = auto()
    SecondaryDS = auto()
    PrimaryKey = auto()
    SecondaryKey = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColDatasources:
            colary.append(col.name)

        return ','.join(colary)