from enum import IntEnum, auto

class ColDisplays(IntEnum):
    WorkSheetName = 0
    use = auto()
    DataSourceName = auto()
    FieldName = auto()
    AdHocCalculatedField = auto()
    role = auto()
    type = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColDisplays:
            colary.append(col.name)

        return ','.join(colary)