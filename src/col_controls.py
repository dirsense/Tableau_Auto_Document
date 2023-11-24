from enum import IntEnum, auto

class ColControls(IntEnum):
    DisplaySheetName = 0
    SourceSheetName = auto()
    DataSourceName = auto()
    FieldName = auto()
    ControlName = auto()
    ControlShortName = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColControls:
            colary.append(col.name)

        return ','.join(colary)