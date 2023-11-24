from enum import IntEnum, auto

class ColFields(IntEnum):
    DataSourceName = 0
    FieldName = auto()
    role = auto()
    type = auto()
    Formula = auto()

    @staticmethod
    def get_colstr():
        colary = []
        for col in ColFields:
            colary.append(col.name)

        return ','.join(colary)

    @staticmethod
    def convert_short_datatype(dtype: str):
        dt = dtype
        if dtype == 'string':
            dt = 'STR'
        elif dtype == 'integer':
            dt = 'INT'
        elif dtype == 'real':
            dt = 'REAL'
        elif dtype == 'boolean':
            dt = 'BOOL'
        elif dtype == 'date':
            dt = 'DATE'
        elif dtype == 'datatime':
            dt = 'TIME'
        return dt