from __future__ import annotations
from enum import IntEnum, auto

class DisplaysType(IntEnum):
    Rows = 0
    Cols = auto()
    Text = auto()
    Tooltip = auto()
    Color = auto()
    Detail = auto()
    Size = auto()
    Shape = auto()
    Angle = auto()

    @staticmethod
    def get_tagstr(dtype: DisplaysType):
        res = ''
        if dtype == DisplaysType.Rows:
            res = 'rows'
        elif dtype == DisplaysType.Cols:
            res = 'cols'
        elif dtype == DisplaysType.Text:
            res = 'text'
        elif dtype == DisplaysType.Tooltip:
            res = 'tooltip'
        elif dtype == DisplaysType.Color:
            res = 'color'
        elif dtype == DisplaysType.Detail:
            res = 'lod'
        elif dtype == DisplaysType.Size:
            res = 'size'
        elif dtype == DisplaysType.Shape:
            res = 'shape'
        elif dtype == DisplaysType.Angle:
            res = 'wedge-size'

        return res

    @staticmethod
    def get_capstr(dtype: DisplaysType):
        res = ''
        if dtype == DisplaysType.Rows:
            res = 'Rows'
        elif dtype == DisplaysType.Cols:
            res = 'Cols'
        elif dtype == DisplaysType.Text:
            res = 'Text'
        elif dtype == DisplaysType.Tooltip:
            res = 'Tooptip'
        elif dtype == DisplaysType.Color:
            res = 'Color'
        elif dtype == DisplaysType.Detail:
            res = 'Detail'
        elif dtype == DisplaysType.Size:
            res = 'Size'
        elif dtype == DisplaysType.Shape:
            res = 'Shape'
        elif dtype == DisplaysType.Angle:
            res = 'Angle'
        
        return res