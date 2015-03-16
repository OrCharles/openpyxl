
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
from .LineSer import *

class ChartLines(Serialisable):

    as_3D = True

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr



class BarChart(Serialisable):

    gapWidth = Typed(expected_type=GapAmount, allow_none=True)
    gapDepth = Typed(expected_type=GapAmount, allow_none=True) #3d
    overlap = Typed(expected_type=Overlap, allow_none=True) #2d
    shape = Typed(expected_type=Shape, allow_none=True) #3d
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    axId = Typed(expected_type=UnsignedInt, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'gapDepth', 'overlap', 'serLines', 'axId', 'extLst')

    def __init__(self,
                 gapWidth=None,
                 overlap=None,
                 serLines=None,
                 axId=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.overlap = overlap
        self.serLines = serLines
        self.axId = axId
        self.extLst = extLst

