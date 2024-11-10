# patch for the openpyxl library to handle the 'xxid' tag
# preventing the folowing error when loading an Excel file
# TypeError: CellStyle.__init__() got an unexpected keyword argument 'xxid'
# code to run after having imported the openpyxl library

from openpyxl.styles.cell_style import CellStyle

CellStyle.__attrs__ = ("numFmtId", "fontId", "fillId", "borderId",
                       "applyAlignment", "applyProtection", "pivotButton", "quotePrefix", "xfId",
                       "xxid")  # add xxid tag

def new_init(self,
             numFmtId=0,
             fontId=0,
             fillId=0,
             borderId=0,
             xfId=None,
             xxid=None,  # add xxid tag
             quotePrefix=None,
             pivotButton=None,
             applyNumberFormat=None,
             applyFont=None,
             applyFill=None,
             applyBorder=None,
             applyAlignment=None,
             applyProtection=None,
             alignment=None,
             protection=None,
             extLst=None,
            ):
    self.numFmtId = numFmtId
    self.fontId = fontId
    self.fillId = fillId
    self.borderId = borderId
    self.xfId = xfId
    self.xxid = xxid  # add xxid tag
    self.quotePrefix = quotePrefix
    self.pivotButton = pivotButton
    self.applyNumberFormat = applyNumberFormat
    self.applyFont = applyFont
    self.applyFill = applyFill
    self.applyBorder = applyBorder
    self.alignment = alignment
    self.protection = protection

CellStyle.__init__ = new_init
