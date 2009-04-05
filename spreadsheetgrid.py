import wx
import wx.grid

from spreadsheet import SpreadSheet, SpreadSheetError, SpreadSheetCell

class ButtonRenderer(wx.grid.PyGridCellRenderer):
    def __init__(self):
        wx.grid.PyGridCellRenderer.__init__(self)

    def DrawBezel(self, grid, dc, rect, isSelected):
        if isSelected:
            state = wx.CONTROL_PRESSED
        else:
            state = 0
        #if not self.IsEnabled():
            #state = wx.CONTROL_DISABLED
        #pt = grid.ScreenToClient(wx.GetMousePosition())
        #if rect.Contains(pt):
        #if isSelected:
            #state = wx.CONTROL_CURRENT
        print state
        wx.RendererNative.Get().DrawPushButton(grid, dc, rect, state)

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        text = grid.GetCellValue(row, col)
        hAlign, vAlign = attr.GetAlignment()
        vAlign = wx.ALIGN_CENTER
        hAlign = wx.ALIGN_CENTER
        dc.SetFont(attr.GetFont())
        bg = grid.GetSelectionBackground()
        #fg = grid.GetSelectionForeground()
        #bg = wx.WHITE
        fg = wx.BLACK
        dc.SetTextBackground(bg)
        dc.SetTextForeground(fg)
        #dc.SetBrush(wx.Brush(bg, wx.SOLID))
        #dc.SetPen(wx.TRANSPARENT_PEN)
        #dc.DrawRectangleRect(rect)
        self.DrawBezel(grid, dc, rect, isSelected)
        grid.DrawTextRectangle(dc, text, rect, hAlign, vAlign)

    def GetBestSize(self, grid, attr, dc, row, col):
        return (10, 10)

    def Clone(self):
        return ButtonRenderer()

class SpreadSheetCellEditor(wx.grid.PyGridCellEditor):
    def __init__(self):
        wx.grid.PyGridCellEditor.__init__(self)

    def Create(self, parent, id, evtHandler):
        self._tc = wx.TextCtrl(parent, id, '')
        self._tc.SetInsertionPoint(0)
        self.SetControl(self._tc)
        if evtHandler:
            self._tc.PushEventHandler(evtHandler)
        
    def SetSize(self, rect):
        self._tc.SetDimensions(rect.x, rect.y, rect.width+2, rect.height+2,
                               wx.SIZE_ALLOW_MINUS_ONE)

    def Show(self, show, attr):
        wx.grid.PyGridCellEditor.Show(self, show, attr)

    def StartingKey(self, evt):
        keycode = evt.GetUnicodeKey()
        if keycode < 128:
            keycode = evt.GetKeyCode()
        cha = chr(keycode) if keycode < 0x80 else unichr(keycode)

        self._tc.SetValue(self._tc.GetValue() + cha)
        self._tc.SetInsertionPointEnd()

    def StartingClick(self):
        #print 'StartingClick'
        pass

    def BeginEdit(self, row, col, grid):
        if grid.GetTable().IsEmptyCell(row, col):
            self.startValue = ''
        else:
            self.startValue = grid.GetTable().GetFormula(row, col)
        self._tc.SetValue(self.startValue)
        self._tc.SetInsertionPointEnd()
        self._tc.SetFocus()

    def EndEdit(self, row, col, grid):
        val = self._tc.GetValue()
        changed = val != self.startValue
        if changed:
            if val.strip():
                grid.GetTable().SetFormula(row, col, val) # update the table
            else:
                grid.GetTable().DeleteCell(row, col)
            #wx.CallAfter(grid.updateSelectedCell(row, col))
            wx.CallAfter(grid.ForceRefresh)

        self.startValue = ''
        self._tc.SetValue('')
        return changed

    def Reset(self):
        self._tc.SetValue(self.startValue)
        self._tc.SetInsertionPointEnd()

    def Clone(self):
        return SpreadSheetCellEditor()


class SpreadSheetTable(wx.grid.PyGridTableBase):
    def __init__(self, spreadsheet, cell_attrs):
        wx.grid.PyGridTableBase.__init__(self)
        self._selected_cell = None
        # Sets of cell position that should have other cell formatting/colours
        self._dependents_pos = set()
        self._precedents_pos = set()
        self._errors_pos = set()
        self._spreadsheet = spreadsheet
        self._cell_attrs = cell_attrs

    def GetNumberRows(self):
        return 99
        
    def GetNumberCols(self):
        return 25

    def IsEmptyCell(self, row, col):
        cellname = self._spreadsheet.getCellName(row, col)
        return self._spreadsheet.getCell(cellname) == None

    def GetFormula(self, row, col):
        cellname = self._spreadsheet.getCellName(row, col)
        cell = self._spreadsheet.getCell(cellname)
        return cell.getFormula() if cell else ''

    def GetValue(self, row, col):
        cellname = self._spreadsheet.getCellName(row, col)
        cell = self._spreadsheet.getCell(cellname)
        cell_pos = (row, col)
        was_formatted_as_error = cell_pos in self._errors_pos
        try:
            self._errors_pos.discard( cell_pos )
            val = cell.getValue() if cell else ''
        except (Exception, SyntaxError, SpreadSheetError), e:
            self._errors_pos.add( cell_pos )
            val = str(e)
        is_formatted_as_error = cell_pos in self._errors_pos
        if was_formatted_as_error != is_formatted_as_error:
            wx.CallAfter(self.GetView().ForceRefresh)
        return val

    def GetAttr(self, row, col, kind):
        pos = (row, col)
        attr = self._cell_attrs['normal']
        if pos in self._errors_pos:
            attr = self._cell_attrs['error']
        elif pos in self._precedents_pos:
            attr = self._cell_attrs['pre']
        elif pos in self._dependents_pos:
            attr = self._cell_attrs['dep']
        else:
            cellname = self._spreadsheet.getCellName(row, col)
            cell = self._spreadsheet.getCell(cellname)
            if cell:
                if cell.getValue() == 'button':
                    attr = self._cell_attrs['buttons']
                else:
                    precedents = len(cell.getPrecedents())
                    dependents = len(cell.getDependents())
                    if (precedents == 0) and (dependents == 0):
                        attr = self._cell_attrs['alone']
                    elif (precedents > 0) and (dependents > 0):
                        attr = self._cell_attrs['intermediate']
                    # Only show global input and output when no cell is selected
                    elif not self._selected_cell:
                        if precedents == 0:
                            attr = self._cell_attrs['input']
                        elif dependents == 0:
                            attr = self._cell_attrs['output']
        attr.IncRef()
        return attr
    
    def SetFormula(self, row, col, formula):
        cellname = self._spreadsheet.getCellName(row, col)
        self._spreadsheet.setCellFormula(cellname, formula)

    def DeleteCell(self, row, col):
        cellname = self._spreadsheet.getCellName(row, col)
        self._spreadsheet.deleteCell(cellname)

    def setSelectedCell(self, row, col):
        cellname = self._spreadsheet.getCellName(row, col)
        old_selected_cell = self._selected_cell
        self._selected_cell = self._spreadsheet.getCell(cellname)
        self._dependents_pos.clear()
        self._precedents_pos.clear()
        if self._selected_cell:
            name2pos = self._spreadsheet.getCellNamePos
            deps = self._selected_cell.getDependents(include_supporting=False)
            self._dependents_pos.update(set([name2pos(c.getName()) 
                                         for c in deps]))
            precs = self._selected_cell.getPrecedents(include_supporting=False)
            self._precedents_pos.update(set([name2pos(c.getName()) 
                                         for c in precs]))
        # Needs refresh if another cell was selected
        return old_selected_cell != self._selected_cell

class SpreadSheetGrid(wx.grid.Grid):
    def __init__(self, parent, spreadsheet, selection_callbacks=None):
        wx.grid.Grid.__init__(self, parent, -1)
        self._selection_callbacks = selection_callbacks or []
        font_size = self.GetCellFont(0, 0).GetPointSize()

        def bg_attr(col):
            attr = wx.grid.GridCellAttr()
            attr.SetBackgroundColour(col)
            return attr

        def button_attr():
            attr = wx.grid.GridCellAttr()
            attr.SetBackgroundColour('#FFFFFF')
            #attr.SetForegroundColour('#000000')
            attr.SetRenderer(ButtonRenderer())
            attr.SetFont(wx.Font(font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            return attr

        attrs = {}
        attrs['normal'] = wx.grid.GridCellAttr()
        attrs['intermediate'] = wx.grid.GridCellAttr()
        attrs['intermediate'].SetTextColour('#444444')
        attrs['pre'] = bg_attr('#DDDDFF')
        attrs['dep'] = bg_attr('#FFDDDD')
        attrs['error'] = bg_attr('#FF8888')
        # TODO: Something is wrong(on OSX). Not same size as default font. :-(
        attrs['error'].SetFont(wx.Font(font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL))
        attrs['alone'] = bg_attr('#EEEEEE')
        attrs['alone'].SetFont(wx.Font(font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        attrs['input'] = bg_attr('#FFFFDD')
        attrs['output'] = bg_attr('#DDFFDD')
        attrs['buttons'] = button_attr()

        self._table = SpreadSheetTable(spreadsheet, attrs)

        self.SetTable(self._table, True)
        self.SetDefaultEditor(SpreadSheetCellEditor())
        self.EnableDragGridSize(False)
        self.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.OnGridSelectCell)
        self.Bind(wx.EVT_MENU, self.OnEditCut, id=wx.ID_CUT)
        self.Bind(wx.EVT_MENU, self.OnEditCopy, id=wx.ID_COPY)
        self.Bind(wx.EVT_MENU, self.OnEditPaste, id=wx.ID_PASTE)
        self.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.OnMouse)

    def OnGridSelectCell(self, event):
        self.updateSelectedCell(event.GetRow(), event.GetCol())
        event.Skip()

    def updateSelectedCell(self, row, col):
        needs_refresh = self._table.setSelectedCell(row, col)
        cellname = self._table._spreadsheet.getCellName(row, col)
        for c in self._selection_callbacks:
            c(cellname, self._table._selected_cell)
        if needs_refresh:
            self.ForceRefresh()

    def OnMouse(self, event):
        edit, col, row = event.AltDown(), event.GetCol(), event.GetRow()
        print edit, col, row
        if self._table.GetValue(row, col) == 'button' and not edit:
        #if not shift:
            print 'button pressed'
        else:
            event.Skip()

    def OnEditCut(self, event):
        print 'SpreadSheetGrid.OnEditCut', event
        event.Skip()

    def OnEditCopy(self, event):
        print 'SpreadSheetGrid.OnEditCopy', event
        event.Skip()

    def OnEditPaste(self, event):
        print 'SpreadSheetGrid.OnEditPaste', event
        event.Skip()

    def paste(self, text):
        row, col = self.GetGridCursorRow(), self.GetGridCursorCol()
        self._table.SetFormula(row, col, text)
        self.ForceRefresh()

    def cut(self):
        row, col = self.GetGridCursorRow(), self.GetGridCursorCol()
        formula = self._table.GetFormula(row, col)
        self._table.DeleteCell(row, col)
        self.ForceRefresh()
        return formula

    def copy(self):
        row, col = self.GetGridCursorRow(), self.GetGridCursorCol()
        formula = self._table.GetFormula(row, col)
        return formula

if __name__ == '__main__':
    class TestFrame(wx.Frame):
        def __init__(self):
            wx.Frame.__init__(self, None, 
                              title='SpreadSheet Grid Editor', size=(640, 480))
            spreadsheet = SpreadSheet()
            spreadsheet.setCellFormula('a1', '2+3')
            spreadsheet.setCellFormula('a2', 'a1')
            spreadsheet.setCellFormula('a3', '"button"')
            spreadsheet._calculate()
            def cell_selection_callback(cell_name, cell):
                print cell_name, cell
            grid = SpreadSheetGrid(self, spreadsheet, [cell_selection_callback])
            #grid.SetCellRenderer(3, 1, ButtonRenderer())
    
    app = wx.App(redirect=False)
    frame = TestFrame()
    frame.Show()
    app.MainLoop()

