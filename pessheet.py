#!/usr/bin/python

# TODO
# * Make it possible to double click on pss-files
# * Autolabels (for dot-files)
# * When you press Del an ugly square appears at end of formula
# * Error-window below editor (sash). Don't use stdout for normal operation. 



# spreadsheet.py

from wx.lib import sheet
import wx
import wx.py.editor 

#import spreadsheet
from spreadsheet import SpreadSheet, SpreadSheetError

import spreadsheetgrid

def iconsize():
    return 16

def iconbitmap(name):
    return wx.Bitmap('icons%d/%s' % (iconsize(), name))

def iconbitmapsize():
    return (iconsize(), iconsize())

def conf_file_name(appname):
    import platform
    import os
    if platform.system() == 'Windows':
        # Get path to the application data directory and create a Todos
        # subdirectory if it does not exist.
        import ctypes
        CSIDL_APPDATA = 0x1a
        CSIDL_FLAG_CREATE = 0x8000
        SHGFP_TYPE_CURRENT = 0
        MAX_DATA_SIZE = 260
        T = ctypes.c_wchar * MAX_DATA_SIZE
        app_data = T()
        f = ctypes.windll.shell32.SHGetFolderPathW
        if f(0, CSIDL_APPDATA | CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, 
             app_data):
            raise RuntimeError('Failed to call SHGetFolderPathW()')
        directory = os.path.join(app_data.value, appname)
        if not os.path.exists(directory):
            os.mkdir(directory)
    else:
        directory = os.path.expanduser('~')
    return os.path.join(directory, '%s.conf' % appname)

class SheetPage(wx.Panel):
    def __init__(self, parent, spreadsheet):
        wx.Panel.__init__(self, parent, id=wx.ID_ANY)
        
        self.SetBackgroundColour(parent.GetBackgroundColour())

        cell_toolbar_line1 = wx.Panel(self)
        cell_toolbar_line1_sizer = wx.BoxSizer(wx.HORIZONTAL)
        cell_toolbar_line1.SetSizer(cell_toolbar_line1_sizer)
        self.selected_cell_textctrl = wx.TextCtrl(cell_toolbar_line1)
        self.selected_cell_textctrl.SetValue('a1')
        cell_toolbar_line1_sizer.Add(self.selected_cell_textctrl, 0, wx.CENTER)
        self.cell_formula_textctrl = wx.TextCtrl(cell_toolbar_line1, style=wx.TE_PROCESS_ENTER)
        cell_toolbar_line1_sizer.Add(self.cell_formula_textctrl, 1, wx.CENTER)

        cell_toolbar_line2 = wx.Panel(self)
        cell_toolbar_line2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        cell_toolbar_line2.SetSizer(cell_toolbar_line2_sizer)
        self.cell_value_type_textctrl = wx.TextCtrl(cell_toolbar_line2)
        self.cell_value_type_textctrl.SetEditable(False)
        cell_toolbar_line2_sizer.Add(self.cell_value_type_textctrl, 0, wx.CENTER)
        self.cell_value_textctrl = wx.TextCtrl(cell_toolbar_line2)
        self.cell_value_textctrl.SetEditable(False)
        cell_toolbar_line2_sizer.Add(self.cell_value_textctrl, 1, wx.CENTER)

        self.grid = spreadsheetgrid.SpreadSheetGrid(self, spreadsheet,
                                                [self.onCellSelect])

        box = wx.BoxSizer(wx.VERTICAL)
        box.Add((5,10) , 0)
        box.Add(cell_toolbar_line1, 0, wx.EXPAND, border=5)
        box.Add(cell_toolbar_line2, 0, wx.EXPAND, border=5)
        box.Add((5,10) , 0)
        box.Add(self.grid, 1, wx.EXPAND)
        self.SetSizer(box)

        self.cell_formula_textctrl.Bind(wx.EVT_KILL_FOCUS, self.onCellFormulaLostFocus)
        self.cell_formula_textctrl.Bind(wx.EVT_TEXT_ENTER, self.onCellFormulaLostFocus)

    def onCellSelect(self, position, cell):
        self.selected_cell_textctrl.SetValue(position)
        formula = ''
        value = ''
        type_name = ''
        if cell:
            formula = cell.getFormula()
            try:
                value = cell.getValue()
                type_name = type(value).__name__
                try:
                    type_name += ' of len %d' % len(value)
                except:
                    pass
            except (Exception, SyntaxError, SpreadSheetError), e:
                value = str(e)
                type_name = '(Error)'
        self.cell_formula_textctrl.SetValue(formula)
        self.cell_value_textctrl.SetValue(str(value))
        self.cell_value_type_textctrl.SetValue(type_name)

    def onCellFormulaLostFocus(self, event):
        value = event.GetEventObject().GetValue()
        row = self.grid.GetGridCursorRow()
        col = self.grid.GetGridCursorCol()
        
        old_value = self.grid.GetTable().GetFormula(row, col)
        if value != old_value:
            self.grid.GetTable().SetFormula(row, col, value)
            self.grid.ForceRefresh()
        event.Skip()


class GraphImage(wx.Window):
    def __init__(self, parent):
        image = wx.Image('graph.gif', wx.BITMAP_TYPE_GIF)
        image = image.ConvertToBitmap()
        wx.Window.__init__(self, parent, wx.ID_ANY)
        self.bmp = wx.StaticBitmap(parent=self, bitmap=image)

class PysApplicationWindow(wx.Frame):

    def __init__(self, parent, id, title, size):
        wx.Frame.__init__(self, parent, id, title, size=size)

        self._filename = None
        self._dirname = ''

        self._additional_paths = []
        self._conf_filename = conf_file_name('pessheet')
        self.loadSettings()
        import sys
        print 'Original path: ', sys.path
        print 'Additional paths: ', self._additional_paths

        self._spreadsheet = SpreadSheet(self._additional_paths)

        ib = wx.IconBundle()
        ib.AddIconFromFile("pessheet.ico", wx.BITMAP_TYPE_ANY)
        self.SetIcons(ib)

        self.id_export = wx.NewId()
        self.SetMenuBar(self.getMenuBar())

        panel = wx.Panel(self)

        main_toolbar = wx.ToolBar(panel, wx.ID_ANY, style=wx.TB_HORIZONTAL)
        main_toolbar.SetToolBitmapSize(iconbitmapsize()) # Needed for Windows XP
        alt = main_toolbar.AddLabelTool
        alt(wx.ID_NEW, '', iconbitmap('document-new.png'), shortHelp='New')
        alt(wx.ID_OPEN, '', iconbitmap('document-open.png'), shortHelp='Open')
        alt(wx.ID_SAVE, '', iconbitmap('document-save.png'), shortHelp='Save')
        alt(wx.ID_SAVEAS, '', iconbitmap('document-save-as.png'), shortHelp='Save as')
        main_toolbar.AddSeparator()
        alt(wx.ID_CUT, '', iconbitmap('edit-cut.png'), shortHelp='Cut')
        alt(wx.ID_COPY, '', iconbitmap('edit-copy.png'), shortHelp='Copy')
        alt(wx.ID_PASTE, '', iconbitmap('edit-paste.png'), shortHelp='Paste')
        #alt(-1, '',  iconbitmap('edit-delete.png'), shortHelp='Delete')
        main_toolbar.AddSeparator()
        alt(wx.ID_UNDO, '', iconbitmap('edit-undo.png'), shortHelp='Undo')
        alt(wx.ID_REDO, '', iconbitmap('edit-redo.png'), shortHelp='Redo')
        main_toolbar.AddSeparator()
        alt(wx.ID_EXIT, '',  iconbitmap('system-log-out.png'), shortHelp='Exit')

        main_toolbar.Realize()

        box = wx.BoxSizer(wx.VERTICAL)
        box.Add(main_toolbar, 0, wx.EXPAND, border=5)
        box.Add((5,5) , 0)

        panel.SetSizer(box)
        notebook = wx.Notebook(panel, wx.ID_ANY, style=wx.LEFT)
        self._sheet_page = SheetPage(notebook, self._spreadsheet)
        #graph_page = GraphImage(notebook)
        script_page = wx.py.editor.EditWindow(None, notebook)
        self._editor = script_page

        notebook.AddPage(self._sheet_page, 'Sheet')
        #notebook.AddPage(graph_page, 'Graph')
        notebook.AddPage(script_page, 'Script')

        box.Add(notebook, 1, wx.EXPAND)

        self.Bind(wx.EVT_MENU, self.OnQuit, id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.OnFileNew, id=wx.ID_NEW)
        self.Bind(wx.EVT_MENU, self.OnFileOpen, id=wx.ID_OPEN)
        self.Bind(wx.EVT_MENU, self.OnFileSave, id=wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, self.OnFileSaveAs, id=wx.ID_SAVEAS)
        self.Bind(wx.EVT_MENU, self.OnFileExport, id=self.id_export)
        self.Bind(wx.EVT_MENU, self.OnEditCut, id=wx.ID_CUT)
        self.Bind(wx.EVT_MENU, self.OnEditCopy, id=wx.ID_COPY)
        self.Bind(wx.EVT_MENU, self.OnEditPaste, id=wx.ID_PASTE)
        self.Bind(wx.EVT_MENU, self.OnOptions, id=wx.ID_PREFERENCES)
        
        script_page.Bind(wx.EVT_KILL_FOCUS, self.onEditorLostFocus)

        self.updateTitle()
        self.CreateStatusBar()
        self.Centre()
        self.Show(True)
        self._sheet_page.grid.SetFocus()

    def loadSettings(self):
        import cPickle
        try:
            settings_dict = cPickle.load(open(self._conf_filename, 'r'))
            self._additional_paths[:] = settings_dict.get('additional_paths', 
                                                          [])[:]
        except IOError, e:
            pass

    def saveSettings(self):
        import cPickle
        settings_dict = {
                'additional_paths': self._additional_paths,
        }
        cPickle.dump(settings_dict, open(self._conf_filename, 'w'))

    def getMenuBar(self):

        menues = [
            ('&File', 
             [
              (wx.ID_NEW, '&New...\tCtrl-N', 'Create new spreadsheet'), 
              (wx.ID_OPEN, '&Open...\tCtrl-O', 'Open a spreadsheet'), 
              (wx.ID_SAVE, '&Save\tCtrl-S', 'Save current spreadsheet'),
              (wx.ID_SAVEAS, 'Save As...\tCtrl-Shift-S', 'Save current spreadsheet as new file'),
              (self.id_export, 'Export...\tCtrl-E', 'Export current spreadsheet into other format'),
              None, #Separator
              (wx.ID_ABOUT, '&About pyssheet', 'Information about this program'),
              (wx.ID_HELP, '&Help', 'Help about using this program'),
              None,
              (wx.ID_EXIT, 'E&xit\tCtrl-Q', 'Terminate the program'),
             ]
            ),
            ('&Edit', 
             [
              (wx.ID_UNDO, 'Undo\tCtrl-Z', 'Undo'), 
              (wx.ID_REDO, 'Redo\tCtrl-Shift-Z', 'Redo'),
              None, 
              (wx.ID_CUT, 'Cut\tCtrl-X', 'Cut selection into Clipboard'), 
              (wx.ID_COPY, 'Copy\tCtrl-C', 'Copy selection into Clipboard'),
              (wx.ID_PASTE, 'Paste\tCtrl-V', 'Paste the content of the Clipboard into the sheet'),
             ]
            ),
            ('&View', 
             [
              (self.onSheet, 'Sheet', 'Show Sheet'), 
              #(wx.ID_ANY, 'Graph', 'Show Calculation Graph'),
              (wx.ID_ANY, 'Script', 'Show Script'),
             ]
            ),
            ('&Tools', 
             [
              (wx.ID_PREFERENCES, '&Options...', 'Set options'), 
             ]
            ),
        ]

        menuBar = wx.MenuBar()
        for menu_name, menu_items in menues:
            menu = wx.Menu()
            menuBar.Append(menu, menu_name)
            for menu_item in menu_items:
                if menu_item:
                    id, name, help = menu_item
                    if isinstance(id, int):
                        menu.Append(id, name, help)
                    else:
                        callback = id
                        assigned_id = menu.Append(wx.ID_ANY, name, help)
                        self.Bind(wx.EVT_MENU, callback, assigned_id)
                else:
                    menu.AppendSeparator()
                    
        return menuBar


    def onSheet(self, event):
        print 'OnSheet called'

    def onEditorLostFocus(self, event):
        script = self._editor.GetText()
        script = '\n'.join(script.splitlines())
        self._spreadsheet.setScript(script)

    def OnQuit(self, event):
        self.Close()

    def getApplicationName(self, version=False):
        return 'PESsheet' + (' v0.03' if version else '')

    def updateTitle(self):
        if self._filename:
            self.SetTitle('%s (%s) - %s' % (self._filename, self._dirname, self.getApplicationName(True)))
        else:
            self.SetTitle('New spreadsheet - %s' % self.getApplicationName(True))

    def OnFileNew(self, event):
        self._filename = None
        self._spreadsheet.clear()
        self.Refresh()
        self.SetStatusText('New spreadsheet created')
        self.updateTitle()

    def OnFileOpen(self, event):
        filedialog = wx.FileDialog(self, "Open spreadsheet file", 
                                   self._dirname, '', 
                                   "%s files (*.pss)|*.pss" % self.getApplicationName(), 
                                   wx.OPEN)
        if filedialog.ShowModal() == wx.ID_OK:
            self._filename = filedialog.GetFilename()
            self._dirname = filedialog.GetDirectory()
            self.load()
        filedialog.Destroy()
        self.Refresh()
      
    def load(self):
        import os
        self.SetStatusText('Loading file %s' % self._filename)
        f = open(os.path.join(self._dirname, self._filename), 'r')
        contents = f.read()
        contents = '\n'.join(contents.splitlines())
        f.close()
        try:
            self._spreadsheet.load(contents)
        except (ImportError, SyntaxError), e:
            print e

        self._editor.SetText(self._spreadsheet.getScript())
        self.SetStatusText('File loaded')
        self.updateTitle()

    def save(self):
        import os
        self.SetStatusText('Saving file %s' % self._filename)
        f = open(os.path.join(self._dirname, self._filename), 'w')
        f.write(self._spreadsheet.save())
        f.close()
        self.SetStatusText('File saved')
        self.updateTitle()

    def OnFileSave(self, event):
        if self._filename:
            self.save()
        else:
            self.OnFileSaveAs(event)

    def OnFileSaveAs(self, event):
        filedialog = wx.FileDialog(self, "Save spreadsheet as", 
                                   self._dirname, '', 
                                   "%s files (*.pss)|*.pss" % self.getApplicationName(), 
                                   wx.SAVE | wx.OVERWRITE_PROMPT)
        if filedialog.ShowModal() == wx.ID_OK:
            self._filename = filedialog.GetFilename()
            self._dirname = filedialog.GetDirectory()
            self.save()
        filedialog.Destroy()
        
    def OnFileExport(self, event):
        export_formats = [
	   ('Python script (*.py)', '.py', self.export_python), 
	   ('Graphviz dot file (*.dot)', '.dot', self.export_dot), 
        ]
        filter_string = '|'.join(d+'|*'+e for d, e, _ in export_formats)
        filedialog = wx.FileDialog(self, 'Export spreadsheet as', 
                                   self._dirname, '', 
				   filter_string, 
                                   wx.SAVE | wx.OVERWRITE_PROMPT)
        if filedialog.ShowModal() == wx.ID_OK:
            filter_index = filedialog.GetFilterIndex()
            _, extension, export_function = export_formats[filter_index]
            dirname = filedialog.GetDirectory()
            filename = filedialog.GetFilename()
            if not filename.endswith(extension):
                filename += extension
            export_function(dirname, filename)
        filedialog.Destroy()
        
    def export_python(self, dirname, filename):
        import os
        self.SetStatusText('Exporting to file %s' % filename)
        f = open(os.path.join(dirname, filename), 'w')
        f.write(self._spreadsheet.asScript())
        f.close()
        self.SetStatusText('Spreadsheet exported')

    def export_dot(self, dirname, filename):
        import os
        self.SetStatusText('Exporting to file %s' % filename)
        f = open(os.path.join(dirname, filename), 'w')
        f.write(self._spreadsheet.asDot())
        f.close()
        self.SetStatusText('Spreadsheet exported')
        #dialog = wx.MessageDialog(None, 'Exporting to dot format not yet implemented', 
			#'Not yet implemented', wx.OK)
        #dialog.ShowModal()
        #dialog.Destroy()

    def OnEditCut(self, event):
        #print 'PysApplicationWindow.OnEditCut', event
        clip_object = None
        if (wx.GetActiveWindow().FindFocus().GetParent() == 
                    self._sheet_page.grid):
                clip_object = self._sheet_page.grid.cut()
        if clip_object:
            text_data = wx.TextDataObject(clip_object)
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(text_data)
                wx.TheClipboard.Close()
        else:
            event.Skip()

    def OnEditCopy(self, event):
        #print 'PysApplicationWindow.OnEditCopy', event
        clip_object = None
        if (wx.GetActiveWindow().FindFocus().GetParent() == 
                    self._sheet_page.grid):
                clip_object = self._sheet_page.grid.copy()
        if clip_object:
            text_data = wx.TextDataObject(clip_object)
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(text_data)
                wx.TheClipboard.Close()
        else:
            event.Skip()

    def OnEditPaste(self, event):
        text_data = wx.TextDataObject()
        success = False
        if wx.TheClipboard.Open():
            success = wx.TheClipboard.GetData(text_data)
            wx.TheClipboard.Close()

        if success:
            if (wx.GetActiveWindow().FindFocus().GetParent() == 
                    self._sheet_page.grid):
                self._sheet_page.grid.paste(text_data.GetText())
            else:
                event.Skip()
        """
        Maybe some useful stuff:
        http://74.125.77.132/search?q=cache:l8ru47_-EuAJ:www.picalo.org/download/picalo-2.32/picalo/gui/Spreadsheet.py+ID_PASTE+wx+bind+TheClipBoard+grid&hl=sv&ct=clnk&cd=1&gl=se&client=firefox-a
        """

    def OnOptions(self, event):
        value = '\n'.join(self._additional_paths)
        d = wx.TextEntryDialog(self, 
                'Additional paths to python packages (one per line)', 
                'Additional paths', value, 
                style=wx.OK|wx.CANCEL|wx.TE_MULTILINE)
        if d.ShowModal() == wx.ID_OK:
            value = str(d.GetValue()).splitlines()
            self._additional_paths[:] = value[:]
            self.saveSettings()
      
        event.Skip()

if __name__ == '__main__':
    app = wx.App(redirect=False)

    bmp = wx.Image('pes_splash.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap()
    wx.SplashScreen(bmp, wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT, 
                    1000, None, -1)

    window_size = wx.GetDisplaySize()
    window_size.Scale(0.8, 0.8)

    main_window = PysApplicationWindow(None, wx.ID_ANY, 
                                       'Python Scriptable SpreadSheet', 
                                       window_size)
    app.MainLoop()

