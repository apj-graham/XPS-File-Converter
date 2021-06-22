import wx
from FileData import FileData
from xlsxwriter.exceptions import FileCreateError

#TODO sort out global files

#TODO sort put onProcessButtonClick (move to another file and break down)

global Files
global Processed
Files = []
Processed = []

class PopupMenu(wx.Menu):
    def __init__(self, parent):
        """
        Initialise the right click menu to rename or delete items from
        the lits of files to process.
        """
        super(PopupMenu, self).__init__()

        self.parent = parent

        rename_option = wx.MenuItem(self, wx.NewId(), 'Rename')
        self.AppendItem(rename_option)
        self.Bind(wx.EVT_MENU, self.OnRename, mmi)

        delete_option = wx.MenuItem(self, wx.NewId(), 'Delete')
        self.AppendItem(delete_option)
        self.Bind(wx.EVT_MENU, self.OnDelete, cmi)

    def OnRename(self, event):
        """
        Renames selected file by changing entry on global file list
        """
        sel = self.parent.list_of_files_to_process.GetSelection()
        old_name = self.parent.list_of_files_to_process.GetString(sel)
        new_name = wx.GetTextFromUser('Rename item', 'Rename dialog', old_name)
        if new_name != '':
            self.parent.list_of_files_to_process.Delete(sel)
            self.parent.list_of_files_to_process.Insert(new_name, sel)
            self._update_file_data(old_name, new_name)

    def _update_file_data(self, old_name, new_name):
        for file_data in Files:
            if file_data.filename == old_name:
                file_data.filename == new_name

    def OnDelete(self, event):
        """
        Deletes file from processing list and global file list
        """
        sel = self.parent.list_of_files_to_process.GetSelection()
        file_name = self.parent.list_of_files_to_process.GetString(sel)
        if sel != -1:
            self.parent.list_of_files_to_process.Delete(sel)
            self._remove_file_data_from_Files(file_name)

    def _remove_file_data_from_Files(self, file_name):
        for file_data in Files:
            if file_data.filename == file_name:
                Files.remove(file_data)

class FileDrop(wx.FileDropTarget):
    def __init__(self, window):
        """
        Initialises drop box
        """
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, file_paths):
        """
        Adds files to drag and drop box
        """
        for file_path in file_paths:
            try:
                file_data = FileData(file_path)
                self.window.Append(file_data.filename)
                Files.append(file_data)
                return True
            except IOError as _:
                self._showErrMessage("Error opening file")
                return False
            except UnicodeDecodeError as _:
                self._showErrMessage("Cannot open non ascii files")
                return False

    def _showErrMessage(self, err_msg):
        """
        Function to show error dialog box
        """
        dialogBox = wx.MessageDialog(None, err_msg +'\n' + str(err_msg))
        dialogBox.ShowModal()

class GraphicalUserInterface(wx.Frame):
    def __init__(self, parent, title):
        """
        Initialises main window of app
        """
        super(GraphicalUserInterface, self).__init__(parent, title=title,
            size=(600, 400))

        self.task_range = 0
        self.count = 0

        self.populate_ui()
        self.Centre()
        self.Show()

    def populate_ui(self):
        """
        Adds all buttons/titles/boxes etc to window
        """

        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(6, 2)

        panel_title = wx.StaticText(panel, label="DATA PROCESSOR", style = wx.ALIGN_CENTRE )
        sizer.Add(panel_title, pos=(0, 1),span = (1,1), flag=wx.EXPAND|wx.ALIGN_CENTRE, border=5)

        drag_n_drop_header = wx.StaticText(panel, label="Files To Process", style = wx.ALIGN_CENTRE )
        sizer.Add(drag_n_drop_header, pos=(1, 0),span = (1,1), flag=wx.EXPAND|wx.ALIGN_CENTRE, border=5)

        processed_files_list_header = wx.StaticText(panel, label="Files Processed", style = wx.ALIGN_CENTRE )
        sizer.Add(processed_files_list_header, pos=(1, 2),span = (1,1), flag=wx.EXPAND|wx.ALIGN_CENTRE, border=5)

        self.processing_indicator = wx.StaticText(panel, label="", style = wx.ALIGN_CENTRE)
        sizer.Add(self.processing_indicator, pos=(4, 1),span = (1,1), flag=wx.EXPAND|wx.ALIGN_CENTRE, border=5)

        self.list_of_files_to_process = wx.ListBox(panel,-1)
        self.list_of_files_to_process.Bind(wx.EVT_RIGHT_DOWN, self.OnRightClick)
        drop_target = FileDrop(self.list_of_files_to_process)
        self.list_of_files_to_process.SetDropTarget(drop_target)
        sizer.Add(self.list_of_files_to_process, pos=(2, 0), span=(4, 1),
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        self.list_of_processed_files = wx.ListBox(panel,-1)
        sizer.Add(self.list_of_processed_files, pos=(2, 2), span=(4, 1),
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)

        process_button = wx.Button(panel, label="Process -->", size=(190, 50))
        process_button.Bind(wx.EVT_BUTTON, self.OnProcessButtonClick)
        sizer.Add(process_button, pos=(2, 1), flag=wx.ALIGN_CENTER, border=5)

        close_button = wx.Button(panel, label="Close", size=(90, 25))
        close_button.Bind(wx.EVT_BUTTON, self.OnClose)
        sizer.Add(close_button, pos=(6, 2), flag=wx.RIGHT|wx.BOTTOM, border=5)

        self.progress_bar = wx.Gauge(panel, range= self.task_range, size=(260, 25))
        sizer.Add(self.progress_bar, pos=(3, 1), flag=wx.ALIGN_CENTER, border=5)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(2)
        panel.SetSizerAndFit(sizer)

    def OnProcessButtonClick(self, e):
        """
        On process button click runs through all files in the list drag and drop box and
        calls the process to convert the file on each one.
        """
        self.count = 0
        self.task_range = len(Files)
        self.progress_bar.SetRange(self.task_range)
        self.progress_bar.SetValue(self.count) #reset progress bar
        self.processing_indicator.SetLabel('Task in Progress')
        if self.task_range != 0:
            for file_data in Files:
                try:
                    if file_data.filename not in Processed:
                        file_data.save_data_as_xlsx()
                        self.list_of_processed_files.Append(file_data.filename)
                        Processed.append(file_data)

                        self.count += 1
                        self.progress_bar.SetValue(self.count)

                        if self.count == self.task_range:
                            self.processing_indicator.SetLabel('Task Completed')
                except FileCreateError as e:
                    self.processing_indicator.SetLabel(f'An excel file with name: {file_data.filename} already exists\nPlease delete that file and try again')

            self.list_of_files_to_process.Clear()
            while len(Files) != 0:
                Files.pop()
        else:
            self.processing_indicator.SetLabel('No Files Selected')

    def OnRightClick(self, e):
        """
        Gets drop down menu at location of click
        """
        self.PopupMenu(PopupMenu(self), e.GetPosition())

    def OnClose(self, e):
        """
        Closes window
        """
        self.Close()
