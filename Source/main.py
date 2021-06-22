import wx
from GUI import GraphicalUserInterface

if __name__ == '__main__':
    app = wx.App()
    GraphicalUserInterface(None, title='EIS Data Processor')
    app.MainLoop()
