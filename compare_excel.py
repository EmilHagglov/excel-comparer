import wx
from gui import ExcelCompareFrame
from utils import setup_logging

if __name__ == '__main__':
    app = wx.App()
    frame = ExcelCompareFrame()
    app.MainLoop()