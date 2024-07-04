import wx
import wx.lib.scrolledpanel as scrolled
import time
import threading
from comparison import compare_excel_files
import os

class ExcelCompareFrame(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title="Emil's Excel Comparer", size=(600, 500))
        panel = wx.Panel(self)
        
        bg_color = wx.Colour(30, 30, 30)
        text_color = wx.Colour(220, 220, 220)
        button_color = wx.Colour(60, 60, 60)
        button_text_color = wx.Colour(220, 220, 220)
        
        panel.SetBackgroundColour(bg_color)
        
        title = wx.StaticText(panel, label="Emil's Excel Comparer")
        font = title.GetFont()
        font.SetPointSize(20)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        title.SetFont(font)
        title.SetForegroundColour(text_color)
        
        self.file1_label = wx.StaticText(panel, label="No file selected")
        self.file1_label.SetForegroundColour(text_color)
        file1_btn = wx.Button(panel, label="File 1")
        file1_btn.SetBackgroundColour(button_color)
        file1_btn.SetForegroundColour(button_text_color)
        file1_btn.Bind(wx.EVT_BUTTON, lambda event: self.browse_file(1))
        
        self.file2_label = wx.StaticText(panel, label="No file selected")
        self.file2_label.SetForegroundColour(text_color)
        file2_btn = wx.Button(panel, label="File 2")
        file2_btn.SetBackgroundColour(button_color)
        file2_btn.SetForegroundColour(button_text_color)
        file2_btn.Bind(wx.EVT_BUTTON, lambda event: self.browse_file(2))
        
        self.compare_btn = wx.Button(panel, label="Compare")
        self.compare_btn.SetBackgroundColour(button_color)
        self.compare_btn.SetForegroundColour(button_text_color)
        self.compare_btn.Bind(wx.EVT_BUTTON, self.compare_files)
        
        self.gauge = wx.Gauge(panel, range=100, size=(300, 25))
        self.gauge.SetValue(0)
        
        self.timer_label = wx.StaticText(panel, label="Time: 0.00 seconds")
        self.timer_label.SetForegroundColour(text_color)
        
        self.output_panel = scrolled.ScrolledPanel(panel)
        self.output_panel.SetBackgroundColour(wx.Colour(40, 40, 40))
        self.output_ctrl = wx.TextCtrl(self.output_panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.output_ctrl.SetBackgroundColour(wx.Colour(40, 40, 40))
        self.output_ctrl.SetForegroundColour(text_color)
        
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(title, 0, wx.ALL | wx.CENTER, 20)
        
        file_sizer = wx.FlexGridSizer(2, 2, 10, 10)
        file_sizer.Add(file1_btn, 0)
        file_sizer.Add(self.file1_label, 0, wx.ALIGN_CENTER_VERTICAL)
        file_sizer.Add(file2_btn, 0)
        file_sizer.Add(self.file2_label, 0, wx.ALIGN_CENTER_VERTICAL)
        
        main_sizer.Add(file_sizer, 0, wx.EXPAND | wx.ALL, 20)
        main_sizer.Add(self.compare_btn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 20)
        main_sizer.Add(self.gauge, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        main_sizer.Add(self.timer_label, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        
        output_sizer = wx.BoxSizer(wx.VERTICAL)
        output_sizer.Add(self.output_ctrl, 1, wx.EXPAND)
        self.output_panel.SetSizer(output_sizer)
        self.output_panel.SetupScrolling()
        
        main_sizer.Add(self.output_panel, 1, wx.EXPAND | wx.ALL, 20)
        
        panel.SetSizer(main_sizer)
        
        self.Centre()
        self.Show()
        
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update_timer, self.timer)
    
    def browse_file(self, file_num):
        with wx.FileDialog(self, "Choose excel-file", wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            pathname = fileDialog.GetPath()
            filename = os.path.basename(pathname)
            if file_num == 1:
                self.file1_label.SetLabel(filename)
                self.file1_path = pathname
            else:
                self.file2_label.SetLabel(filename)
                self.file2_path = pathname
    
    def compare_files(self, event):
        if hasattr(self, 'file1_path') and hasattr(self, 'file2_path'):
            self.compare_btn.Disable()
            self.gauge.SetValue(0)
            self.timer_label.SetLabel("Time: 0.00 seconds")
            
            self.start_time = time.time()
            self.timer.Start(100)
            
            threading.Thread(target=self.run_comparison, daemon=True).start()
        else:
            self.output_ctrl.SetValue("Select both files before comparing.")
    
    def run_comparison(self):
        result = compare_excel_files(self.file1_path, self.file2_path)
        wx.CallAfter(self.comparison_done, result)
    
    def comparison_done(self, result):
        self.timer.Stop()
        elapsed_time = time.time() - self.start_time
        self.timer_label.SetLabel(f"Time: {elapsed_time:.2f} seconds")
        self.gauge.SetValue(100)
        
        self.output_ctrl.SetValue(result)
        self.compare_btn.Enable()
    
    def update_timer(self, event):
        elapsed_time = time.time() - self.start_time
        self.timer_label.SetLabel(f"Time: {elapsed_time:.2f} seconds")
        self.gauge.Pulse()