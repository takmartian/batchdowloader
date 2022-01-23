# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version 3.10.1-0-g8feb16b)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import time
import pandas as pd
import requests
from threading import Thread
from pubsub import pub


###########################################################################
## Class mainFrame
###########################################################################

class mainFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"Excel图片地址批量下载器", pos=wx.DefaultPosition,
                          size=wx.Size(500, 306), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.progress_value = 0
        self.progress_range = 0

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        fgSizer1 = wx.FlexGridSizer(0, 3, 0, 0)
        fgSizer1.AddGrowableCol(1)
        fgSizer1.SetFlexibleDirection(wx.BOTH)
        fgSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        fgSizer1.Add((0, 10), 1, wx.EXPAND, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"选择Excel文件：", wx.DefaultPosition, wx.DefaultSize,
                                           wx.ALIGN_RIGHT)
        self.m_staticText1.Wrap(-1)

        fgSizer1.Add(self.m_staticText1, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.txtExcelPath = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                        wx.TE_READONLY)
        fgSizer1.Add(self.txtExcelPath, 0, wx.ALL | wx.EXPAND, 5)

        self.btnBrowseXlsFile = wx.Button(self, wx.ID_ANY, u"浏览...", wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.btnBrowseXlsFile, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.m_staticText2 = wx.StaticText(self, wx.ID_ANY, u"选择工作表：", wx.DefaultPosition, wx.DefaultSize,
                                           wx.ALIGN_RIGHT)
        self.m_staticText2.Wrap(-1)

        fgSizer1.Add(self.m_staticText2, 0, wx.ALL | wx.EXPAND, 5)

        dropSelSheetChoices = []
        self.dropSelSheet = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, dropSelSheetChoices, 0)
        self.dropSelSheet.SetSelection(0)
        fgSizer1.Add(self.dropSelSheet, 0, wx.ALL | wx.EXPAND, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"选择下载地址列：", wx.DefaultPosition, wx.DefaultSize,
                                           wx.ALIGN_RIGHT)
        self.m_staticText3.Wrap(-1)

        fgSizer1.Add(self.m_staticText3, 0, wx.ALL | wx.EXPAND, 5)

        dropSelColChoices = []
        self.dropSelCol = wx.ComboBox(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                      dropSelColChoices, 0)
        fgSizer1.Add(self.dropSelCol, 0, wx.ALL | wx.EXPAND, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.m_staticText6 = wx.StaticText(self, wx.ID_ANY, u"姓名列：", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT)
        self.m_staticText6.Wrap(-1)

        fgSizer1.Add(self.m_staticText6, 0, wx.ALL | wx.EXPAND, 5)

        dropNameColChoices = []
        self.dropNameCol = wx.ComboBox(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                       dropNameColChoices, 0)
        fgSizer1.Add(self.dropNameCol, 0, wx.ALL | wx.EXPAND, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"下载文件夹：", wx.DefaultPosition, wx.DefaultSize,
                                           wx.ALIGN_RIGHT)
        self.m_staticText4.Wrap(-1)

        fgSizer1.Add(self.m_staticText4, 0, wx.ALL | wx.EXPAND, 5)

        self.txtDLFolder = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                       wx.TE_READONLY)
        fgSizer1.Add(self.txtDLFolder, 0, wx.ALL | wx.EXPAND, 5)

        self.btnBrowseDLFolder = wx.Button(self, wx.ID_ANY, u"浏览...", wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.btnBrowseDLFolder, 0, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 10)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.progressBar = wx.Gauge(self, wx.ID_ANY, 100, wx.DefaultPosition, wx.DefaultSize, wx.GA_HORIZONTAL)
        self.progressBar.SetValue(0)
        fgSizer1.Add(self.progressBar, 0, wx.ALL | wx.EXPAND, 5)

        self.dlCount = wx.StaticText(self, wx.ID_ANY, u"0/0", wx.DefaultPosition, wx.DefaultSize, 0)
        self.dlCount.Wrap(-1)

        fgSizer1.Add(self.dlCount, 0, wx.ALL, 5)

        fgSizer1.Add((0, 0), 1, wx.EXPAND, 5)

        self.btnStart = wx.Button(self, wx.ID_ANY, u"开始下载喽！", wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer1.Add(self.btnStart, 0, wx.ALL, 5)

        self.SetSizer(fgSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btnBrowseXlsFile.Bind(wx.EVT_BUTTON, self.Browse_xls)
        self.dropSelSheet.Bind(wx.EVT_CHOICE, self.refreshCol)
        self.btnBrowseDLFolder.Bind(wx.EVT_BUTTON, self.Browse_Folder)
        self.btnStart.Bind(wx.EVT_BUTTON, self.Start_Download)

        self.dlCount.Bind(wx.EVT_IDLE, self.LblProgressOnIdle)
        self.progressBar.Bind(wx.EVT_IDLE, self.ProgressBarOnIdle)

        pub.subscribe(self.set_progress, "updateProgress")

    def __del__(self):
        pass

    # Virtual event handlers, override them in your derived class
    def Browse_xls(self, event):
        filterXls = "表格文件(*.xls, *.xlsx, *.csv)|*.xls;*.xlsx;*.csv"
        fileDialog = wx.FileDialog(self, message="选择表格文件", wildcard=filterXls, style=wx.FD_OPEN)
        dialogResult = fileDialog.ShowModal()
        if dialogResult != wx.ID_OK:
            return

        excelPath = fileDialog.GetPath()
        self.txtExcelPath.SetValue(excelPath)

        self.dropSelSheet.Clear()
        if excelPath[-4:] == "xlsx" or excelPath[-4:] == ".xls":
            # 读取excel表名
            xls_sheets = pd.ExcelFile(excelPath).sheet_names
            for sheetsList in xls_sheets:
                self.dropSelSheet.AppendItems(sheetsList)

            xls = pd.read_excel(excelPath)
        elif excelPath[-4:] == ".csv":
            # 如果是csv
            self.dropSelSheet.AppendItems("无")
            xls = pd.read_csv(excelPath)

        # 读取列名
        self.dropSelCol.Clear()
        for colList in xls.columns.values:
            self.dropSelCol.AppendItems(colList)
            self.dropNameCol.AppendItems(colList)
        fileDialog.Destroy()

        self.dropSelCol.Clear()
        for colList in xls.columns.values:
            self.dropSelCol.AppendItems(colList)
        fileDialog.Destroy()


    def refreshCol(self, event):
        excelPath = self.txtExcelPath.GetValue()
        if excelPath[-4:] == "xlsx" or excelPath[-4:] == ".xls":
            excelPath = self.txtExcelPath.GetValue()
            xls = pd.read_excel(excelPath, sheet_name=self.dropSelSheet.GetStringSelection())
            self.dropSelCol.Clear()
            for colList in xls.columns.values:
                self.dropSelCol.AppendItems(colList)
                self.dropNameCol.AppendItems(colList)


    def Browse_Folder(self, event):
        dlg = wx.DirDialog(self, u"选择保存图片的文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            savePath = dlg.GetPath()
            self.txtDLFolder.SetValue(savePath)
        dlg.Destroy()


    def Start_Download(self, event):
        xlsPath = self.txtExcelPath.GetValue()
        xlsSheet = self.dropSelSheet.GetStringSelection()
        xlsCol = self.dropSelCol.GetStringSelection()
        xlsNameCol = self.dropNameCol.GetStringSelection()
        dlPath = self.txtDLFolder.GetValue()
        if dlPath[-1:] != "/":
            dlPath = dlPath + "/"

        if dlPath == "":
            wx.MessageBox("请选择需要下载的文件夹。", "错误")
            return
        elif xlsNameCol == "":
            wx.MessageBox("请选择表格中的姓名列用于保存文件名。", "错误")
            return
        elif xlsPath == "":
            wx.MessageBox("请选择Excel文件。", "错误")
            return
        elif self.dropSelCol.GetStringSelection() == "":
            wx.MessageBox("请选择列。", "错误")
            return

        DownloadImages(xlsPath, xlsSheet, xlsNameCol, dlPath, xlsCol)

    def set_progress(self, p_value, p_range):
        self.progress_value = p_value
        self.progress_range = p_range

    def LblProgressOnIdle(self, event):
        self.dlCount.SetLabel("%d/%d" % (self.progress_value, self.progress_range))

    def ProgressBarOnIdle(self, event):
        self.progressBar.SetRange(self.progress_range)
        self.progressBar.SetValue(self.progress_value)




###########################################################################
# Threads
###########################################################################
class DownloadImages(Thread):
    def __init__(self, xlsPath, xlsSheet, xlsNameCol, dlPath, xlsCol):
        self.xlsPath = xlsPath
        self.xlsSheet = xlsSheet
        self.xlsNameCol = xlsNameCol
        self.dlPath = dlPath
        self.xlsCol = xlsCol
        Thread.__init__(self)
        self.start()

    def run(self):
        if self.xlsPath[-4:] == "xlsx" or self.xlsPath[-4:] == ".xls":
            df = pd.read_excel(self.xlsPath, sheet_name=self.xlsSheet, keep_default_na=False)
        elif self.xlsPath[-4:] == ".csv":
            df = pd.read_csv(self.xlsPath, keep_default_na=False)

        urls = df[self.xlsCol]
        names = df[self.xlsNameCol]
        curDLNum = 0
        failDLNum = 0
        failNames = ""

        for i in range(len(urls)):
            try:
                if urls[i] == "":
                    continue
                urlSplited = urls[i].split(",")
                for j in urlSplited:
                        dotIndex = j.rfind(".")
                        extName = j[dotIndex:len(j)]
                        r = requests.request('get', j)
                        with open(self.dlPath + names[i] + str('img') + str(time.time()) + extName, 'wb') as f:
                            f.write(r.content)
                        f.close()
                curDLNum += 1
                pub.sendMessage("updateProgress", p_value=curDLNum, p_range=len(urls))
            except:
                failDLNum += 1
                failNames = "%s, %s" % (failNames, names[i])
                continue

        wx.MessageBox("下载完毕，成功%d条，失败%d条。\n下列是失败清单：\n%s。" % (curDLNum, failDLNum, failNames))
