import wx
import UI.imgDownloader as ui


if __name__ == '__main__':
    app = wx.App(False)
    frmMain = ui.mainFrame(None)
    frmMain.Show(True)

    app.MainLoop()
