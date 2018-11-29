Imports eZstd.eZAPI.APIWindows
Imports eZstd.eZAPI
Imports Microsoft.Office.Interop.Excel
Public Class Test_FixWndSize
    Private WithEvents ExcelApp As Application
    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelApp = New Application
        ExcelApp.Visible = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        NoResizeExcel(ExcelApp)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ResizeExcel(ExcelApp)
    End Sub
    ''' <summary>
    ''' 禁止Excel程序窗口的缩放
    ''' </summary>
    ''' <param name="App"></param>
    ''' <remarks></remarks>
    Sub NoResizeExcel(ByVal App As Application)
        App.ScreenUpdating = False
        App.WindowState = XlWindowState.xlNormal
        With App
            .Width = 400
            .Height = 453
        End With
        App.ScreenUpdating = True
        'Dim hWnd As IntPtr = FindWindow("XLMAIN", App.Caption)
        Dim hWnd As IntPtr = App.Hwnd
        FixWindow(hWnd)
    End Sub

    ''' <summary>
    ''' 允许Excel程序窗口的缩放
    ''' </summary>
    ''' <param name="App"></param>
    ''' <remarks></remarks>
    Sub ResizeExcel(ByVal App As Application)
        App.WindowState = XlWindowState.xlNormal
        'Dim hWnd As IntPtr = FindWindow("XLMAIN", App.Caption)
        Dim hWnd As IntPtr = App.Hwnd
        unFixWindow(hWnd)
    End Sub

    ''' <summary>
    ''' 通过释放窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
    ''' </summary>
    ''' <param name="hWnd">要释放大小的窗口的句柄</param>
    ''' <remarks></remarks>
    Private Sub FixWindow(ByVal hWnd As IntPtr)
        Dim hStyle As Integer = GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE)
        '禁用最大化的标头及拖拽视窗
        SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle And Not WindowStyle.WS_MAXIMIZEBOX And Not WindowStyle.WS_EX_APPWINDOW)
    End Sub

    ''' <summary>
    ''' 通过禁用窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
    ''' </summary>
    ''' <param name="hWnd">要固定大小的窗口的句柄</param>
    ''' <remarks></remarks>
    Private Sub unFixWindow(ByVal hWnd As IntPtr)
        Dim hStyle As Integer = GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE)
        SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle Or WindowStyle.WS_MAXIMIZEBOX Or WindowStyle.WS_EX_APPWINDOW)
    End Sub

End Class
