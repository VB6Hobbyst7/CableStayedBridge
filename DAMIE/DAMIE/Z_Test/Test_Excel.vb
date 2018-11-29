Imports Microsoft.Office.Interop.Excel

Class Test_Excel
    Inherits System.Windows.Forms.Form
    Public WithEvents App As Application
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public wkbk As Workbook

    Private Sub Test_Excel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        App = New Application
        App.Visible = True
        App.WindowState = XlWindowState.xlMaximized
        Me.wkbk = App.Workbooks.Add
        Call test()
        Me.Visible = False
    End Sub
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(79, 48)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Test_Excel
        '
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Test_Excel"
        Me.ResumeLayout(False)

    End Sub


    Sub test()
        Dim sht As Worksheet
        sht = wkbk.Worksheets(1)
        Dim cht As Chart
        cht = sht.Shapes.AddChart(XlChartType.xlXYScatterSmoothNoMarkers).Chart

        Dim X(0 To 40) As Double
        Dim Y(0 To 40) As Double
        Dim j As Integer
        Dim i As Single
        For i = 0 To 4 Step 0.1
            X(j) = i
            Y(j) = (4 - (i - 2) ^ 2) ^ 0.5
            j = j + 1
        Next
        Dim s As Series
        s = cht.SeriesCollection.NewSeries()
        s.XValues = X
        s.Values = Y

        Dim ax1 As Axis
        ax1 = cht.Axes(XlAxisType.xlCategory)
        Dim ax2 As Axis
        ax2 = cht.Axes(XlAxisType.xlValue)
        Dim ptArea As PlotArea
        ptArea = cht.PlotArea
        With ax1
            .MinimumScale = 0
            .MaximumScale = 5
            .MajorUnit = 0.5
            .MinorUnit = 0.1
            .MajorGridlines.Delete()
        End With

        With ax2
            .MinimumScale = 0
            .MaximumScale = 2.5
            .MajorUnit = 0.5
            .MinorUnit = 0.1
            .MajorGridlines.Delete()
        End With


        With cht.ChartArea
            .Top = 0
            .Left = 0
            .Height = 225
            .Width = 425
        End With
        '
        With ptArea
            .InsideTop = 0
            .InsideLeft = 0
            .InsideHeight = 200
            .InsideWidth = 400
        End With

    End Sub
End Class

