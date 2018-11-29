Imports Microsoft.Office.Interop
Imports DAMIE.Miscellaneous
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form
Imports eZstd.eZAPI.APIWindows

Public Class Diafrm_Output_Word

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

    ''' <summary>
    ''' 所有选择的要进行输出的图形
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure Drawings_For_Output
        Public PlanView As ClsDrawing_PlanView
        Public SectionalView As ClsDrawing_ExcavationElevation
        Public MntDrawings As List(Of ClsDrawing_Mnt_Base)

        Public Function Count() As UShort
            Dim SumUp As Short = MntDrawings.Count
            If PlanView IsNot Nothing Then
                SumUp += 1
            End If
            If SectionalView IsNot Nothing Then
                SumUp += 1
            End If
            Return SumUp
        End Function

        Public Sub New(Sender As Diafrm_Output_Word)
            MntDrawings = New List(Of ClsDrawing_Mnt_Base)
        End Sub
    End Structure

#End Region

#Region "  ---  Field定义"

    'word页面中正文区域的宽度，用来限定图片的宽度
    Private ContentWidth As Single

    '“当天”的日期值
    Private dateThisday As Date

    ''' <summary>
    ''' 窗口中的所有列表框listbox对象
    ''' </summary>
    ''' <remarks>此数组是为了便于后面的统一操作：清空内容、全部选择，取消全选</remarks>
    Private F_arrListBoxes(0 To 1) As ListBox

    ''' <summary>
    ''' 所有选择的要进行输出的图形
    ''' </summary>
    ''' <remarks></remarks>
    Private F_SelectedDrawings As Drawings_For_Output

    ''' <summary>
    ''' 程序中所有图表的窗口的句柄值，用来对窗口进行禁用或者启用窗口
    ''' </summary>
    ''' <remarks></remarks>
    Private WindowHandles As IntPtr()
#End Region

#Region "  ---  Properties"
    ''' <summary>
    ''' Word的Application对象
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents WdApp As Word.Application
    ''' <summary>
    ''' Word的Application对象
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Application As Word.Application
        Get
            Return WdApp
        End Get
    End Property

    Private WdDoc As Word.Document
    ''' <summary>
    ''' 输出到的word.document对象
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Document As Word.Document
        Get
            Return WdDoc
        End Get
    End Property


#End Region

#End Region

#Region "  ---  窗口的加载与关闭"

    ''' <summary>
    ''' Showdialog式加载
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frm_Output_Word_Load(sender As Object, e As EventArgs) Handles Me.Load
        '刷新时间
        dateThisday = APPLICATION_MAINFORM.MainForm.Form_Rolling.Rollingday
        '设置初始界面
        LabelDate.Text = dateThisday.ToString("yyyy/MM/dd")
        ChkBxSelect.CheckState = CheckState.Unchecked
        CheckBox_PlanView.Checked = False
        CheckBox_SectionalView.Checked = False
        btnExport.Enabled = False
        '为数组中的每一个元素赋值，以便于后面的统一操作：清空内容、全部选择，取消全选
        F_arrListBoxes(0) = ListBoxMonitor_Dynamic
        F_arrListBoxes(1) = ListBoxMonitor_Static
        '    
        F_SelectedDrawings = New Drawings_For_Output(Me)
        '刷新主程序与界面
        Dim AllDrawing As AmeDrawings = GlobalApplication.Application.ExposeAllDrawings
        ' ---------- 禁用所有绘图窗口
        Me.WindowHandles = GlobalApplication.GetWindwosHandles(AllDrawing)
        For Each H As IntPtr In WindowHandles
            eZstd.eZAPI.APIWindows.EnableWindow(H, False)
        Next
        '
        Call RefreshUI(AllDrawing)
    End Sub

    ''' <summary>
    ''' 从主程序中提取所有的图表，以进行输出之用。同时刷新窗口中的可供输出的图形
    ''' </summary>
    ''' <param name="AllDrawing">主程序对象中所有的图表</param>
    ''' <remarks>在此方法中，将提取主程序中的所有图表对象，而且将其显示在输入窗口的列表框中</remarks>
    Private Sub RefreshUI(ByVal AllDrawing As AmeDrawings)

        '-------------------------1、剖面图---------------------------------------------
        With CheckBox_SectionalView
            Dim Sectional As ClsDrawing_ExcavationElevation = AllDrawing.SectionalView
            .Tag = Sectional
            If Sectional IsNot Nothing Then
                .Enabled = True
            Else
                .Checked = False
                .Enabled = False
            End If
        End With
        '--------------------------2、开挖平面图----------------------------------------------
        With CheckBox_PlanView
            Dim Plan As ClsDrawing_PlanView = AllDrawing.PlanView
            .Tag = Plan
            If Plan IsNot Nothing Then
                .Enabled = True
            Else
                .Checked = False
                .Enabled = False
            End If
        End With

        '---------------------------------3、监测曲线图---------------------------------------

        '清空两个列表框中的所有项目
        For Each lstbox As ListBox In Me.F_arrListBoxes
            With lstbox
                .Items.Clear()
                .DisplayMember = LstbxDisplayAndItem.DisplayMember
            End With
        Next
        For Each sht As ClsDrawing_Mnt_Base In AllDrawing.MonitorData
            Select Case sht.Type
                Case DrawingType.Monitor_Incline_Dynamic, DrawingType.Monitor_Dynamic
                    ListBoxMonitor_Dynamic.Items.Add(New LstbxDisplayAndItem(sht.Chart_App_Title, sht))

                Case DrawingType.Monitor_Static, DrawingType.Monitor_Incline_MaxMinDepth
                    ListBoxMonitor_Static.Items.Add(New LstbxDisplayAndItem(sht.Chart_App_Title, sht))
            End Select
        Next
    End Sub

    ''' <summary>
    ''' Word退出时，将对应的文档与word程序的变量设置为nothing
    ''' </summary>
    ''' <param name="docBeingClosed"></param>
    ''' <param name="cancel"></param>
    ''' <remarks></remarks>
    Private Sub word_Quit(ByVal docBeingClosed As Word.Document, ByRef cancel As Boolean) Handles WdApp.DocumentBeforeClose
        Dim a As Single = 0
        If docBeingClosed.Name = WdDoc.Name Then
            WdDoc = Nothing
            WdApp = Nothing
        End If
    End Sub

    Private Sub Diafrm_Output_Word_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        For Each H As IntPtr In WindowHandles
            EnableWindow(H, True)
        Next
    End Sub

#End Region

    ''' <summary>
    ''' 将结果输出到word中
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

        If Me.F_SelectedDrawings.Count > 0 Then

            If BackgroundWorker1.IsBusy <> True Then
                ' Start the asynchronous operation.
                BackgroundWorker1.RunWorkerAsync(Me.F_SelectedDrawings)
            End If
        Else
            Exit Sub
        End If
    End Sub

#Region "  ---  后台线程的操作"

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim selectedDrawings As Drawings_For_Output = DirectCast(e.Argument, Drawings_For_Output)

        '打开Word程序
        If WdApp Is Nothing Then
            WdApp = New Word.Application
        End If
        '指定进行输入的Word文档
        If WdDoc Is Nothing Then
            '以模板文件打开新文档
            Dim t_path As String = System.IO.Path.Combine(My.Settings.Path_Template, _
                                                          Constants.FolderOrFileName.File_Template.Word_Output)
            WdDoc = WdApp.Documents.Add(Template:=t_path)
            '文档的正文宽度，用以限制图形的宽度
            ContentWidth = GetContentWidth(WdDoc)
        End If
        '设置界面的可见性
        With WdApp
            If .Visible = True Then .Visible = True '即保持原来的可见性
            .ScreenUpdating = False
        End With

        '  --------------------------- 输出 ------------------------------------

        Call ExportToWord(WdApp, selectedDrawings)

        '  --------------------------- 输出 ------------------------------------
        WdApp.Visible = True
        WdApp.ScreenUpdating = True
    End Sub

    Private Sub ExportToWord(ByVal WdApp As Word.Application, ByVal selectedDrawings As Drawings_For_Output)
        Dim rg As Word.Range = WdDoc.Range(Start:=0)
        '在写入标题部分内容时所占的进度
        Dim intProgressForStartPart As Integer = 10
        '一共要导出的元素个数
        Dim intElementsCount As Integer = selectedDrawings.Count
        '每一个导出的元素所占的进度
        Dim sngUnit As Single = (100 - intProgressForStartPart) / intElementsCount
        '实时的进度值
        Dim intProgress As Integer = intProgressForStartPart
        Try
            '写入标题项
            Call Export_OverView(rg)
        Catch ex As Exception
            MessageBox.Show("写入概述部分出错，但可以继续工作。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgressForStartPart)
        End Try

        ' ------------- 取消绘图窗口的禁用 ------------------
        '一定要在将绘图窗口中的图形导出到Word之前取消窗口的禁用，
        '否则的话，当调用这些窗口的Application属性时，就会出现报错：应用程序正在使用中。
        For Each H As IntPtr In WindowHandles
            EnableWindow(H, True)
        Next

        '输出每一个选定的图形
        ' ------------- 开挖平面图 ------------------
        Try
            Dim D As ClsDrawing_PlanView = selectedDrawings.PlanView
            If D IsNot Nothing Then
                Dim page As Visio.Page = D.Page
                '
                NewLine(rg, Constants.ParagraphStyle.Title_2)
                rg.InsertAfter("开挖平面图：")
                '
                Call Export_VisioPlanview(page, rg)
                '
                intProgress += sngUnit
                APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress)
            End If
        Catch ex As Exception
            MessageBox.Show("导出Visio开挖平面图出错，但可以继续工作。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
         "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        ' ------------- 剖面标高图 -------------------------
        Try
            Dim D As ClsDrawing_ExcavationElevation = selectedDrawings.SectionalView
            If D IsNot Nothing Then
                '
                NewLine(rg, Constants.ParagraphStyle.Title_2)
                rg.InsertAfter("开挖剖面图：")
                '
                Call Export_ExcelChart(D.Chart, rg)
                '
                intProgress += sngUnit
                APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress)
            End If
        Catch ex As Exception
            MessageBox.Show("导出Excel开挖剖面图出错，但可以继续工作。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
         "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        ' ---------------------- 监测曲线图 --------------------
        Dim cht As Excel.Chart
        For Each Drawing As ClsDrawing_Mnt_Base In selectedDrawings.MntDrawings
            Try
                Select Case Drawing.Type
                    ' ------------- 测斜曲线图 ---------------------------------------------------
                    Case DrawingType.Monitor_Incline_Dynamic
                        Dim D As ClsDrawing_Mnt_Incline = CType(Drawing, ClsDrawing_Mnt_Incline)
                        cht = D.Chart
                        '
                        NewLine(rg, Constants.ParagraphStyle.Title_2)
                        rg.InsertAfter(D.Chart_App_Title)
                        '
                        Call Export_ExcelChart(cht, rg)

                        ' ------------- 动态监测曲线图 ---------------------------------------------
                    Case DrawingType.Monitor_Dynamic
                        Dim D As ClsDrawing_Mnt_OtherDynamics = CType(Drawing, ClsDrawing_Mnt_OtherDynamics)
                        cht = D.Chart

                        '
                        NewLine(rg, Constants.ParagraphStyle.Title_2)
                        rg.InsertAfter(D.Chart_App_Title)

                        Call Export_ExcelChart(cht, rg)

                        ' ------------- 静态监测曲线图 ---------------------------------------------
                    Case DrawingType.Monitor_Static
                        Dim D As ClsDrawing_Mnt_Static = CType(Drawing, ClsDrawing_Mnt_Static)
                        cht = D.Chart
                        '
                        NewLine(rg, Constants.ParagraphStyle.Title_2)
                        rg.InsertAfter(D.Chart_App_Title)

                        Call Export_ExcelChart(cht, rg)
                        ' ------------- 静态监测曲线图 ---------------------------------------------
                    Case DrawingType.Monitor_Incline_MaxMinDepth
                        Dim D As ClsDrawing_Mnt_MaxMinDepth = CType(Drawing, ClsDrawing_Mnt_MaxMinDepth)
                        cht = D.Chart
                        '
                        NewLine(rg, Constants.ParagraphStyle.Title_2)
                        rg.InsertAfter(D.Chart_App_Title)

                        Call Export_ExcelChart(cht, rg)
                    Case Else

                End Select
            Catch ex As Exception
                MessageBox.Show("导出监测曲线图""" & Drawing.Chart_App_Title.ToString & """出错，但可以继续工作。" &
                                vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                                ex.TargetSite.Name, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Finally
                intProgress += sngUnit
                APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress)

            End Try
        Next
    End Sub

    ''' <summary>
    ''' 输出开头的一些概况信息
    ''' </summary>
    ''' <param name="Range"></param>
    ''' <remarks>包括标题、施工日期，施工工况等</remarks>
    Private Sub Export_OverView(ByRef Range As Word.Range)

        With Range
            If .End > 1 Then        '说明当前range不是在文档的开头，那么就要新起一行
                NewLine(Range, Constants.ParagraphStyle.Title_1)
            Else                    '说明当前range就是在文档的开头，那么就直接设置段落样式为“标题”就可以了。
                Range.ParagraphFormat.Style = Constants.ParagraphStyle.Title_1
            End If
            .InsertAfter(My.Settings.ProjectName & " 实测数据动态分析：" & dateThisday.ToShortDateString)
            '
            NewLine(Range, Constants.ParagraphStyle.Content)
            .InsertAfter("施工日期： " & dateThisday.ToLongDateString + Chr(13) + "施工工况： （***）")
        End With

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.Close()
        '
        APPLICATION_MAINFORM.MainForm.HideProgress("图形导出完成！")
    End Sub

#End Region

#Region "  ---  图表输出操作"

    ''' <summary>
    ''' 导出excel中的chart对象到word中
    ''' </summary>
    ''' <param name="cht">excel中的chart对象</param>
    ''' <param name="range">此时word文档中的全局range的位置或者范围</param>
    ''' <remarks>由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
    Private Sub Export_ExcelChart(ByVal cht As Excel.Chart, ByRef range As Word.Range)
        cht.Application.ScreenUpdating = False
        Try
            ' 下面复制Chart的操作中，如果监测曲线图所使用的Chart模板有问题，则可能会出错。
            Dim chtObj As Excel.ChartObject = cht.Parent
            chtObj.Activate()
            chtObj.Copy()       ' 或者用  cht.ChartArea.Copy()都可以。
        Catch ex As Exception
            MessageBox.Show("导出监测曲线图""" & cht.Application.Caption.ToString & """出错（请检查是否是用户使用的Chart模板有问题），跳过此图的导出。" &
                          vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                          ex.TargetSite.Name, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '刷新excel屏幕
            cht.Application.ScreenUpdating = True
            Return
        End Try
        '设置word.range的格式
        With range
            '新起一行，并设置新段落的段落样式为图片
            NewLine(range, Constants.ParagraphStyle.picture)

            '进行粘贴，下面也可以用：DataType:=23
            .PasteSpecial(DataType:=Word.WdPasteDataType.wdPasteOLEObject, _
                          Placement:=Word.WdOLEPlacement.wdInLine)

        End With

        Dim shp As Word.InlineShape
        range.Select()
        With range.Application.Selection
            .MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)
            shp = .InlineShapes(1)
        End With
        '约束图形的宽度，将其限制在word页面的正文宽度之内
        Call WidthRestrain(shp, ContentWidth)

        '刷新excel屏幕
        cht.Application.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' 未使用。导出excel的工作表中的所有形状到word中
    ''' </summary>
    ''' <param name="sheet">要进行粘贴的Excel中的工作表对象</param>
    ''' <param name="range">此时word文档中的全局range的位置或者范围</param>
    ''' <remarks>在此方法中，会将Excel的工作表中的所有形状进行选择，然后进行组合，最后将其输出到word中；
    ''' 由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
    Private Sub Export_ExcelSheet(ByVal sheet As Excel.Worksheet, ByRef range As Word.Range)
        With sheet
            .Application.ScreenUpdating = False

            '在excel中将工作表里的所有形状进行复制并组合
            .Shapes.SelectAll()
            Dim shprg As Excel.ShapeRange = .Application.Selection.ShapeRange
            If shprg.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                Dim shp As Excel.Shape
                shp = shprg(0)          'ShapeRange中的第一个形状的下标值为0
                shp.Copy()
            Else
                shprg.Group.Copy()
            End If

        End With

        With range
            '新起一行，并设置新段落的段落样式为图片
            NewLine(range, Constants.ParagraphStyle.picture)

            ' ----------------------------  进行粘贴。这里的DataType不能指定为wdPasteOLEObject，
            '因为从Excel中复制过来的图片，它不是一个OLE对象。
            .PasteSpecial(DataType:=23, Placement:=Word.WdOLEPlacement.wdInLine)

            Dim shp As Word.Shape
            ' 获取刚刚粘贴过来的图片，此时图片很可能不是以嵌入式粘贴进来的。
            shp = .ShapeRange(1)
            '约束图形的宽度
            Call WidthRestrain(shp, ContentWidth)
            '将shape转换为inlineshape
            shp.ConvertToInlineShape()
        End With
        '刷新excel屏幕
        sheet.Application.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' 导出visio的Page中的所有形状到word中
    ''' </summary>
    ''' <param name="Page">要进行粘贴的Visio中的页面</param>
    ''' <param name="range">此时word文档中的全局range的位置或者范围</param>
    ''' <remarks>在此方法中，会将visio指定页面中的所有形状进行选择，然后进行组合，最后将其输出到word中；
    ''' 由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
    Private Sub Export_VisioPlanview(ByVal Page As Visio.Page, ByRef range As Word.Range)

        Dim app As Visio.Application = Page.Application
        '
        Dim wnd As Visio.Window
        wnd = app.ActiveWindow
        wnd.Page = Page


        With wnd
            .Activate()
            '这里要将ShowChanges设置为True，否则下面的SelectAll()方法会被禁止。
            app.ShowChanges = True
            .SelectAll()
            '  ---------------------- 耗时代码1：复制Visio的Page中的所有形状
            '而且在这一步的时候Visio的窗口中可能会弹出子窗口
            .Selection.Copy()
            '这一步也可能会导致Visio的窗口中弹出子窗口
            .DeselectAll()

            '关闭所有的子窗口
            'Debug.Print(app.ActiveWindow.Windows.Count)     ‘即使只显示出一个子窗口，这里也会返回10
            'For Each subWnd As Visio.Window In app.ActiveWindow.Windows
            '    subWnd.Visible = False
            'Next

            '根据实际情况：每次都只弹出“外部数据”这一子窗口，所以在这里就只对其进行单独隐藏。
            .Windows.ItemFromID(Visio.VisWinTypes.visWinIDExternalData).Visible = False

            '让窗口的显示适应页面
            .ViewFit = Visio.VisWindowFit.visFitPage
        End With
        With range
            '新起一行，并设置新段落的段落样式为图片
            NewLine(range, Constants.ParagraphStyle.picture)

            '  ---------------------- 耗时代码2：将Visio的Page中的所有形状粘贴到Word中（不可以用：DataType:=23）
            .PasteSpecial(DataType:=Word.WdPasteDataType.wdPasteOLEObject, Placement:=Word.WdOLEPlacement.wdInLine)


            Dim shp As Word.InlineShape
            range.Select()
            With range.Application.Selection
                .MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)
                shp = .InlineShapes(1)
            End With
            '约束图形的宽度，将其限制在word页面的正文宽度之内
            Call WidthRestrain(shp, ContentWidth)
        End With

        '刷新visio屏幕
        app.ShowChanges = True
    End Sub

#End Region

#Region "  ---  零碎方法"

    ''' <summary>
    ''' 新起一段
    ''' </summary>
    ''' <param name="range">Range对象，可以为选区范围或者光标插入点</param>
    ''' <param name="PrphStyle">新起一段的段落样式</param>
    ''' <remarks></remarks>
    Private Sub NewLine(ByRef range As Word.Range, ByVal PrphStyle As String)
        With range
            .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            .InsertParagraphAfter()
            .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            .ParagraphFormat.Style = PrphStyle
        End With
    End Sub

    ''' <summary>
    ''' 限制粘贴到word中的图形的宽度
    ''' </summary>
    ''' <param name="shape">粘贴过来的图形，可能为shape对象或者inlineshape对象</param>
    ''' <param name="PageWidth_Content">用来限制图片宽度的值，一般取页面的正文版面的宽度值</param>
    ''' <remarks></remarks>
    Private Sub WidthRestrain(ByVal shape As Object, ByVal PageWidth_Content As Single)
        With shape
            Dim W As Single = .width
            If W > PageWidth_Content Then
                Dim H As Single = .height
                Dim AspectRatio As Double = H / W
                '
                .width = PageWidth_Content
                .height = .width * AspectRatio
            End If
        End With
    End Sub

    ''' <summary>
    ''' 获取word页面的正文范围的宽度，用来限定图片的宽度值
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <returns>word页面的正文范围的宽度，以磅为单位</returns>
    ''' <remarks></remarks>
    Private Function GetContentWidth(ByVal doc As Word.Document) As Single
        '正文的区域的宽度CW
        Dim CW As Single
        Dim ps As Word.PageSetup = doc.PageSetup
        With ps
            Dim W As Single = .PageWidth
            Dim Margin As Single = .LeftMargin + .RightMargin
            CW = W - Margin
        End With
        Return CW
    End Function

#End Region

#Region "  ---  一般界面操作"

    ''' <summary>
    ''' 对列表中的项目进行全选或者取消全部选择
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ChkBxSelect_Click(sender As Object, e As EventArgs) Handles ChkBxSelect.Click

        Select Case ChkBxSelect.CheckState
            Case CheckState.Checked
                '执行全选操作
                With CheckBox_PlanView
                    If .Enabled Then
                        .Checked = True
                    End If
                End With
                With CheckBox_SectionalView
                    If .Enabled Then
                        .Checked = True
                    End If
                End With
                For Each lstbox As ListBox In F_arrListBoxes
                    Dim n As Short = lstbox.Items.Count
                    For index As Short = 0 To n - 1
                        lstbox.SetSelected(index, True)
                    Next
                Next

            Case CheckState.Unchecked, CheckState.Indeterminate
                '在UI上跳过中间状态，直接进入“取消全部选择”，并执行取消选择的操作。
                If ChkBxSelect.CheckState = CheckState.Indeterminate Then
                    ChkBxSelect.CheckState = CheckState.Unchecked
                End If
                '执行取消全部选择操作
                With CheckBox_PlanView
                    .Checked = False
                End With
                With CheckBox_SectionalView
                    .Checked = False
                End With
                For Each lstbox As ListBox In F_arrListBoxes
                    lstbox.ClearSelected()
                Next

        End Select
    End Sub

    '选择复选框或者列表项——更新选择的图表对象
    Private Sub CheckBox_PlanView_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_PlanView.CheckedChanged
        If CheckBox_PlanView.Checked Then
            Me.F_SelectedDrawings.PlanView = DirectCast(CheckBox_PlanView.Tag, ClsDrawing_PlanView)
        Else
            Me.F_SelectedDrawings.PlanView = Nothing
        End If
        '
        Call SelectedDrawingsChanged(Me.F_SelectedDrawings)
    End Sub
    Private Sub CheckBox_SectionalView_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_SectionalView.CheckedChanged
        If CheckBox_SectionalView.Checked Then
            Me.F_SelectedDrawings.SectionalView = DirectCast(CheckBox_SectionalView.Tag, ClsDrawing_ExcavationElevation)
        Else
            Me.F_SelectedDrawings.SectionalView = Nothing
        End If
        '
        Call SelectedDrawingsChanged(Me.F_SelectedDrawings)
    End Sub
    Private Sub ListBoxMonitor_Dynamic_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles ListBoxMonitor_Dynamic.SelectedIndexChanged, ListBoxMonitor_Static.SelectedIndexChanged
        Me.F_SelectedDrawings.MntDrawings.Clear()
        '
        Dim Drawing As ClsDrawing_Mnt_Base
        Dim Items = ListBoxMonitor_Dynamic.SelectedItems
        For Each lstboxItem As LstbxDisplayAndItem In Items
            Drawing = DirectCast(lstboxItem.Value, ClsDrawing_Mnt_Base)
            Me.F_SelectedDrawings.MntDrawings.Add(Drawing)
        Next
        Items = ListBoxMonitor_Static.SelectedItems
        For Each lstboxItem As LstbxDisplayAndItem In Items
            Drawing = DirectCast(lstboxItem.Value, ClsDrawing_Mnt_Base)
            Me.F_SelectedDrawings.MntDrawings.Add(Drawing)
        Next
        '
        Call SelectedDrawingsChanged(Me.F_SelectedDrawings)
    End Sub

    '选择复选框或者列表项——更新滚动线程与窗口界面
    ''' <summary>
    ''' ！选择的图形发生改变时，更新滚动线程与窗口界面。
    ''' </summary>
    ''' <param name="Selected_Drawings">更新后的要进行滚动的图形</param>
    ''' <remarks>此方法不能直接Handle复选框的CheckedChanged或者列表框的SelectedIndexChanged事件，
    ''' 因为此方法必须是在更新了Me.F_SelectedDrawings属性之后，才能去更新窗口界面。</remarks>
    Private Sub SelectedDrawingsChanged(ByVal Selected_Drawings As Drawings_For_Output)
        If Selected_Drawings.Count > 0 Then
            btnExport.Enabled = True
        Else
            btnExport.Enabled = False
        End If
    End Sub
    '
    Private Sub ProgressBar_PlanView_Click(sender As Object, e As EventArgs) Handles ProgressBar_PlanView.Click
        With CheckBox_PlanView
            If .Enabled Then
                If .Checked Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End If
        End With
    End Sub
    Private Sub ProgressBar_SectionalView_Click(sender As Object, e As EventArgs) Handles ProgressBar_SectionalView.Click
        With CheckBox_SectionalView
            If .Enabled Then
                If .Checked Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End If
        End With
    End Sub

#End Region

End Class