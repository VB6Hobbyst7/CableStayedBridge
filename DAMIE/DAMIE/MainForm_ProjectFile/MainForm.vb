Imports Microsoft.Office.Interop
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.DataBase
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form
Imports System.Windows.Forms
Imports System.Threading

Namespace GlobalApp_Form
    ''' <summary>
    ''' 程序的主界面
    ''' </summary>
    ''' <remarks></remarks>
    Public Class APPLICATION_MAINFORM

#Region "  ---  定义与声明"

#Region "  ---  字段定义"

        ''' <summary>
        ''' 全局的主程序
        ''' </summary>
        ''' <remarks></remarks>
        Private GlbApp As GlobalApplication

#End Region

#Region "  ---  属性值的定义"

        Private Shared F_main_Form As APPLICATION_MAINFORM
        ''' <summary>
        ''' 共享属性，用来索引启动窗口对象
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>这一操作并不多余，对于多线程操作，有可能会出现在其他线程不能正确地
        ''' 调用到这个唯一的主程序对象，此时可以用这个属性来返回其实例对象。</remarks>
        Public Shared ReadOnly Property MainForm As APPLICATION_MAINFORM
            Get
                Return F_main_Form
            End Get
        End Property

#Region "  ---  操作窗口"

        ''' <summary>
        ''' 图形滚动窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmRolling As frmRolling
        ''' <summary>
        ''' 图形滚动窗口
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Form_Rolling As frmRolling
            Get
                Return frmRolling
            End Get
            Set(value As frmRolling)
                frmRolling = value
            End Set
        End Property

        ''' <summary>
        ''' 生成剖面标高图窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmSectionView As frmDrawElevation
        ''' <summary>
        ''' 生成剖面标高图窗口
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Form_SectionalView As frmDrawElevation
            Get
                Return frmSectionView
            End Get
            Set(value As frmDrawElevation)
                frmSectionView = value
            End Set
        End Property

        ''' <summary>
        ''' 绘制测斜曲线的窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmMnt_Incline As frmDrawing_Mnt_Incline
        ''' <summary>
        ''' 绘制测斜曲线的窗口
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Form_Mnt_Incline As frmDrawing_Mnt_Incline
            Get
                Return frmMnt_Incline
            End Get
            Set(value As frmDrawing_Mnt_Incline)
                frmMnt_Incline = value
            End Set
        End Property

        ''' <summary>
        ''' 绘制其他监测曲线的窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmMnt_Others As frmDrawing_Mnt_Others
        ''' <summary>
        ''' 绘制其他监测曲线的窗口
        ''' </summary>
        ''' <value></value>
        ''' <remarks>从此窗口中可以生成非测斜曲线的其他曲线，并且包括其时间分布与空间分布的形式</remarks>
        Public Property Form_Mnt_Others As frmDrawing_Mnt_Others
            Get
                Return frmMnt_Others
            End Get
            Set(value As frmDrawing_Mnt_Others)
                frmMnt_Others = value
            End Set
        End Property

        ''' <summary>
        ''' 对项目文件进行操作的窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmProjectFile As DataBase.frmProjectFile
        ''' <summary>
        ''' 对项目文件进行操作的窗口
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Form_ProjectFile As DataBase.frmProjectFile
            Get
                Return Me.frmProjectFile
            End Get
        End Property

        ''' <summary>
        ''' 打开Visio的开挖平面图的窗口
        ''' </summary>
        ''' <remarks></remarks>
        Private frmVisioPlanView As frmDrawingPlan
        ''' <summary>
        ''' 打开Visio的开挖平面图的窗口
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Form_VisioPlanView As frmDrawingPlan
            Get
                If Me.frmVisioPlanView Is Nothing Then
                    Me.frmVisioPlanView = New frmDrawingPlan
                End If
                Return Me.frmVisioPlanView
            End Get
        End Property
#End Region

#Region "  - - 逻辑标志 布尔值"

        ''' <summary>
        ''' 布尔值，用来指示主程序中的绘图窗口是否有新添加的，或者是否被关闭
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DrawingWindowChanged As Boolean

#End Region

#End Region

#End Region

#Region "  ---  主程序的加载与关闭"

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

            ' This call is required by the designer.
            InitializeComponent()
            ' Add any initialization after the InitializeComponent() call.
            '----------------------------
            Me.GlbApp = New GlobalApplication
            '为关键字段赋初始值()
            APPLICATION_MAINFORM.F_main_Form = Me
            '----------------------------
            Call MainUI_ProjectNotOpened()
            '获取与文件或文件夹路径有关的数据
            Call GetPath()
            '----------------------------
            '创建新窗口，窗口在创建时默认是不隐藏的。
            Me.frmSectionView = New frmDrawElevation
            Me.frmMnt_Incline = New frmDrawing_Mnt_Incline
            Me.frmMnt_Others = New frmDrawing_Mnt_Others
            Me.frmRolling = New frmRolling()

            ' ----------------------- 设置MDI窗体的背景
            For Each C As Control In Me.Controls
                If String.Compare(C.GetType.ToString, "System.Windows.Forms.MdiClient", True) = 0 Then
                    Dim MDIC As MdiClient = DirectCast(C, MdiClient)
                    With MDIC
                        .BackgroundImage = My.Resources.线条背景
                        .BackgroundImageLayout = ImageLayout.Tile
                    End With
                    Exit For
                End If
            Next
            ' ----------------------- 设置主程序窗口启动时的状态
            With Me
                Dim mysettings As New mySettings_UI
                Dim winState As FormWindowState = mysettings.WindowState
                Select Case winState
                    Case FormWindowState.Maximized
                        .WindowState = winState
                    Case FormWindowState.Minimized
                        .WindowState = FormWindowState.Normal
                    Case FormWindowState.Normal
                        .WindowState = winState
                        .Location = mysettings.WindowLocation
                        .Size = mysettings.WindowSize
                End Select
            End With
            '在新线程中进行程序的一些属性的初始值的设置
            Dim thd As New Thread(AddressOf Me.myDefaltSettings)
            With thd
                .Name = "程序的一些属性的初始值的设置"
                .Start()
            End With

        End Sub
        ''' <summary>
        ''' 程序的一些属性的初始值的设置，这些属性是与UI线程无关的属性，以在新的非UI线程中进行设置。
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub myDefaltSettings()
            Dim setting1 As New mySettings_Application
            With setting1

                Dim struct As New ClsDrawing_PlanView.MonitorPointsInformation
                With struct
                    .ShapeName_MonitorPointTag = "Tag"
                    .pt_CAD_BottomLeft = New PointF(309598.527, -119668.436)
                    .pt_CAD_UpRight = New PointF(536642.644, 201852.14)
                    .pt_Visio_BottomLeft_ShapeID = 197
                    .pt_Visio_UpRight_ShapeID = 217
                End With
                .MonitorPointsInfo = struct
                '在下面的Save方法中，不知为何为出现两次相同的报错：System.IO.FileNotFoundException
                '可以明确其于多线程无关，但是好在此异常对于程序的运行无影响。
                .Save()
            End With

        End Sub

        ''' <summary>
        ''' 主程序加载
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>在主程序界面加载时，先启动Splash Screen，再为关键字段赋值，
        ''' 然后用一个伪窗口作为主程序界面的背景，最后关闭Splash Screen窗口。</remarks>
        Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
            '-----------------------根据程序的启动方式的不同，作出不同的操作
            Dim s As ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
            If s.Count = 1 Then     '
                Dim StartFilePath As String = s.Item(0)
                If File.Exists(StartFilePath) Then
                    If String.Compare(Path.GetExtension(StartFilePath), Constants.AMEApplication.FileExtension, True) = 0 Then
                        Call OpenProjectFile(StartFilePath)
                        Call MainUI_ProjectOpened()
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' 获取与文件或文件夹路径有关的数据,并保存在My.Settings中
        ''' </summary>
        ''' <remarks>用来保证程序在不同机器或文件夹间迁移时能够正常索引</remarks>
        Private Sub GetPath()
            With My.Settings

                '主程序.exe文件所在的文件夹路径，比如：F:\基坑数据\程序编写\群坑分析\AME\bin\Debug
                .Path_MainForm = My.Application.Info.DirectoryPath

                '"Templates"
                .Path_Template = System.IO.Path.Combine(.Path_MainForm, Constants.FolderOrFileName.Folder.Template)

                '用来进行输出的文件夹
                .Path_Output = System.IO.Path.Combine(.Path_MainForm, Constants.FolderOrFileName.Folder.Output)

                .Path_DataBase = System.IO.Path.Combine(.Path_MainForm, Constants.FolderOrFileName.Folder.DataBase)
            End With

        End Sub

        ''' <summary>
        ''' 退出主程序
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>
        ''' 这里用FormClosing事件来控制主程序的退出，而不用Dispose事件来控制，
        ''' 是为了解决其子窗口在各自的FormClosing事件中只是将其进行隐藏并取消了默认的关闭操作，
        ''' 所以这里在主程序最先触发的FormClosing事件中就直接执行Me.Dispose()方法，这样就可以
        ''' 跳过子窗口的FormClosing事件，而直接退出主程序了。
        ''' </remarks>
        Private Sub mainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
            Dim AllDrawing As AmeDrawings = GlbApp.ExposeAllDrawings
            If AllDrawing.Count > 0 Then
                Dim result As DialogResult = MessageBox.Show("还有图表未处理，是否要关闭所有绘图并退出程序", "tip",
                                                             MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
                                                             MessageBoxDefaultButton.Button3)
                Select Case result
                    Case Windows.Forms.DialogResult.Yes     '关闭AME主程序，同时也关闭所有的绘图程序。
                        Me.Hide()
                        '' --- 通过新的工作线程来执行绘图程序的关闭。
                        'Dim thd As New Thread(AddressOf Me.QuitDrawingApplications)
                        'With thd
                        '    .Name = "关闭所有绘图程序"
                        '    .Start(AllDrawing)
                        '    thd.Join()
                        'End With
                        ' --- 
                        '通过Main Thread来执行绘图程序的关闭。
                        Call Me.QuitDrawingApplications(AllDrawing)
                    Case Windows.Forms.DialogResult.No       '不关闭图表，但关闭程序

                    Case Windows.Forms.DialogResult.Cancel   '图表与程序都不关闭
                        e.Cancel = True
                        Exit Sub
                End Select
            End If
            ' ---------------------- 先隐藏主界面，以达到更好的UI效果。
            Me.Hide()
            '--------------------------- 断后工作
            Try
                '关闭隐藏的Excel数据库中的所有工作簿
                For Each wkbk As Excel.Workbook In GlbApp.ExcelApplication_DB.Workbooks
                    wkbk.Close(False)
                Next
                GlbApp.ExcelApplication_DB.Quit()
                '这一步非常重要哟
                GlbApp.ExcelApplication_DB = Nothing
            Catch ex As Exception
                '有可能会出现数据库文件已经被关闭的情况
            End Try

            '保存主程序窗口关闭时的界面位置与大小
            With Me
                Dim mysetting As New mySettings_UI
                mysetting.WindowLocation = .Location
                mysetting.WindowSize = .Size
                mysetting.WindowState = .WindowState
                mysetting.Save()
            End With

            '---------------------------
            '这里在主程序最先触发的FormClosing事件中就直接执行Me.Dispose()方法，
            '这样就可以跳过子窗口的FormClosing事件，而直接退出主程序了。
            Me.Dispose()
        End Sub

        ''' <summary>
        ''' 关闭程序中的所有绘图所在的程序，如Excel或者Visio的程序
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub QuitDrawingApplications(ByVal AllDrawing As AmeDrawings)

            With AllDrawing
                '开挖剖面图
                If .SectionalView IsNot Nothing Then
                    .SectionalView.Close(False)
                End If
                'Visio开挖平面图
                If .PlanView IsNot Nothing Then
                    .PlanView.Close(False)
                End If
                '监测曲线图
                For Each MntDrawing As ClsDrawing_Mnt_Base In .MonitorData
                    MntDrawing.Close(False)
                Next
            End With
        End Sub

        ''' <summary>
        ''' 退出程序
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItemExit_Click(sender As Object, e As EventArgs) Handles MenuItemExit.Click
            Me.OnFormClosing(New System.Windows.Forms.FormClosingEventArgs(CloseReason.ApplicationExitCall, False))
        End Sub

#End Region

#Region "  ---  一般界面操作"
        '项目文件的新建与打开
        Private Sub MenuItem_NewProject_Click(sender As Object, e As EventArgs) Handles MenuItem_NewProject.Click
            Call Me.NewProjectFile()
        End Sub
        Private Sub MenuItem_OpenProject_Click(sender As Object, e As EventArgs) Handles MenuItem_OpenProject.Click
            Dim FilePath As String = ""
            With OpenFileDialog1
                .Title = "选择项目文件"
                Dim FileExtension As String = AMEApplication.FileExtension
                .Filter = FileExtension & "文件(*" & FileExtension & ")|*" & FileExtension
                .FilterIndex = 1
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    FilePath = .FileName
                End If
            End With
            If FilePath.Length > 0 Then
                Call Me.OpenProjectFile(FilePath)
            End If
        End Sub
        Private Sub MenuItem_EditProject_Click(sender As Object, e As EventArgs) Handles MenuItem_EditProject.Click
            Call Me.EditProjectFile()
        End Sub
        '拖拽操作
        Private Sub APPLICATION_MAINFORM_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
            Dim FileDrop As String() = e.Data.GetData(DataFormats.FileDrop)
            ' DoSomething with the Files or Directories that are droped in.
            Dim filepath As String = FileDrop(0)
            Dim ext As String = Path.GetExtension(filepath)

            If String.Compare(ext, AMEApplication.FileExtension, True) = 0 Then
                Call Me.OpenProjectFile(filepath)
            Else
                MessageBox.Show("Can not open file" & filepath & ". Verify that the file is a(an)" _
                                & AMEApplication.FileExtension & " file.",
                                 "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub
        Private Sub APPLICATION_MAINFORM_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
            ' See if the data includes text.
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                ' There is text. Allow copy.
                e.Effect = DragDropEffects.Copy
            Else
                ' There is no text. Prohibit drop.
                e.Effect = DragDropEffects.None
            End If

        End Sub

        'Arrange,子窗口重排
        ''' <summary>
        ''' 子窗口水平排列
        ''' </summary>
        Private Sub ChildrenFormAlligment_Horizontal(sender As Object, e As EventArgs) Handles MenuItem_Arrange_Horizontal.Click
            Me.LayoutMdi(MdiLayout.TileHorizontal)
        End Sub
        ''' <summary>
        ''' 子窗口竖直排列
        ''' </summary>
        Private Sub ChildrenFormAlligment_Vertical(sender As Object, e As EventArgs) Handles MenuItem_Arrange_Vertical.Click
            Me.LayoutMdi(MdiLayout.TileVertical)
        End Sub
        ''' <summary>
        ''' 子窗口层叠
        ''' </summary>
        Private Sub ChildrenFormAlligment_Cascade(sender As Object, e As EventArgs) Handles MenuItem_Arrange_Cascade.Click
            Me.LayoutMdi(MdiLayout.Cascade)
        End Sub

#End Region

#Region "  ---  绘制图表"

        ''' <summary>
        ''' 生成剖面标高图
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItemSectionalView_Click(sender As Object, e As EventArgs) Handles MenuItemSectionalView.Click
            '数据库中所有的开挖区域
            With GlbApp.DataBase
                If .ID_Components.Count = 0 Then
                    MessageBox.Show("没有发现基坑标高相关的数据，请先在项目中添加相关数据。",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    With frmSectionView
                        .MdiParent = Me
                        .Show()
                        .WindowState = FormWindowState.Maximized
                    End With
                End If
            End With

        End Sub
        '
        ''' <summary>
        ''' 生成Visio的平面开挖图
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Sub ShowForm_DrawingVisioPlanView(sender As Object, e As EventArgs) Handles MenuItemPlanView.Click
            Call GlobalApplication.Application.DrawVisioPlanView()
        End Sub

        ''' <summary>
        ''' 生成测斜曲线图
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItemDataMonitored_Click(sender As Object, e As EventArgs) Handles MenuItemMntData_Incline.Click
            With frmMnt_Incline
                .MdiParent = Me
                .Show()
                .WindowState = FormWindowState.Maximized
            End With
        End Sub

        ''' <summary>
        ''' 生成其他监测数据曲线图
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItemOtherCurves_Click(sender As Object, e As EventArgs) Handles MenuItemMntData_Others.Click

            'With Me.frmMnt_Others
            '    myAPI.SetParent(.Handle, Me.PictureBox_Background.Handle)
            '    .WindowState = FormWindowState.Maximized
            '    .Show()
            'End With
            With frmMnt_Others
                .MdiParent = Me
                .Show()
                .WindowState = FormWindowState.Maximized
            End With
        End Sub

        ''' <summary>
        ''' 执行Rolling操作，以进行窗口的同步滚动
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub StartRolling(sender As Object, e As EventArgs) Handles TlStrpBtn_Roll.Click
            With frmRolling
                .MdiParent = Me
                .Show()
                .WindowState = FormWindowState.Maximized
            End With
        End Sub

#End Region

#Region "  ---  对话框式界面"


        ''' <summary>
        ''' 将最终结果输出的Word中的窗口对象
        ''' </summary>
        ''' <remarks></remarks>
        Private frm_Output_Word As Diafrm_Output_Word
        ''' <summary>
        ''' 将最终结果输出的Word中的
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Sub ExportToWord(sender As Object, e As EventArgs) Handles MenuItemExport.Click
            If frm_Output_Word Is Nothing Then frm_Output_Word = New Diafrm_Output_Word
            frm_Output_Word.ShowDialog()
        End Sub

#End Region

#Region "  ---  项目文件的新建、打开、保存等"

        ''' <summary>
        ''' 新建项目文件
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub NewProjectFile()

            Me.frmProjectFile = New DataBase.frmProjectFile
            With Me.frmProjectFile
                .ProjectState = ProjectState.NewProject
                .ShowDialog()
            End With

        End Sub

        ''' <summary>
        ''' 打开项目文件
        ''' </summary>
        ''' <param name="FilePath">打开的文件的文件路径（此文件已经确定为项目后缀的文件）</param>
        ''' <remarks>单纯的打开项目文件并不要求打开操作项目文件的窗口</remarks>
        Private Sub OpenProjectFile(ByVal FilePath As String)
            '新开一个线程来执行打开文件的操作
            Dim thd As New Thread(AddressOf Me.OpenFile)
            With thd
                .Name = "打开项目文件"
                .Start(FilePath)
            End With
        End Sub
        ''' <summary>
        ''' 在工作者线程中执行具体的打开文件的工作
        ''' </summary>
        ''' <param name="FilePath"></param>
        ''' <remarks></remarks>
        Private Sub OpenFile(ByVal FilePath As String)
            '在主程序界面上显示出进度条
            Me.ShowProgressBar_Marquee()
            '
            With GlbApp
                .ProjectFile = New clsProjectFile(FilePath)
                .ProjectFile.LoadFromXmlFile()
                .DataBase = New ClsData_DataBase(.ProjectFile.Contents)
            End With
            '隐藏进度条
            Me.HideProgress("File Opened")
        End Sub

        ''' <summary>
        ''' 编辑项目文件
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub EditProjectFile()
            If Me.frmProjectFile Is Nothing Then
                Me.frmProjectFile = New DataBase.frmProjectFile
            End If
            With Me.frmProjectFile
                .ProjectState = ProjectState.EditProject
                .ShowDialog()
            End With
        End Sub

        ''' <summary>
        ''' 保存项目文件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItem_SaveProject_Click(sender As Object, e As EventArgs) Handles MenuItem_SaveProject.Click
            '执行保存文件的操作
            With GlbApp.ProjectFile
                If .FilePath = Nothing Then
                    Call MenuItem_SaveAsProject_Click(sender, e)
                Else
                    .SaveToXmlFile()
                End If
            End With

        End Sub

        ''' <summary>
        ''' 关闭项目文件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItem_CloseProject_Click(sender As Object, e As EventArgs) Handles MenuItem_CloseProject.Click

        End Sub

        ''' <summary>
        ''' 另存为项目文件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItem_SaveAsProject_Click(sender As Object, e As EventArgs) Handles MenuItem_SaveAsProject.Click
            Dim FinalPathToSave As String = ""
            With Me.SaveFileDialog1
                Dim FileExtension As String = AMEApplication.FileExtension
                .Filter = FileExtension & "文件(*" & FileExtension & ")|*" & FileExtension
                .DefaultExt = Constants.AMEApplication.FileExtension
                .AddExtension = True
                .OverwritePrompt = True
                '打开对话框
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    '获取选择的路径
                    FinalPathToSave = .FileName
                    If FinalPathToSave.Length > 0 Then
                        With GlbApp.ProjectFile
                            '执行保存文件的操作
                            .FilePath = FinalPathToSave
                            .SaveToXmlFile()
                        End With
                    End If
                End If
            End With
        End Sub

#End Region

#Region "  ---  数据的提取与格式化（从Excel或Word）"

        Private Sub ExtractDataFromExcel(sender As Object, e As EventArgs) Handles ToolStripMenuItemExtractDataFromExcel.Click
            Dim a As New frmDeriveData_Excel
            a.ShowDialog()
        End Sub

        Private Sub ExtractDataFromWord(sender As Object, e As EventArgs) Handles ToolStripMenuItemExtractDataFromWord.Click
            Dim a As New frmDeriveData_Word
            Try
                a.ShowDialog()
            Catch ex As System.Reflection.TargetInvocationException
                Debug.Print("窗口关闭出错！")
            End Try
        End Sub

        Private Sub VisioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VisioToolStripMenuItem.Click
            Dim a As New Visio_DataRecordsetLinkToShape
            a.Show()
        End Sub

#End Region

        ''' <summary>
        ''' 在Visio平面图中绘制监测点位图
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MenuItemDrawingPoints_Click(sender As Object, e As EventArgs) Handles MenuItemDrawingPoints.Click
            Call GlobalApplication.Application.DrawingPointsInVisio()
        End Sub
    End Class
End Namespace