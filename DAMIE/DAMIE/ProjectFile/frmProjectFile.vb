Imports System.Windows.Forms
Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Constants

Namespace DataBase

    ''' <summary>
    ''' 对项目文件进行操作的窗口，比如新建项目、打开项目、编辑项目等
    ''' </summary>
    ''' <remarks></remarks>
    Public Class frmProjectFile

#Region "  ---  定义与声明"

#Region "  ---  Events"

        ''' <summary>
        ''' 工作簿列表框中的工作簿对象发生变化时触发的事件
        ''' </summary>
        ''' <param name="Sender"></param>
        ''' <param name="FileContents"></param>
        ''' <param name="clearListBox_Progress">是否要清空项目文件中的施工进度列表框中的内容</param>
        ''' <remarks></remarks>
        Private Event WorkBookInProjectChanged(ByVal Sender As Object, ByVal FileContents As clsData_FileContents, ByVal clearListBox_Progress As Boolean)

#End Region

#Region "  ---  Constants"

        ''' <summary>
        ''' 列表框的DataSource中用来表示Value的属性名
        ''' </summary>
        ''' <remarks></remarks>
        Private Const cstValueMember As String = LstbxDisplayAndItem.ValueMember
        ''' <summary>
        ''' 列表框的DataSource中用来显示在UI界面中的DisplayMember的属性名
        ''' </summary>
        ''' <remarks></remarks>
        Private Const cstDisplayMember As String = LstbxDisplayAndItem.DisplayMember

#End Region

#Region "  ---  Properties"

        ''' <summary>
        ''' 窗口打开时的作用状态
        ''' </summary>
        ''' <remarks></remarks>
        Private P_ProjectState As ProjectState
        ''' <summary>
        ''' 窗口打开时的作用状态
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProjectState As ProjectState
            Get
                Return Me.P_ProjectState
            End Get
            Set(value As ProjectState)
                Me.P_ProjectState = value
            End Set
        End Property

#End Region

#Region "  ---  Fields"
        Private GlbApp As GlobalApplication = GlobalApplication.Application

        ''' <summary>
        ''' 窗口中的界面所反映出的项目文件内容。
        ''' 此变量是为了在窗口中点击确认的时候赋值给主程序，
        ''' 而如果窗口不点击确定的话，主程序的变量就不会被更新。
        ''' </summary>
        ''' <remarks></remarks>
        Private F_NewFileContents As clsData_FileContents
        ''' <summary>
        ''' 在打开此窗口时，程序中的数据库文件对象，此对象在此窗口中是只读的。
        ''' </summary>
        ''' <remarks></remarks>
        Private F_OldFileContents As New clsData_FileContents
#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

        ''' <summary>
        ''' 窗口加载
        ''' </summary>
        Private Sub frmProjectContents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            '这一步的New非常重要，因为在每一次编辑完成后都会将其传址给主程序。
            F_NewFileContents = New clsData_FileContents
            '
            APPLICATION_MAINFORM.MainForm.AllowDrop = False
            Dim GlobalProjectfile As clsProjectFile = GlbApp.ProjectFile
            If GlobalProjectfile IsNot Nothing Then
                Me.F_OldFileContents = GlobalProjectfile.Contents
                Call FileContentsToUI(Me.F_OldFileContents)
                '在Label中更新此项目文件的绝对路径
                Me.LabelProjectFilePath.Text = GlobalProjectfile.FilePath

            End If

            '
            Call Me.BindProperty()

            '根据不同的窗口状态设置窗口的样式
            Select Case Me.P_ProjectState
                Case Miscellaneous.ProjectState.NewProject
                    With Me
                        .Text = "New Project"
                    End With

                Case Miscellaneous.ProjectState.EditProject
                    '将项目文件中的内容更新的窗口中
                    With Me
                        .Text = "Edit Project"
                    End With
            End Select

        End Sub
        ''' <summary>
        ''' 为列表框绑定文本显示与数据的属性值
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BindProperty()
            '设置列表框的显示文本的属性
            Dim AllListControl As New List(Of ListControl)
            With AllListControl
                .Clear()
                .Add(Me.CmbbxPlan)
                .Add(Me.CmbbxPointCoordinates)
                .Add(Me.CmbbxProgressWkbk)
                .Add(Me.CmbbxSectional)
                .Add(Me.CmbbxWorkingStage)
                .Add(Me.LstbxSheetsProgressInProject)
                .Add(Me.LstbxSheetsProgressInWkbk)
                .Add(Me.LstBxWorkbooks)
            End With
            For Each lstControl As ListControl In AllListControl
                With lstControl
                    .DisplayMember = cstDisplayMember
                    .ValueMember = cstValueMember
                End With
            Next
        End Sub

        Private Sub frmProjectFile_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
            APPLICATION_MAINFORM.MainForm.AllowDrop = True
        End Sub
#End Region

        '项目文件刷新到窗口UI
        ''' <summary>
        ''' 在打开窗口时将主程序中的ProjectFile对象的信息反映到窗口的控件中
        ''' </summary>
        ''' <param name="FileContents"></param>
        ''' <remarks></remarks>
        Public Sub FileContentsToUI(ByVal FileContents As clsData_FileContents)
            If FileContents IsNot Nothing Then
                With FileContents

                    '显示出项目中的工作簿
                    Dim listWkbks As New List(Of LstbxDisplayAndItem)
                    For Each wkbk As Workbook In .lstWkbks
                        listWkbks.Add(New LstbxDisplayAndItem(wkbk.FullName, wkbk))
                    Next
                    Me.LstBxWorkbooks.DataSource = listWkbks

                    '显示出项目文件中的施工进度工作表
                    Dim DataSource_SheetsProgressInProject As New List(Of LstbxDisplayAndItem)
                    For Each shtProgress As Worksheet In .lstSheets_Progress
                        DataSource_SheetsProgressInProject.Add(New LstbxDisplayAndItem(DisplayedText:=shtProgress.Name, _
                                                                                  Value:=shtProgress))
                    Next
                    Me.LstbxSheetsProgressInProject.DataSource = DataSource_SheetsProgressInProject
                End With
                '将数据源更新到所有的列表框
                RaiseEvent WorkBookInProjectChanged(Nothing, FileContents, False)
            End If
        End Sub

#Region "  ---  !项目内容与窗口显示的交互"

        '项目中的数据库工作簿发生变化时的事件——窗口控件中列表框的刷新
        ''' <summary>
        ''' 项目中的数据库工作簿发生变化时的事件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="clearListBox_Progress">是否要清空项目文件中的施工进度列表框中的内容</param>
        ''' <remarks></remarks>
        Private Sub _WorkBookInProjectChanged(ByVal Sender As Object, ByVal FileContents As clsData_FileContents, _
                                              ByVal clearListBox_Progress As Boolean) Handles Me.WorkBookInProjectChanged

            '' 整个项目文件中所有的工作簿中的所有工作表的集合，用来作为列表控件中DataSource。
            Dim lstAllWorksheets As New List(Of LstbxDisplayAndItem)
            '下面这一项“无”很重要，因为当数据库中对于某一个项目（如开挖剖面，测点坐标），
            '如果没有相应的数据，就应该选择“无”。
            lstAllWorksheets.Add(New LstbxDisplayAndItem(" 无", LstbxDisplayAndItem.NothingInListBox.None))

            '
            Me.CmbbxProgressWkbk.Items.Clear()
            '
            For Each lstItem As LstbxDisplayAndItem In Me.LstBxWorkbooks.DataSource
                Dim wkbk As Workbook = DirectCast(lstItem.Value, Workbook)

                '更新项目中的工作簿的组合框列表
                Me.CmbbxProgressWkbk.Items.Add(New LstbxDisplayAndItem(wkbk.Name, wkbk))

                '提取工作簿中的所有工作表
                For Each sht As Worksheet In wkbk.Worksheets
                    lstAllWorksheets.Add(New LstbxDisplayAndItem(DisplayedText:=wkbk.Name & " : " & sht.Name, _
                                                                      value:=sht))
                Next
            Next

            '更新： 将数据源更新到所有的列表框
            Call ListControlRefresh(lstAllWorksheets, FileContents)

            If clearListBox_Progress Then
                ' 项目文件中的施工进度工作表的列表框的刷新
                '这里应该判断项目文件中保存的施工进度工作表中，有哪些是位于被删除的那个工作簿的，然后将属于那个工作簿的工作表进行移除。
                '现在为了简单起见，先直接将其清空。
                With Me.LstbxSheetsProgressInProject
                    .DataSource = Nothing
                    '.Items.Clear()
                    '.DisplayMember = LstbxDisplayAndItem.DisplayMember
                End With
            End If

        End Sub
        ''' <summary>
        ''' 项目中的数据库工作簿发生变化时，更新窗口中的相关列表框中的数据对象
        ''' </summary>
        ''' <param name="lstAllsheet"></param>
        ''' <param name="FileContents"></param>
        ''' <remarks></remarks>
        Private Sub ListControlRefresh(ByVal lstAllsheet As List(Of LstbxDisplayAndItem), _
                                      ByVal FileContents As clsData_FileContents)
            Dim intSheetsCount As Short = lstAllsheet.Count
            Dim arrAllSheets1(0 To intSheetsCount - 1) As LstbxDisplayAndItem
            Dim arrAllSheets2(0 To intSheetsCount - 1) As LstbxDisplayAndItem
            Dim arrAllSheets3(0 To intSheetsCount - 1) As LstbxDisplayAndItem
            Dim arrAllSheets4(0 To intSheetsCount - 1) As LstbxDisplayAndItem
            '这里一定要生成副本，因为如果是同一个引用变量，那么设置到三个控件的DataSource属性中后，
            '如果一个列表组合框的选择项发生变化,那个另外两个控件的选择项也会同步变化。
            arrAllSheets1 = lstAllsheet.ToArray
            arrAllSheets2 = arrAllSheets1.Clone
            arrAllSheets3 = arrAllSheets1.Clone
            arrAllSheets4 = arrAllSheets1.Clone

            '设置各种列表框中的数据以及选择的项lstbxWorksheetsInProjectFileChanged
            With Me

                '开挖平面工作表
                With CmbbxPlan
                    .DataSource = arrAllSheets1
                    .ValueMember = cstValueMember
                    If FileContents.Sheet_PlanView IsNot Nothing Then
                        Call Me.SheetToComboBox(CmbbxPlan, FileContents.Sheet_PlanView)
                    Else
                        .SelectedValue = LstbxDisplayAndItem.NothingInListBox.None
                    End If
                End With

                '测点坐标工作表
                With CmbbxPointCoordinates
                    .DataSource = arrAllSheets2
                    .ValueMember = cstValueMember
                    If FileContents.Sheet_PointCoordinates IsNot Nothing Then
                        Call Me.SheetToComboBox(CmbbxPointCoordinates, FileContents.Sheet_PointCoordinates)
                    Else
                        .SelectedValue = LstbxDisplayAndItem.NothingInListBox.None
                    End If
                End With

                '开挖剖面工作表
                With CmbbxSectional
                    .DataSource = arrAllSheets3
                    .ValueMember = cstValueMember
                    If FileContents.Sheet_Elevation IsNot Nothing Then
                        Call Me.SheetToComboBox(CmbbxSectional, FileContents.Sheet_Elevation)
                    Else
                        .SelectedValue = LstbxDisplayAndItem.NothingInListBox.None
                    End If
                End With

                '开挖工况工作表
                With CmbbxWorkingStage
                    .DataSource = arrAllSheets4
                    .ValueMember = cstValueMember
                    If FileContents.Sheet_WorkingStage IsNot Nothing Then
                        Call Me.SheetToComboBox(CmbbxWorkingStage, FileContents.Sheet_WorkingStage)
                    Else
                        .SelectedValue = LstbxDisplayAndItem.NothingInListBox.None
                    End If
                End With

            End With

            '为施工进度列表服务的组合列表框
            With CmbbxProgressWkbk
                If .Items.Count = 0 Then
                    '工作簿中的工作表列表框
                    Me.LstbxSheetsProgressInWkbk.DataSource = Nothing
                    '上面将DataSource设置为Nothing会清空DisplayMember属性的值，那么下次再向列表框中添加成员时，
                    '其DisplayMember就为Nothing了，所以必须在下面设置好其DisplayMember的值。
                    Me.LstbxSheetsProgressInWkbk.DisplayMember = LstbxDisplayAndItem.DisplayMember
                Else
                    .SelectedIndex = 0
                End If
            End With

        End Sub
        ''' <summary>
        ''' 将项目文件中的工作表添加到组合列表框中
        ''' </summary>
        ''' <param name="cmbx">进行添加的组合列表框</param>
        ''' <param name="destinationSheet">要添加到组合列表框中的工作表对象</param>
        ''' <remarks></remarks>
        Private Sub SheetToComboBox(ByVal cmbx As ComboBox, ByVal destinationSheet As Worksheet)
            Dim lstbxItem As LstbxDisplayAndItem
            Dim sht As Worksheet
            With cmbx
                For Each lstbxItem In .Items
                    If Not lstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                        '有可能会出现列表中的项目不能转换为工作表对象的情况，比如第一项""
                        sht = DirectCast(lstbxItem.Value, Worksheet)
                        If ExcelFunction.SheetCompare(sht, destinationSheet) Then
                            .SelectedItem = lstbxItem
                            Exit For
                        End If
                    End If
                Next
            End With
        End Sub

#End Region

#Region "  ---  一般界面操作"

        '添加或者移除工作簿
        ''' <summary>
        ''' 在项目中添加数据库的工作簿
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnAddWorkbook_Click(sender As Object, e As EventArgs) Handles btnAddWorkbook.Click
            Dim FilePath As String = ""
            With APPLICATION_MAINFORM.MainForm.OpenFileDialog1
                .Title = "选择Excel数据工作簿"
                .Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb"
                .FilterIndex = 1
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    FilePath = .FileName
                End If
            End With
            If FilePath.Length > 0 Then
                Dim wkbk As Workbook = Nothing
                '先看选择的工作簿是否已经在数据库的Excel程序中打开
                Dim blnOpenedInApp As Boolean = False
                For Each wkbkOpened As Workbook In GlbApp.ExcelApplication_DB.Workbooks
                    If String.Compare(wkbkOpened.FullName, FilePath) = 0 Then
                        wkbk = wkbkOpened
                        blnOpenedInApp = True
                        Exit For
                    End If
                Next

                '
                If Not blnOpenedInApp Then      '如果此工作簿还没有在Excel程序中打开
                    '则将其打开，并添加到列表框中
                    wkbk = GlbApp.ExcelApplication_DB.Workbooks.Open(Filename:=FilePath, _
                                                                                 UpdateLinks:=False, _
                                                                                 [ReadOnly]:=True)
                    '为列表框中添加新元素
                    With LstBxWorkbooks
                        Dim DataSource As New List(Of LstbxDisplayAndItem)
                        If .DataSource IsNot Nothing Then
                            For Each i As LstbxDisplayAndItem In .DataSource
                                DataSource.Add(i)
                            Next
                        End If
                        DataSource.Add(New LstbxDisplayAndItem(FilePath, wkbk))
                        .DataSource = DataSource
                    End With
                    '
                    RaiseEvent WorkBookInProjectChanged(btnAddWorkbook, Me.F_NewFileContents, False)

                Else        '说明此工作簿已经在Excel中打开
                    '则先检查它是否已经添加到了列表框中
                    Dim lstbxItem As LstbxDisplayAndItem
                    Dim blnShownInListbox As Boolean = False
                    For Each lstbxItem In LstBxWorkbooks.Items
                        Dim wkbkInproject As Workbook = DirectCast(lstbxItem.Value, Workbook)
                        If String.Compare(wkbkInproject.FullName, FilePath) = 0 Then
                            blnShownInListbox = True
                        End If
                    Next
                    If Not blnShownInListbox Then
                        '为列表框中添加新元素
                        With LstBxWorkbooks
                            Dim DataSource As New List(Of LstbxDisplayAndItem)
                            If .DataSource IsNot Nothing Then
                                For Each i As LstbxDisplayAndItem In .DataSource
                                    DataSource.Add(i)
                                Next
                            End If
                            DataSource.Add(New LstbxDisplayAndItem(FilePath, wkbk))
                            .DataSource = DataSource
                        End With
                        '
                        RaiseEvent WorkBookInProjectChanged(btnAddWorkbook, Me.F_NewFileContents, False)

                    End If
                End If

                '再看将工作簿对象添加到列表中
            End If
        End Sub
        ''' <summary>
        ''' 在项目中移除数据库工作簿
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks>在从列表框中移除项目时，一定要注意，由于集合的数据结构，所以删除时要从后面的开始删除，
        ''' 否则可能会出现索引不到的情况。</remarks>
        Private Sub BtnRemoveWorkbook_Click(sender As Object, e As EventArgs) Handles BtnRemoveWorkbook.Click
            With LstBxWorkbooks
                Dim count As Byte = .SelectedIndices.Count
                If count > 0 Then

                    Dim DataSource As New List(Of LstbxDisplayAndItem)
                    If .DataSource IsNot Nothing Then
                        For Each i As LstbxDisplayAndItem In .DataSource
                            DataSource.Add(i)
                        Next
                    End If
                    '
                    Dim wkbk As Workbook, lstbxItem As LstbxDisplayAndItem
                    Dim index As Byte
                    For i As Short = count - 1 To 0 Step -1
                        index = .SelectedIndices(i)
                        lstbxItem = DataSource.Item(index)
                        '
                        wkbk = DirectCast(lstbxItem.Value, Workbook)
                        wkbk.Close(False)
                        DataSource.RemoveAt(index)
                    Next
                    .DataSource = DataSource
                    RaiseEvent WorkBookInProjectChanged(BtnRemoveWorkbook, Me.F_NewFileContents, True)
                End If
            End With
        End Sub

        '对施工进度表的列表项进行操作
        ''' <summary>
        ''' 在项目中添加施工进度工作表
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub BtnAddSheet_Click(sender As Object, e As EventArgs) Handles BtnAddSheet.Click

            Select Case Me.P_ProjectState

                Case Miscellaneous.ProjectState.NewProject       '直接用LstbxDisplayAndItem对象来进行比较

                    '项目文件中的施工进度工作表的列表框中的所有工作表项
                    Dim OldDataSource As Object = Me.LstbxSheetsProgressInProject.DataSource
                    Dim NewDataSource As New List(Of LstbxDisplayAndItem)
                    If Me.LstbxSheetsProgressInProject.DataSource IsNot Nothing Then
                        For Each i As LstbxDisplayAndItem In OldDataSource
                            NewDataSource.Add(i)
                        Next
                    End If
                    '
                    For Each lstItem_Wkbk As LstbxDisplayAndItem In Me.LstbxSheetsProgressInWkbk.SelectedItems
                        '看选择的工作表是否已经包含在项目工作表中
                        If Not NewDataSource.Contains(lstItem_Wkbk) Then
                            NewDataSource.Add(lstItem_Wkbk)
                        End If
                    Next
                    Me.LstbxSheetsProgressInProject.DataSource = NewDataSource

                Case Miscellaneous.ProjectState.EditProject      '用工作表的路径来进行比较
                    '工作簿中选择的工作表
                    Dim sht_InWkbk As Worksheet
                    '项目文件中已经存在的工作表
                    Dim sht_InProject As Worksheet
                    '
                    Dim ItemsAddToProject As New List(Of LstbxDisplayAndItem)
                    Dim OldDataSource_SheetsInProject As Object = Me.LstbxSheetsProgressInProject.DataSource
                    Dim NewDataSource_SheetsInProject As New List(Of LstbxDisplayAndItem)
                    '
                    If OldDataSource_SheetsInProject Is Nothing Then
                        For Each lstbxItem_Wkbk As LstbxDisplayAndItem In Me.LstbxSheetsProgressInWkbk.SelectedItems
                            NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk)
                        Next
                        Me.LstbxSheetsProgressInProject.DataSource = NewDataSource_SheetsInProject

                    Else

                        For Each lstbxItem_Wkbk As LstbxDisplayAndItem In OldDataSource_SheetsInProject
                            NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk)
                        Next
                        '
                        For Each lstbxItem_Wkbk As LstbxDisplayAndItem In Me.LstbxSheetsProgressInWkbk.SelectedItems
                            '判断两个工作表是否相等
                            Dim blnSheetsMatched As Boolean = False
                            '
                            sht_InWkbk = DirectCast(lstbxItem_Wkbk.Value, Worksheet)
                            For Each lstbxItem_Project As LstbxDisplayAndItem In OldDataSource_SheetsInProject
                                sht_InProject = DirectCast(lstbxItem_Project.Value, Worksheet)

                                If ExcelFunction.SheetCompare(sht_InProject, sht_InWkbk) = True Then
                                    blnSheetsMatched = True
                                    Exit For
                                End If
                            Next    'lstbxItem_Project

                            '如果两个工作表不匹配，则添加到项目文件中。
                            If Not blnSheetsMatched Then
                                NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk)
                            End If

                        Next        'lstbxItem_Wkbk
                        Me.LstbxSheetsProgressInProject.DataSource = NewDataSource_SheetsInProject
                    End If

            End Select
        End Sub
        ''' <summary>
        ''' 在项目中移除施工进度工作表
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub BtnRemoveSheet_Click(sender As Object, e As EventArgs) Handles BtnRemoveSheet.Click
            With LstbxSheetsProgressInProject
                Dim DataSource As New List(Of LstbxDisplayAndItem)
                If .DataSource IsNot Nothing Then
                    For Each i As LstbxDisplayAndItem In .DataSource
                        DataSource.Add(i)
                    Next
                End If
                Dim count As Byte = .SelectedIndices.Count
                For i As Short = count - 1 To 0 Step -1
                    Dim index As Byte = .SelectedIndices.Item(i)
                    DataSource.RemoveAt(index)
                Next
                .DataSource = DataSource
            End With
        End Sub
        ''' <summary>
        ''' 工作簿的组合列表框的选择项发生变化时，更新施工进度工作表的列表框
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CmbbxProgressWkbk_SelectedValueChanged(sender As Object, e As EventArgs) Handles CmbbxProgressWkbk.SelectedValueChanged
            Dim wkbk As Workbook
            Dim lstItem As LstbxDisplayAndItem = CmbbxProgressWkbk.SelectedItem
            If lstItem Is Nothing Then
                Me.LstbxSheetsProgressInProject.DataSource = Nothing
                Me.LstbxSheetsProgressInWkbk.DataSource = Nothing
            Else
                wkbk = DirectCast(lstItem.Value, Workbook)
                ''-- 这里的shtsObject只能声明为Object，而不能声明为Worksheets接口，否则会出现异常：
                '无法将类型为“System.__ComObject”的 COM 对象强制转换为接口类型
                '“Microsoft.Office.Interop.Excel.Worksheets”。此操作失败的原因是对 IID 
                '为“{000208B1-0000-0000-C000-000000000046}”的接口的 COM 组件调用 QueryInterface 
                '因以下错误而失败: 不支持此接口 (异常来自 HRESULT:0x80004002 (E_NOINTERFACE))。
                ''-- 但是这里一定要如此调用，而不用dim sht as worksheet=wkbk.Worksheets(i)去一次一次地
                '调用每一个单个的工作表，是为了避免多次调用Worksheets接口。
                Dim shtsObject As Object = wkbk.Worksheets
                Dim shtCount As Short = shtsObject.Count
                Dim arrSheets(0 To shtCount - 1) As LstbxDisplayAndItem
                For i = 0 To shtCount - 1
                    Dim sht As Worksheet = shtsObject(i + 1)    'Worksheets接口的集合中的第一个元素的下标值为1
                    arrSheets(i) = New LstbxDisplayAndItem(sht.Name, sht)
                Next
                LstbxSheetsProgressInWkbk.DataSource = arrSheets
            End If
        End Sub

        '窗口的最终处理：确定，取消等
        ''' <summary>
        ''' 将界面中的内容保存到XML文档中
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click

            ''提取界面中绑定的数据
            'Me.F_FileContents = Me.UIToFileContents()

            '新开一个线程以将FileContents中的内容更新到程序的DataBase中。
            Dim thd As New Thread(AddressOf Me.RefreshDataBase)
            With thd
                .Name = "在项目文件窗口关闭时刷新程序中的数据库"
                .Start(Me.F_NewFileContents)
            End With
            '
            Me.Close()
        End Sub
        Private Sub RefreshDataBase(ByVal FileContents As clsData_FileContents)
            With GlbApp
                With .ProjectFile
                    .Contents = FileContents
                    If Me.ProjectState = Miscellaneous.ProjectState.NewProject Then
                        .FilePath = Nothing
                    End If
                End With
                .DataBase = New ClsData_DataBase(FileContents)
            End With
            '将主程序的界面刷新为打开了文件后的界面
            APPLICATION_MAINFORM.MainForm.MainUI_ProjectOpened()
        End Sub

        ''' <summary>
        ''' 鼠标移动进Panel时引发的事件。
        ''' </summary>
        ''' <remarks>此时将Panel设置为获得焦点。</remarks>
        Private Sub PanelFather_MouseEnter(sender As Object, e As EventArgs) _
            Handles PanelFather.MouseEnter, LstbxSheetsProgressInProject.Click, LstbxSheetsProgressInWkbk.Click
            With Me.PanelFather
                .Focus()
            End With
        End Sub

#End Region

#Region "  ---  选择列表框内容时进行赋值"

        '开挖平面分块工作表
        Private Sub CmbbxPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbbxPlan.SelectedIndexChanged
            Dim LstbxItem As LstbxDisplayAndItem = CmbbxPlan.SelectedItem
            If LstbxItem IsNot Nothing Then
                If Not LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                    Me.F_NewFileContents.Sheet_PlanView = DirectCast(LstbxItem.Value, Worksheet)
                Else
                    Me.F_NewFileContents.Sheet_PlanView = Nothing
                End If
            End If
        End Sub
        '监测点位坐标的数据工作表
        Private Sub CmbbxPointCoordinates_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbbxPointCoordinates.SelectedIndexChanged
            Dim LstbxItem As LstbxDisplayAndItem = CmbbxPointCoordinates.SelectedItem
            If LstbxItem IsNot Nothing Then
                If Not LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                    Me.F_NewFileContents.Sheet_PointCoordinates = DirectCast(LstbxItem.Value, Worksheet)
                Else
                    Me.F_NewFileContents.Sheet_PointCoordinates = Nothing
                End If
            End If
        End Sub
        '开挖剖面标高图的数据工作表
        Private Sub CmbbxSectional_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbbxSectional.SelectedIndexChanged
            Dim LstbxItem As LstbxDisplayAndItem = CmbbxSectional.SelectedItem
            If LstbxItem IsNot Nothing Then
                If Not LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                    Me.F_NewFileContents.Sheet_Elevation = DirectCast(LstbxItem.Value, Worksheet)
                Else
                    Me.F_NewFileContents.Sheet_Elevation = Nothing
                End If
            End If
        End Sub
        '开挖工况的数据工作表
        Private Sub CmbbxWorkingStage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbbxWorkingStage.SelectedIndexChanged
            Dim LstbxItem As LstbxDisplayAndItem = CmbbxWorkingStage.SelectedItem
            If LstbxItem IsNot Nothing Then
                If Not LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                    Me.F_NewFileContents.Sheet_WorkingStage = DirectCast(LstbxItem.Value, Worksheet)
                Else
                    Me.F_NewFileContents.Sheet_WorkingStage = Nothing
                End If
            End If
        End Sub
        '施工进度工作表
        Private Sub LstbxSheetsProgressInProject_DataSourceChanged(sender As Object, e As EventArgs) Handles LstbxSheetsProgressInProject.DataSourceChanged
            Dim lstSheetProgress As New List(Of Worksheet)
            If LstbxSheetsProgressInProject.DataSource IsNot Nothing Then
                For Each LstbxItem As LstbxDisplayAndItem In LstbxSheetsProgressInProject.DataSource
                    Dim sht As Worksheet = DirectCast(LstbxItem.Value, Worksheet)
                    lstSheetProgress.Add(sht)
                Next
            End If

            Me.F_NewFileContents.lstSheets_Progress = lstSheetProgress
        End Sub
        '所有代表项目数据库的工作簿文件
        Private Sub LstBxWorkbooks_DataSourceChanged(sender As Object, e As EventArgs) Handles LstBxWorkbooks.DataSourceChanged
            '项目中所有的工作簿
            Dim lstWkbk As New List(Of Workbook)
            If LstBxWorkbooks.DataSource IsNot Nothing Then
                For Each LstbxItem As LstbxDisplayAndItem In LstBxWorkbooks.DataSource
                    lstWkbk.Add(DirectCast(LstbxItem.Value, Workbook))
                Next
            End If
            Me.F_NewFileContents.lstWkbks = lstWkbk
        End Sub

#End Region

    End Class
End Namespace
