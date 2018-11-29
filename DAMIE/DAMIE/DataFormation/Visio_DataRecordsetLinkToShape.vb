Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Miscellaneous.AdoForExcel
Imports DAMIE.Miscellaneous
''' <summary>
''' Excel数据到Visio形状
''' </summary>
''' <remarks></remarks>
Public Class Visio_DataRecordsetLinkToShape


#Region "  ---  Properties"
    ''' <summary>
    ''' 进行形状链接的文档
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents F_vsoDoc As Visio.Document
    ''' <summary>
    ''' 进行形状链接的文档，设置此属性时会触发vsoDocumentChanged事件
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property vsoDoc As Visio.Document
        Get
            Return F_vsoDoc
        End Get
        Set(value As Visio.Document)
            Me.F_vsoDoc = value
            RaiseEvent vsoDocumentChanged(value)
        End Set
    End Property

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 当进行数据链接的Visio文档发生改变时触发
    ''' </summary>
    ''' <remarks></remarks>
    Private Event vsoDocumentChanged As Action(Of Visio.Document)

    ''' <summary>
    ''' 在Visio文档通过验证，表示可以进行数据链接之时触发
    ''' </summary>
    ''' <remarks></remarks>
    Private Event ShapeIDValidated As Action

    ''' <summary>
    ''' Visio的Application对象，此对象不包含在“群坑分析”的主程序中的那个Visio的Application对象
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents F_vsoApplication As Visio.Application

    ''' <summary>
    ''' 进行形状链接的绘图页面
    ''' </summary>
    ''' <remarks></remarks>
    Private F_vsoPage As Visio.Page

    ''' <summary>
    ''' 进行链接的数据记录集
    ''' </summary>
    ''' <remarks></remarks>
    Private F_vsoDataRs As Visio.DataRecordset

    ''' <summary>
    ''' 在数据记录集中标识“形状ID”的字段列的下标值。在数据记录集中，每一行中的第一列（个）数据的下标值为0。
    ''' </summary>
    ''' <remarks></remarks>
    Private F_IndexOfShapeID As Integer

    Private F_arrCombobox(0 To 2) As ComboBox

#End Region

#Region "  ---  构造函数与窗体的加载"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        '设置组合列表框中要进行显示的属性
        Dim DisplayMember As String = LstbxDisplayAndItem.DisplayMember
        With Me
            F_arrCombobox(0) = ComboBox_Page
            F_arrCombobox(1) = ComboBox_DataRs
            F_arrCombobox(2) = ComboBox_Column_ShapeID
            '
            For Each cbb As ComboBox In Me.F_arrCombobox
                cbb.DisplayMember = DisplayMember
            Next
            '
            Me.btnLink.Enabled = False
            '
        End With

    End Sub

    Private Sub Visio_DataRecordsetLinkToShape_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '如果程序中已经有打开的Visio文档，则将此文档作为默认的进行形状链接的文档
        Dim GlobalApp As GlobalApplication = GlobalApplication.Application
        If GlobalApp IsNot Nothing Then
            With GlobalApp
                If .PlanView_VisioWindow IsNot Nothing Then
                    Dim doc As Visio.Document = .PlanView_VisioWindow.Page.Document
                    Me.vsoDoc = doc
                    Me.txtbxVsoDoc.Text = doc.FullName
                End If
            End With
        End If

    End Sub

#End Region

#Region "  ---  获取集合中的成员"
    ''' <summary>
    ''' 从Visio文档中返回其中的所有Page对象的数组
    ''' </summary>
    ''' <param name="Doc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPagesFromDoc(ByVal Doc As Visio.Document) As LstbxDisplayAndItem()
        Dim pagesCount As Short = Doc.Pages.Count
        Dim arrItems(0 To pagesCount - 1) As LstbxDisplayAndItem
        Dim i As Short = 0
        For Each page As Visio.Page In Doc.Pages
            arrItems(i) = New LstbxDisplayAndItem(page.Name, page)
            i += 1
        Next
        Return arrItems
    End Function

    ''' <summary>
    ''' 从Visio文档中返回其中的所有DataRecordset对象的数组
    ''' </summary>
    ''' <param name="Doc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDataRsFromDoc(ByVal Doc As Visio.Document) As LstbxDisplayAndItem()
        Dim DRSsCount As Short = Doc.DataRecordsets.Count
        Dim arrItems(0 To DRSsCount - 1) As LstbxDisplayAndItem
        Dim i As Short = 0
        For Each DRS As Visio.DataRecordset In Doc.DataRecordsets
            arrItems(i) = New LstbxDisplayAndItem(DRS.Name, DRS)
            i += 1
        Next
        Return arrItems
    End Function

    ''' <summary>
    ''' 从Visio文档的数据记录集中返回其中的字段列对象的数组
    ''' </summary>
    ''' <param name="DRS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetColumnsFromDataRS(ByVal DRS As Visio.DataRecordset) As LstbxDisplayAndItem()
        Dim ColumnsCount As Integer = DRS.DataColumns.Count
        Dim arrItems(0 To ColumnsCount - 1) As LstbxDisplayAndItem
        Dim i As Integer = 0
        For Each Column As Visio.DataColumn In DRS.DataColumns
            '在数据记录集中，第一列数据的Index为0。
            arrItems(i) = New LstbxDisplayAndItem(Column.DisplayName, i)
            i += 1
        Next
        Return arrItems
    End Function

#End Region

#Region "  ---  组合框的选择项发生改变"

    Private Sub ComboBox_DataRs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_DataRs.SelectedIndexChanged
        Dim lstItem As LstbxDisplayAndItem = Me.ComboBox_DataRs.SelectedItem
        Try
            Dim DataRs As Visio.DataRecordset = DirectCast(lstItem.Value, Visio.DataRecordset)
            Me.F_vsoDataRs = DataRs
            '更新数据记录集中的字段列。
            Me.ComboBox_Column_ShapeID.DataSource = GetColumnsFromDataRS(DataRs)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "选择数据记录集出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Me.btnLink.Enabled = False
    End Sub

    Private Sub ComboBox_Page_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Page.SelectedIndexChanged
        Dim lstItem As LstbxDisplayAndItem = Me.ComboBox_Page.SelectedItem
        Try
            Me.F_vsoPage = DirectCast(lstItem.Value, Visio.Page)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "选择Visio页面出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Me.btnLink.Enabled = False
    End Sub

    Private Sub ComboBox_Column_ShapeID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Column_ShapeID.SelectedIndexChanged
        Dim lstItem As LstbxDisplayAndItem = Me.ComboBox_Column_ShapeID.SelectedItem
        Try
            Me.F_IndexOfShapeID = DirectCast(lstItem.Value, Integer)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "选择形状ID的字段出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Me.btnLink.Enabled = False
    End Sub

#End Region

#Region "  ---  按钮操作"
    '验证
    Private Sub btnValidate_Click(sender As Object, e As EventArgs) Handles btnValidate.Click
        If ValidateShapes(Me.F_vsoPage, Me.F_vsoDataRs, Me.F_IndexOfShapeID) Then
            MessageBox.Show("形状ID验证成功！", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            RaiseEvent ShapeIDValidated()
        Else
            Me.btnLink.Enabled = False
        End If
    End Sub
    '链接
    Private Sub btnLink_Click(sender As Object, e As EventArgs) Handles btnLink.Click
        If PassDataRecordsetToShape(Me.F_vsoDataRs, Me.F_vsoPage) Then
            MessageBox.Show("形状ID数据链接到形状成功！", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            MessageBox.Show("形状ID数据链接到形状失败！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

#End Region

    '选择新的Visio文档
    ''' <summary>
    ''' 选择新的Visio文档
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnChooseVsoDoc_Click(sender As Object, e As EventArgs) Handles BtnChooseVsoDoc.Click

        Dim FilePath As String = ""
        With Me.OpenFileDialog1
            .Title = "选择进行数据链接的Visio文档"
            .Filter = "Visio文件(*.vsd)|*.vsd"
            .FilterIndex = 2
            .Multiselect = False
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FilePath = .FileName
            Else
                Exit Sub
            End If
        End With
        If FilePath.Length > 0 Then
            '
            Me.txtbxVsoDoc.Text = FilePath
            '
            If Me.F_vsoApplication Is Nothing Then
                Me.F_vsoApplication = New Visio.Application
            End If
            '
            With Me.F_vsoApplication
                Try
                    Me.vsoDoc = .Documents.Open(FilePath)
                Catch ex As Exception
                    MessageBox.Show("Visio文档打开出错，请检查后重新打开。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                .Visible = True
            End With
        End If

    End Sub

    '验证
    ''' <summary>
    ''' 验证页面中是否包含所有位于数据记录集中所记录的形状ID。
    ''' </summary>
    ''' <param name="page"></param>
    ''' <param name="DRS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateShapes(ByVal page As Visio.Page, ByVal DRS As Visio.DataRecordset, ByVal intIndexOfShapeID As Integer) As Boolean
        Dim blnValidated As Boolean = True
        If DRS IsNot Nothing Then
            Dim lngRowIDs() As Integer = DRS.GetDataRowIDs("")
            Dim shp As Visio.Shape
            For Each id As Integer In lngRowIDs
                Dim RowData As Object() = DRS.GetRowData(id)
                Dim shapeID As Object = RowData(intIndexOfShapeID)
                Try
                    shp = page.Shapes.ItemFromID(shapeID)
                Catch ex As Exception
                    blnValidated = False
                    Dim Result = MessageBox.Show("在页面中没有找到与形状ID """ & shapeID & """ 相匹配的形状。请仔细检查记录的形状ID值。", _
                                    "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
                    Select Case Result
                        Case Windows.Forms.DialogResult.OK

                        Case Windows.Forms.DialogResult.Cancel
                            '不再提示这条错误
                            Exit For
                    End Select
                End Try
            Next
        Else
            blnValidated = False
        End If
        Return blnValidated
    End Function

    '链接
    ''' <summary>
    ''' 将Visio中的外部数据链接到Page中的指定形状。
    ''' 此操作的作用：通过Visio的数据图形功能，在对应的形状上显示出它所链接的数据，比如此图形对应的开挖深度。
    ''' </summary>
    ''' <param name="DataRS">数据链接的源数据记录集</param>
    ''' <param name="Page">要进行数据链接的形状所在的Page</param>
    ''' <param name="ColumnIndex_PrimaryKey">在数据记录集中，用来记录形状的名称的数据所在的列号。如果是第一列，则为0.</param>
    ''' <param name="DeleteDataRecordset">是否要在数据记录集的数据链接到形状后，将此数据记录集删除。</param>
    ''' <remarks></remarks>
    Private Function PassDataRecordsetToShape(ByVal DataRS As Visio.DataRecordset, ByVal Page As Visio.Page, _
                                         Optional ByVal ColumnIndex_PrimaryKey As Short = 0, _
                                         Optional ByVal DeleteDataRecordset As Boolean = False)
        Dim blnSucceeded As Boolean = True
        Dim IDs() As Integer
        '  ------------------ GetDataRowIDs --------------------- 
        '获取数据记录集内所有行的 ID 组成的数组，其中每一行均代表一个数据记录。
        '若要不应用筛选器（即获取所有行），则传递一个空字符串 ("") 即可。
        IDs = DataRS.GetDataRowIDs("")
        '
        Dim shp As Visio.Shape
        Try
            For Each RowID As Integer In IDs
                Dim shapeID As Integer = DataRS.GetRowData(RowID)(ColumnIndex_PrimaryKey)
                'ItemFromID可以进行页面或者形状集合中的全局索引，即可以索引子形状中的嵌套形状，而Item一般只能索引其下的子形状。
                shp = Page.Shapes.ItemFromID(shapeID)
                shp.LinkToData(DataRS.ID, RowID, False)
            Next
        Catch ex As Exception
            blnSucceeded = False
        End Try

        '是否要在数据记录集的数据链接到形状后，将此数据记录集删除。
        If DeleteDataRecordset Then
            DataRS.Delete()
            DataRS = Nothing
        End If
        Return blnSucceeded
    End Function

#Region "  ---  用户定义的事件"

    ''' <summary>
    ''' Visio文档改变
    ''' </summary>
    ''' <param name="vsoDoc"></param>
    ''' <remarks></remarks>
    Private Sub DocumentChanged(ByVal vsoDoc As Visio.Document) Handles Me.vsoDocumentChanged
        '
        Me.ComboBox_Page.DataSource = GetPagesFromDoc(vsoDoc)
        Me.ComboBox_DataRs.DataSource = GetDataRsFromDoc(vsoDoc)
        '
        Me.btnValidate.Enabled = True
        Me.btnLink.Enabled = False
        '
    End Sub

    ''' <summary>
    ''' 形状ID验证成功
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Visio_DataRecordsetLinkToShape_ShapeIDValidated() Handles Me.ShapeIDValidated
        Me.btnLink.Enabled = True
    End Sub

    ''' <summary>
    ''' Visio程序关闭
    ''' </summary>
    ''' <param name="app"></param>
    ''' <remarks></remarks>
    Private Sub F_vsoApplication_BeforeQuit(app As Visio.Application) Handles F_vsoApplication.BeforeQuit
        Me.F_vsoApplication = Nothing
    End Sub

    ''' <summary>
    ''' Visio文档关闭
    ''' </summary>
    ''' <param name="Doc"></param>
    ''' <remarks></remarks>
    Private Sub F_vsoDoc_BeforeDocumentClose(Doc As Visio.Document) Handles F_vsoDoc.BeforeDocumentClose
        Me.F_vsoDoc = Nothing
        Me.F_vsoDataRs = Nothing
        '
        Call ClearUI()
    End Sub

    ''' <summary>
    ''' 委托：在主程序界面上清空列表框的显示
    ''' </summary>
    ''' <remarks></remarks>
    Private Delegate Sub BeforeDocumentCloseHander()
    ''' <summary>
    ''' 在主程序界面上清空列表框的显示
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearUI()
        With Me
            If .InvokeRequired Then
                '非UI线程，再次封送该方法到UI线程
                Me.BeginInvoke(New BeforeDocumentCloseHander(AddressOf Me.ClearUI))
            Else
                Me.txtbxVsoDoc.Text = ""
                Me.btnValidate.Enabled = False
                Me.btnLink.Enabled = False
                For Each cbb As ComboBox In Me.F_arrCombobox
                    Try
                        With cbb
                            .DataSource = Nothing
                            .DisplayMember = LstbxDisplayAndItem.DisplayMember
                        End With
                    Catch ex As Exception
                        Debug.Print("重新设置列表框的数据源出错！")
                    End Try
                Next
            End If
        End With
    End Sub

#End Region


End Class