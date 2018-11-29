Imports Microsoft.Office.Interop
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.Miscellaneous
Imports DAMIE.GlobalApp_Form
Imports DAMIE.All_Drawings_In_Application.ClsDrawing_PlanView
Imports DAMIE.Constants.xmlNodeNames.VisioPlanView_MonitorPoints
Imports System.IO
Imports System.Xml
''' <summary>
''' 绘制开挖平面图图的窗口界面
''' </summary>
''' <remarks></remarks>
Public Class frmDrawingPlan

#Region "  ---  Fields"

    Private F_HasMonitorPointinfos As Boolean

    Private MonitorPointinfos As MonitorPointsInformation
    ''' <summary>
    ''' 整个项目的文件路径
    ''' </summary>
    ''' <remarks></remarks>
    Private ProjectFilePath As String

#End Region

#Region "  ---  窗口的加载与关闭"
    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        With Me
            .ChkBx_PointInfo.Checked = False
            .ChkBx_PointInfo_CheckedChanged(Nothing, Nothing)
        End With
        '设置监测点位信息的初始数据
        Me.ProjectFilePath = GlobalApplication.Application.ProjectFile.FilePath
        Call ImportFromXmlFile(Me.ProjectFilePath)
        '
        'Dim settings As New mySettings_Application
        'Call PointsInfoToUI(settings.MonitorPointsInfo)

    End Sub

    ''' <summary>
    ''' 窗口关闭
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

#End Region

    ''' <summary>
    ''' 打开新的Visio的开挖平面图。如果平面图已经打开，则不能再打开新的平面图。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnChooseVisioPlanView_Click(sender As Object, e As EventArgs) _
        Handles btnChooseVisioPlanView.Click
        Dim VisioFilepath As String = String.Empty
        Dim OpenFileDialg As New OpenFileDialog
        With OpenFileDialg
            .Title = "选择基坑开挖平面图"
            .Filter = "Visio Documents  (*.vsd)|*.vsd"
            .FilterIndex = 1
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                VisioFilepath = .FileName
            End If
        End With
        Me.TextBoxFilePath.Text = VisioFilepath
    End Sub

    ''' <summary>
    ''' 生成开挖平面图
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ConstructVisioPlanView(sender As Object, e As EventArgs) Handles BtnGenerate.Click
        '检查程序中是否已经有了打开的Visio绘图
        If GlobalApplication.Application.PlanView_VisioWindow Is Nothing Then
            Dim VsoFilepath As String = Me.TextBoxFilePath.Text
            Dim blnFilePathValidated As Boolean = False
            If VsoFilepath.Length > 0 Then
                If File.Exists(VsoFilepath) Then
                    If Path.GetExtension(VsoFilepath) = ".vsd" Then
                        blnFilePathValidated = True
                    End If
                End If
            End If
            If blnFilePathValidated Then
                Try
                    Me.Hide()
                    Dim PointsInfo As ClsDrawing_PlanView.MonitorPointsInformation = Nothing

                    '提取监测点位的信息
                    If F_HasMonitorPointinfos Then
                        PointsInfo = UIToPointsInfo()
                    End If
                    '提取开挖平面图的信息
                    Dim visioWindow As New ClsDrawing_PlanView(
                                            strFilePath:=VsoFilepath, type:=DrawingType.Vso_PlanView,
                                            PageName_PlanView:=Me.TextBoxPageName.Text,
                                            ShapeID_AllRegions:=Me.TextBoxAllRegions.Text,
                                            InfoBoxID:=Me.TextBoxInfoBoxID.Text,
                                            HasMonitorPointsInfo:=Me.F_HasMonitorPointinfos, MonitorPointsInfo:=PointsInfo)
                    Me.Close()
                Catch ex As Exception
                    MessageBox.Show("Visio平面图打开出错，请重新打开。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    Me.Visible = True
                    GlobalApplication.Application.PlanView_VisioWindow = Nothing
                End Try
            Else
                MessageBox.Show("Visio文档不符合规范，请重新选择。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            End If
        Else
            '不能打开多个Visio平面图
            MessageBox.Show("Visio平面图已经打开。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End If
    End Sub

    ''' <summary>
    ''' 启用或禁用监测点位信息的设置区域
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ChkBx_PointInfo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBx_PointInfo.CheckedChanged
        Select Case Me.ChkBx_PointInfo.CheckState
            Case CheckState.Checked
                Me.Panel1.Enabled = True
                Me.F_HasMonitorPointinfos = True
            Case CheckState.Unchecked
                Me.Panel1.Enabled = False
                Me.F_HasMonitorPointinfos = False
        End Select
    End Sub

#Region "  ---  一般界面操作"
    '文本框的字符格式验证
    Private Sub ValidateForSingle(sender As Object, e As EventArgs) _
        Handles txtbx_Pt_BL_CAD_X.Validating, txtbx_Pt_BL_CAD_Y.Validating,
        txtbx_Pt_UR_CAD_X.Validating, txtbx_Pt_UR_CAD_Y.Validating
        Dim ctrl As TextBox = DirectCast(sender, TextBox)
        Dim T As String = ctrl.Text
        Try
            Dim Coordinate As Single = CSng(T)
            With ctrl
                .Text = T.TrimStart({"0"c})
            End With
        Catch ex As Exception
            With ctrl
                .Text = "0.0"
            End With
        End Try
    End Sub
    Private Sub ValidateForInteger(sender As Object, e As EventArgs) Handles txtbx_Pt_BL_ShapeID.Validating, txtbx_Pt_UR_ShapeID.Validating
        Dim ctrl As TextBox = DirectCast(sender, TextBox)
        Dim T As String = ctrl.Text
        Try
            Dim Coordinate As Integer = CInt(T)
            With ctrl
                .Text = T.TrimStart({"0"c})
            End With
        Catch ex As Exception
            With ctrl
                .Text = "0"
            End With
        End Try
    End Sub
    '数据的导入与导出
    Private Sub Btn_Import_Click(sender As Object, e As EventArgs) Handles Btn_Import.Click
        '这里要再索引一次，是为了避免在此窗口被打开的过程中，更新了项目文件，
        '而如果这里不再次索引，那Me.ProjectFilePath就还是原来的那个未更新的项目文件。
        Me.ProjectFilePath = GlobalApplication.Application.ProjectFile.FilePath
        Call ImportFromXmlFile(Me.ProjectFilePath)
    End Sub
    Private Sub Btn_Export_Click(sender As Object, e As EventArgs) Handles Btn_Export.Click
        Dim xmlDoc As New XmlDocument()
        If Me.ProjectFilePath Is Nothing Then
            '用打开文件对话框选择 .ame 项目文件
        End If
        xmlDoc.Load(Me.ProjectFilePath)
        Dim eleRoot As XmlElement = xmlDoc.SelectSingleNode(My.Settings.ProjectName)
        Dim ElePtInfo As XmlElement = eleRoot.SelectSingleNode(Nd1_MonitorPoints)
        If ElePtInfo Is Nothing Then ElePtInfo = eleRoot.AppendChild(xmlDoc.CreateElement(Nd1_MonitorPoints))
        '
        Me.MonitorPointinfos = UIToPointsInfo()
        Call ExportToXmlFile(ElePtInfo, Me.MonitorPointinfos)
        '
        xmlDoc.Save(Me.ProjectFilePath)
        MessageBox.Show("导出到项目文件成功！", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.None)
    End Sub
    '
    ''' <summary>
    ''' 将监测点位信息的属性值显示在窗口界面中
    ''' </summary>
    ''' <param name="PointsInfo"></param>
    ''' <remarks></remarks>
    Private Sub PointsInfoToUI(ByVal PointsInfo As ClsDrawing_PlanView.MonitorPointsInformation)
        With PointsInfo
            Me.txtbx_ShapeName_MonitorPointTag.Text = .ShapeName_MonitorPointTag
            Me.txtbx_Pt_BL_ShapeID.Text = .pt_Visio_BottomLeft_ShapeID
            Me.txtbx_Pt_UR_ShapeID.Text = .pt_Visio_UpRight_ShapeID
            Me.txtbx_Pt_BL_CAD_X.Text = .pt_CAD_BottomLeft.X
            Me.txtbx_Pt_BL_CAD_Y.Text = .pt_CAD_BottomLeft.Y
            Me.txtbx_Pt_UR_CAD_X.Text = .pt_CAD_UpRight.X
            Me.txtbx_Pt_UR_CAD_Y.Text = .pt_CAD_UpRight.Y
        End With
    End Sub
    ''' <summary>
    ''' 根据窗口界面中输入的监测点位数据，来返回对应的结构体属性。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UIToPointsInfo() As ClsDrawing_PlanView.MonitorPointsInformation
        Dim PointsInfo As New ClsDrawing_PlanView.MonitorPointsInformation()
        Try
            With PointsInfo
                .ShapeName_MonitorPointTag = Me.txtbx_ShapeName_MonitorPointTag.Text
                .pt_CAD_BottomLeft = New PointF(CSng(Me.txtbx_Pt_BL_CAD_X.Text), CSng(Me.txtbx_Pt_BL_CAD_Y.Text))
                .pt_CAD_UpRight = New PointF(CSng(Me.txtbx_Pt_UR_CAD_X.Text), CSng(Me.txtbx_Pt_UR_CAD_Y.Text))
                .pt_Visio_BottomLeft_ShapeID = CInt(Me.txtbx_Pt_BL_ShapeID.Text)
                .pt_Visio_UpRight_ShapeID = CInt(Me.txtbx_Pt_UR_ShapeID.Text)
            End With
        Catch ex As Exception
            MessageBox.Show("测点绘制与定位的数据格式不正确，请调整", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New ClsDrawing_PlanView.MonitorPointsInformation()
        End Try
        Return PointsInfo
    End Function

#End Region

#Region "  ---  Visio平面图中，测点的位置与标签信息的导入与导出"
    '导入
    ''' <summary>
    ''' 从项目文件中导入测点的数据
    ''' </summary>
    ''' <param name="Path_xmlDocFile"></param>
    ''' <remarks></remarks>
    Private Sub ImportFromXmlFile(ByVal Path_xmlDocFile As String)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(Path_xmlDocFile)
        '
        Dim eleRoot As XmlElement = xmlDoc.SelectSingleNode(My.Settings.ProjectName)
        Dim ElePtInfo As XmlElement = eleRoot.SelectSingleNode(Nd1_MonitorPoints)
        If ElePtInfo IsNot Nothing Then
            Me.MonitorPointinfos = ImportFromXmlElement(ElePtInfo)
        Else
            MessageBox.Show("项目文件中没有监测点位布置信息的数据")
        End If
        '
        Call PointsInfoToUI(Me.MonitorPointinfos)
    End Sub
    ''' <summary>
    ''' 从xml的节点中导入其子节点中保存的数据
    ''' </summary>
    ''' <param name="EleParent"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ImportFromXmlElement(ByVal EleParent As XmlElement) As MonitorPointsInformation
        Dim PointsInfo As New MonitorPointsInformation
        Try
            With PointsInfo
                .ShapeName_MonitorPointTag = EleParent.SelectSingleNode(Nd2_ShapeName_MonitorPointTag).InnerText
                .pt_Visio_BottomLeft_ShapeID = EleParent.SelectSingleNode(Nd2_pt_Visio_BottomLeft_ShapeID).InnerText
                .pt_Visio_UpRight_ShapeID = EleParent.SelectSingleNode(Nd2_pt_Visio_UpRight_ShapeID).InnerText
                .pt_CAD_BottomLeft = New PointF(EleParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_X).InnerText,
                                               EleParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_Y).InnerText)
                .pt_CAD_UpRight = New PointF(EleParent.SelectSingleNode(Nd2_pt_CAD_UpRight_X).InnerText,
                                               EleParent.SelectSingleNode(Nd2_pt_CAD_UpRight_Y).InnerText)
            End With
        Catch ex As Exception
            MessageBox.Show("从文件导入出错" & vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                            ex.TargetSite.Name, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return PointsInfo
    End Function
    '导出
    ''' <summary>
    ''' 将测点的绘制与定位数据保存到xml文档的某节点中。
    ''' </summary>
    ''' <param name="xmlParent"></param>
    ''' <param name="Pointinfos"></param>
    ''' <remarks></remarks>
    Private Sub ExportToXmlFile(ByVal xmlParent As XmlElement, ByVal Pointinfos As MonitorPointsInformation)
        Try
            Dim Doc As XmlDocument = xmlParent.OwnerDocument
            With Pointinfos
                '
                Dim Nd_ShapeName_MonitorPointTag As XmlElement
                Dim Nd_pt_CAD_BottomLeft_X As XmlElement
                Dim Nd_pt_CAD_BottomLeft_Y As XmlElement
                Dim Nd_pt_CAD_UpRight_X As XmlElement
                Dim Nd_pt_CAD_UpRight_Y As XmlElement
                Dim Nd_pt_Visio_BottomLeft_ShapeID As XmlElement
                Dim Nd_pt_Visio_UpRight_ShapeID As XmlElement
                '
                Nd_ShapeName_MonitorPointTag = xmlParent.SelectSingleNode(Nd2_ShapeName_MonitorPointTag)
                Dim blnHasChildNode As Boolean = False
                If Nd_ShapeName_MonitorPointTag IsNot Nothing Then
                    blnHasChildNode = True
                End If
                '后面认为：如果节点中没有子节点“ShapeName_MonitorPointTag”，则没有其他的子节点；
                '而如果有此节点, 则认为其他的子节点也都存在
                If blnHasChildNode Then
                    Nd_pt_CAD_BottomLeft_X = xmlParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_X)
                    Nd_pt_CAD_BottomLeft_Y = xmlParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_Y)
                    Nd_pt_CAD_UpRight_X = xmlParent.SelectSingleNode(Nd2_pt_CAD_UpRight_X)
                    Nd_pt_CAD_UpRight_Y = xmlParent.SelectSingleNode(Nd2_pt_CAD_UpRight_Y)
                    Nd_pt_Visio_BottomLeft_ShapeID = xmlParent.SelectSingleNode(Nd2_pt_Visio_BottomLeft_ShapeID)
                    Nd_pt_Visio_UpRight_ShapeID = xmlParent.SelectSingleNode(Nd2_pt_Visio_UpRight_ShapeID)
                Else
                    Nd_ShapeName_MonitorPointTag = xmlParent.AppendChild(Doc.CreateElement(Nd2_ShapeName_MonitorPointTag))
                    Nd_pt_CAD_BottomLeft_X = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_BottomLeft_X))
                    Nd_pt_CAD_BottomLeft_Y = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_BottomLeft_Y))
                    Nd_pt_CAD_UpRight_X = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_UpRight_X))
                    Nd_pt_CAD_UpRight_Y = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_UpRight_Y))
                    Nd_pt_Visio_BottomLeft_ShapeID = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_Visio_BottomLeft_ShapeID))
                    Nd_pt_Visio_UpRight_ShapeID = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_Visio_UpRight_ShapeID))
                End If
                Nd_ShapeName_MonitorPointTag.InnerText = .ShapeName_MonitorPointTag
                Nd_pt_CAD_BottomLeft_X.InnerText = .pt_CAD_BottomLeft.X
                Nd_pt_CAD_BottomLeft_Y.InnerText = .pt_CAD_BottomLeft.Y
                Nd_pt_CAD_UpRight_X.InnerText = .pt_CAD_UpRight.X
                Nd_pt_CAD_UpRight_Y.InnerText = .pt_CAD_UpRight.Y
                Nd_pt_Visio_BottomLeft_ShapeID.InnerText = .pt_Visio_BottomLeft_ShapeID
                Nd_pt_Visio_UpRight_ShapeID.InnerText = .pt_Visio_UpRight_ShapeID
                '
            End With
        Catch ex As Exception
            MessageBox.Show("导出到文件出错。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                            ex.TargetSite.Name, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class