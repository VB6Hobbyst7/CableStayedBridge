Imports System.IO
Imports System.Xml
Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Constants
Imports DAMIE.Constants.xmlNodeNames
Imports DAMIE.Miscellaneous
Imports DAMIE.GlobalApp_Form

Namespace DataBase

    ''' <summary>
    ''' 项目文件类，对应于每一个本地的项目文件。
    ''' 它主要实现XML文档中的内容与程序中的FileContents对象的交互。
    ''' </summary>
    ''' <remarks>此类并不实现与界面的UI交互。</remarks>
    Public Class clsProjectFile

#Region "  ---  属性值定义"

        ''' <summary>
        ''' 项目文件的路径
        ''' </summary>
        ''' <remarks></remarks>
        Private P_FilePath As String
        ''' <summary>
        ''' 项目文件的路径
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilePath As String
            Get
                Return Me.P_FilePath
            End Get
            Set(value As String)
                Me.P_FilePath = value
            End Set
        End Property

        ''' <summary>
        ''' 项目文件中记录的内容的实际对象
        ''' </summary>
        ''' <remarks></remarks>
        Private P_FileContents As clsData_FileContents
        ''' <summary>
        ''' ！项目文件中记录的内容的实际对象，即其中的Workbook对象、Worksheet对象等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Contents As clsData_FileContents
            Get
                Return Me.P_FileContents
            End Get
            Set(value As clsData_FileContents)
                Me.P_FileContents = value
            End Set

        End Property

#End Region

#Region "  ---  字段值定义"

        ''' <summary>
        ''' 整个程序中用来放置各种隐藏的Excel数据文档的Application对象
        ''' </summary>
        ''' <remarks></remarks>
        Private F_Application As Application

        ''' <summary>
        ''' 指示此项目文件是否是有效文件，即文件中的数据是否正常，文件中索引的工作簿或者工作表是否正常
        ''' </summary>
        ''' <remarks>只要有一个不正常，则为False</remarks>
        Private F_blnFileValid As Boolean

        Private F_lstErrorMessage As New List(Of String)
#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="xmlFilePath">与程序进行交互的XML文档的路径，如果不指定，则为空字符</param>
        ''' <remarks></remarks>
        Public Sub New(Optional ByVal xmlFilePath As String = "")
            Me.P_FilePath = xmlFilePath
            Me.F_Application = GlobalApplication.Application.ExcelApplication_DB
        End Sub

        '将项目中的内容写入XML文档
        ''' <summary>
        ''' 将设置好的项目写入文件
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SaveToXmlFile()
            Dim FileContents As clsData_FileContents = Me.P_FileContents
            Dim xmlDoc As New XmlDocument
            Dim sheet As Worksheet
            ' --- 写入根节点
            Dim eleRoot As XmlNode = xmlDoc.CreateElement(My.Settings.ProjectName)
            xmlDoc.AppendChild(eleRoot)
            Dim eleDataBase As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd1_DataBasePaths)
            eleRoot.AppendChild(eleDataBase)
            ' ------------ 写入整个项目的所有工作簿的绝对路径
            For Each wkbk As Workbook In FileContents.lstWkbks
                Dim eleWkbks As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_WorkbooksInProject)
                eleWkbks.InnerText = wkbk.FullName
                eleDataBase.AppendChild(eleWkbks)
            Next

            '--------- 写入施工进度工作表                       
            Dim iProgress As Short = 1
            For Each sheet In FileContents.lstSheets_Progress
                Dim eleProgress As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_Progress)
                Call WriteChildNodes(xmlDoc, eleProgress, sheet)
                eleDataBase.AppendChild(eleProgress)
            Next

            '--------- 写入开挖剖面工作表
            Dim eleSecional As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_SectionalView)
            eleDataBase.AppendChild(eleSecional)
            sheet = FileContents.Sheet_Elevation
            Call WriteChildNodes(xmlDoc, eleSecional, sheet)

            '-------- 写入测点坐标工作表
            Dim elePoint As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_PointCoordinates)
            eleDataBase.AppendChild(elePoint)
            sheet = FileContents.Sheet_PointCoordinates
            Call WriteChildNodes(xmlDoc, elePoint, sheet)


            '-------- 写入测点坐标工作表
            Dim eleWorkingStage As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_WorkingStage)
            eleDataBase.AppendChild(eleWorkingStage)
            sheet = FileContents.Sheet_WorkingStage
            Call WriteChildNodes(xmlDoc, eleWorkingStage, sheet)


            '-------- 写入开挖分块平面图
            Dim elePlan As XmlElement = xmlDoc.CreateElement(DataBasePath.Nd2_PlanView)
            eleDataBase.AppendChild(elePlan)
            sheet = FileContents.Sheet_PlanView
            Call WriteChildNodes(xmlDoc, elePlan, sheet)

            '保存文档
            xmlDoc.Save(Me.P_FilePath)
        End Sub
        ''' <summary>
        ''' 将每一个工作表项目写入XML文档中，此方法在ParentElement下创建两个子节点
        ''' </summary>
        ''' <param name="xmlDoc">写入节点的xml文档</param>
        ''' <param name="ParentElement">节点元素，要写入的子节点就是在此节点之下的</param>
        ''' <param name="sheet">要写入的Excel工作表</param>
        ''' <remarks>在此方法中，将指定工作表所在的工作簿的绝对路径，与此工作表的名称，作为两个子节点，
        ''' 写入到父节点ParentElement中。</remarks>
        Private Sub WriteChildNodes(ByVal xmlDoc As XmlDocument, ByVal ParentElement As XmlElement, ByVal sheet As Worksheet)
            If sheet IsNot Nothing Then
                With xmlDoc

                    '节点：工作簿路径
                    Dim wkbk As Workbook
                    Dim eleFilePath1 As XmlElement = .CreateElement(DataBasePath.Nd3_FilePath)
                    wkbk = sheet.Parent
                    eleFilePath1.InnerText = wkbk.FullName

                    '节点：工作表名称
                    Dim eleShtName As XmlElement = .CreateElement(DataBasePath.Nd3_SheetName)
                    eleShtName.InnerText = sheet.Name

                    '文件写入
                    ParentElement.AppendChild(eleFilePath1)
                    ParentElement.AppendChild(eleShtName)
                End With
            End If
        End Sub
        '从XML文件读取并检测文件中的成员是否存在
        ''' <summary>
        ''' 从项目文件中读取数据，并打开相应的Excel程序与工作簿
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub LoadFromXmlFile()
            '载入文档
            Dim xmlDoc As New XmlDocument
            xmlDoc.Load(Me.P_FilePath)
            '
            Dim FC As New clsData_FileContents
            Dim eleRoot As XmlNode = xmlDoc.SelectSingleNode(My.Settings.ProjectName)
            '这里可以尝试用GetElementById
            Dim Node_DataBase As XmlElement = eleRoot.SelectSingleNode(DataBasePath.Nd1_DataBasePaths)
            If Node_DataBase Is Nothing Then Exit Sub
            ' ---------------------- 读取文档 ------------------------ 
            ' ---------------------- 读取文档 ------------------------ 

            Dim eleWkbks As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_WorkbooksInProject)
            For Each eleWkbk As XmlElement In eleWkbks
                Dim strWkbkPath As String = eleWkbk.InnerText
                Dim wkbk As Workbook = ExcelFunction.MatchOpenedWkbk(strWkbkPath, Me.F_Application, OpenIfNotOpened:=True)
                If wkbk IsNot Nothing Then
                    FC.lstWkbks.Add(wkbk)
                Else     '此工作簿不存在，或者是没有成功赋值
                    Me.F_blnFileValid = False
                    Me.F_lstErrorMessage.Add("The Specified Workbook is not found : " & strWkbkPath)
                End If
            Next

            ' ---------------- 施工进度工作表
            Dim blnNodeForWorksheetValidated As Boolean
            Dim eleSheetsProgress As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_Progress)
            For Each eleSheetProgress As XmlElement In eleSheetsProgress
                Dim shtProgress As Worksheet = ValidateNodeForWorksheet(eleSheetProgress, FC, blnNodeForWorksheetValidated)
                If blnNodeForWorksheetValidated Then FC.lstSheets_Progress.Add(shtProgress)
            Next eleSheetProgress

            ' ---------------- 开挖平面图工作表

            Dim eleSheetPlanView As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_PlanView)
            Dim shtPlanView = ValidateNodeForWorksheet(eleSheetPlanView.Item(0), FC, blnNodeForWorksheetValidated)
            FC.Sheet_PlanView = shtPlanView

            ' ---------------- 开挖剖面图工作表

            Dim eleSheetSectionalView As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_SectionalView)
            Dim shtSectionalView As Worksheet = ValidateNodeForWorksheet(eleSheetSectionalView.Item(0), FC, blnNodeForWorksheetValidated)
            FC.Sheet_Elevation = shtSectionalView

            ' ---------------- 测点坐标工作表

            Dim eleSheetPointCoordinates As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_PointCoordinates)
            Dim shtPoint As Worksheet = ValidateNodeForWorksheet(eleSheetPointCoordinates.Item(0), FC, blnNodeForWorksheetValidated)
            FC.Sheet_PointCoordinates = shtPoint

            ' ---------------- 开挖工况工作表

            Dim eleWorkingStage As XmlNodeList = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_WorkingStage)
            Dim shtWorkingStage As Worksheet = ValidateNodeForWorksheet(eleWorkingStage.Item(0), FC, blnNodeForWorksheetValidated)
            FC.Sheet_WorkingStage = shtWorkingStage
            '
            Me.P_FileContents = FC
            '刷新主程序界面显示
            APPLICATION_MAINFORM.MainForm.MainUI_ProjectOpened()
        End Sub
        ''' <summary>
        ''' 检测工作表节点的有效性，此类节点中包含了两个子节点，一个是此工作表所在的工作簿的路径，一个是此工作表的名称；
        ''' 如果检测通过，则返回此Worksheet对象，否则返回Nothing
        ''' </summary>
        ''' <param name="WorksheetNode"></param>
        ''' <param name="FileContents">用来放置项目文件中记录的工作簿或者工作表对象的变量</param>
        ''' <param name="blnNodeForWorksheetValidated"></param>
        ''' <returns>要返回的工作表对象，如果验证不通过，则返回Nothing</returns>
        ''' <remarks></remarks>
        Private Function ValidateNodeForWorksheet(ByVal WorksheetNode As XmlElement, _
                                                  ByVal FileContents As clsData_FileContents, _
                                                  ByRef blnNodeForWorksheetValidated As Boolean) As Worksheet
            '要返回的工作表对象，如果验证不通过，则返回Nothing
            Dim ValidSheet As Worksheet = Nothing
            '
            blnNodeForWorksheetValidated = False
            '节点中记录的工作簿路径
            Dim strWkbkPath As String
            Dim ndWkbkPath As XmlNode = WorksheetNode.SelectSingleNode(DataBasePath.Nd3_FilePath)
            If ndWkbkPath Is Nothing Then   '说明此节点中没有记录工作表所在的工作簿信息，也就是说，此节点中没有记录值
                Return ValidSheet
            Else
                strWkbkPath = ndWkbkPath.InnerText
            End If

            '节点中记录的工作表名称
            Dim strSheetName As String
            Dim ndShetName As XmlNode = WorksheetNode.Item(DataBasePath.Nd3_SheetName)
            If ndShetName Is Nothing Then
                Return ValidSheet
            Else
                strSheetName = ndShetName.InnerText
            End If


            '---先检测工作表所在的工作簿是否在有效的并成功打开和返回的工作簿列表中
            Dim ValidWkbk As Workbook = Nothing
            For Each Wkbk As Workbook In FileContents.lstWkbks
                If String.Compare(strWkbkPath, Wkbk.FullName, True) = 0 Then
                    ValidWkbk = Wkbk
                    Exit For
                End If
            Next Wkbk

            '---- 根据工作簿的有效性与否执行相应的操作
            If ValidWkbk IsNot Nothing Then        '说明工作簿有效

                '开始检测工作表的有效性

                ValidSheet = ExcelFunction.MatchWorksheet(ValidWkbk, strSheetName)
                If ValidSheet IsNot Nothing Then
                    blnNodeForWorksheetValidated = True
                    '  
                Else     '此工作簿不存在，或者是没有成功赋值
                    Me.F_blnFileValid = False
                    Me.F_lstErrorMessage.Add("The Specified Worksheet" & strSheetName & "is not found in workbook: " & strWkbkPath)
                End If

            Else            '说明节点中记录的工作簿无效
                Me.F_blnFileValid = False
                Me.F_lstErrorMessage.Add("The Specified Workbook for worksheet " & strSheetName _
                                         & " is not found : " & strWkbkPath)
            End If

            '返回检测结果
            Return ValidSheet
        End Function

    End Class
End Namespace