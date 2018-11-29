Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.Constants.Data_Drawing_Format
Imports DAMIE.DataBase

Namespace DataBase

    ''' <summary>
    ''' 最原始的数据库类
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsData_DataBase

#Region "  ---  定义与声明"

#Region "  ---  Fields"

        ''' <summary>
        ''' 数据库中的剖面标高工作表
        ''' </summary>
        ''' <remarks></remarks>
        Private F_shtElevation As Worksheet


#End Region

#Region "  ---  Events"

        ''' <summary>
        ''' 在DataBase中的表示剖面标高的基坑ID属性被整体修改赋值时触发。
        ''' </summary>
        ''' <param name="dic_Data"></param>
        ''' <remarks></remarks>
        Public Shared Event dic_IDtoComponentsChanged(ByVal dic_Data As Dictionary(Of String, clsData_ExcavationID))

        ''' <summary>
        ''' 在DataBase中的表示基坑的施工进度属性被整体修改赋值时触发。
        ''' </summary>
        ''' <param name="ProcessRange"></param>
        ''' <remarks></remarks>
        Public Shared Event ProcessRangeChanged(ByVal ProcessRange As List(Of clsData_ProcessRegionData))

        ''' <summary>
        ''' 在DataBase中的表示基坑中的开挖工况汇总信息被整体修改赋值时触发。
        ''' </summary>
        ''' <param name="NewWorkingStage">以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
        ''' 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高</param>
        ''' <remarks></remarks>
        Public Shared Event WorkingStageChanged(ByVal NewWorkingStage As Dictionary(Of String, List(Of clsData_WorkingStage)))

#End Region

#Region "  ---  Properties"


        Private F_sheet_Points_Coordinates As Worksheet
        ''' <summary>
        ''' 工作表：监测点编号与对应坐标(在CAD中的坐标)
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property sheet_Points_Coordinates As Worksheet
            Get
                Return Me.F_sheet_Points_Coordinates
            End Get
        End Property

        ''' <summary>
        ''' 以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
        ''' </summary>
        ''' <remarks>返回一个字典，其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</remarks>
        Private F_ID_Components As New Dictionary(Of String, clsData_ExcavationID)(StringComparer.OrdinalIgnoreCase)
        ''' <summary>
        ''' 以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
        ''' </summary>
        ''' <value></value>
        ''' <returns>返回一个字典，其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</returns>
        ''' <remarks></remarks>
        Public Property ID_Components As Dictionary(Of String, clsData_ExcavationID)
            Get
                Return F_ID_Components
            End Get
            Set(value As Dictionary(Of String, clsData_ExcavationID))
                Me.F_ID_Components = value
                RaiseEvent dic_IDtoComponentsChanged(value)
            End Set
        End Property

        ''' <summary>
        ''' 整个工程项目中的所有的基坑区域的集合
        ''' </summary>
        ''' <remarks></remarks>
        Private P_lst_ProcessRegion As New List(Of clsData_ProcessRegionData)
        ''' <summary>
        ''' 整个工程项目中的所有的基坑区域的集合，其中包含了与此基坑区域相关的各种信息，比如区域名称，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProcessRegion As List(Of clsData_ProcessRegionData)
            Get
                Return Me.P_lst_ProcessRegion
            End Get
            Set(value As List(Of clsData_ProcessRegionData))
                Me.P_lst_ProcessRegion = value
                RaiseEvent ProcessRangeChanged(value)
            End Set
        End Property

        ''' <summary>
        ''' 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据
        ''' </summary>
        ''' <remarks></remarks>
        Private F_ShapeIDAndFinishedDate As clsData_ShapeID_FinishedDate
        ''' <summary>
        ''' 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据
        ''' </summary>
        ''' <value></value>
        ''' <returns>返回一个数组，数组中有两个元素，其中每一个元素都是以向量的形式记录了这两列数据。</returns>
        ''' <remarks>向量的第一个元素的下标值为0。</remarks>
        Public Property ShapeIDAndFinishedDate As clsData_ShapeID_FinishedDate
            Get
                Return F_ShapeIDAndFinishedDate
            End Get
            Private Set(value As clsData_ShapeID_FinishedDate)
                F_ShapeIDAndFinishedDate = value
            End Set
        End Property

        ''' <summary>
        ''' 以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
        ''' 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高
        ''' </summary>
        ''' <remarks></remarks>
        Private F_WorkingStage As Dictionary(Of String, List(Of clsData_WorkingStage))
        ''' <summary>
        ''' 以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
        ''' 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property WorkingStage As Dictionary(Of String, List(Of clsData_WorkingStage))
            Get
                Return Me.F_WorkingStage
            End Get
            Private Set(value As Dictionary(Of String, List(Of clsData_WorkingStage)))
                Me.F_WorkingStage = value
                RaiseEvent WorkingStageChanged(value)
            End Set
        End Property

#End Region

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="FileContents">根据项目文件来提取相应的工作簿和工作表对象</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal FileContents As clsData_FileContents)
            If FileContents IsNot Nothing Then

                With FileContents
                    '开挖剖面工作表及其数据提取
                    Me.F_shtElevation = .Sheet_Elevation
                    Try
                        If Me.F_shtElevation IsNot Nothing Then
                            'Me.ID_Components的赋值一定要在施工进度工作表的赋值之前，因为后面在赋值时可能会用到Me.ID_Components的值。
                            Me.ID_Components = ExtractElevation(F_shtElevation)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("基坑ID及剖面标高信息提取出错，请检查""剖面标高""工作表的数据格式是否正确。" &
                                        vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                                        ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Me.ID_Components = Nothing
                    End Try

                    '施工进度工作表及其数据提取
                    If .lstSheets_Progress IsNot Nothing Then
                        Try
                            Me.ProcessRegion = GetDataForProcessRegion(.lstSheets_Progress)
                        Catch ex As Exception
                            MessageBox.Show("基坑群的施工进度信息提取出错，请检查""施工进度""工作表的数据格式是否正确。" &
                                            vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                                            ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Me.ProcessRegion = Nothing
                        End Try
                    End If

                    '记录测点坐标的工作表
                    Me.F_sheet_Points_Coordinates = .Sheet_PointCoordinates

                    'Excel中记录Visio中开挖分块的形状的相关信息的工作表
                    Try
                        If .Sheet_PlanView IsNot Nothing Then
                            Me.F_ShapeIDAndFinishedDate = GetShapeID_FinishedDatePair(.Sheet_PlanView)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("开挖分块信息提取出错，请检查""开挖分块""工作表的数据格式是否正确。" &
                                        vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                                        ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Me.F_ShapeIDAndFinishedDate = Nothing
                    End Try

                    '开挖工况信息
                    Try
                        'Dim wkbk As Workbook = Me.F_shtElevation.Parent
                        'Dim shtWorkingStage As Worksheet = wkbk.Worksheets.Item("开挖工况")
                        If .Sheet_WorkingStage IsNot Nothing Then
                            Me.WorkingStage = GetWorkingStage(.Sheet_WorkingStage)
                        Else
                            Me.WorkingStage = Nothing
                        End If
                    Catch ex As Exception
                        MessageBox.Show("提取开挖工况信息出错！" & vbCrLf & ex.Message & vbCrLf &
                                        "报错位置：" & ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Me.WorkingStage = Nothing
                    End Try
                End With
            End If
        End Sub

        '剖面标高工作表的数据提取
        ''' <summary>
        ''' 创建字典：dic_IDtoElevationAndData
        ''' </summary>
        ''' <param name="shtElevation"></param>
        ''' <returns> 创建一个字典，以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
        ''' 其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</returns>
        ''' <remarks></remarks>
        Public Function ExtractElevation(ByVal shtElevation As Worksheet) As Dictionary(Of String, clsData_ExcavationID)
            Dim IDtoElevationAndData As New Dictionary(Of String, clsData_ExcavationID)
            '创建一个字典dic_IDtoElevationAndData，以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
            '其key为基坑的ID，value为[此基坑的标高项目,对应的标高]的数组
            With shtElevation
                Const cstRowID As Byte = Data_Drawing_Format.DB_Sectional.RowNum_ID
                Const cstColFirstID As Byte = Data_Drawing_Format.DB_Sectional.ColNum_FirstID
                '
                Dim ColumnCount As Integer = .UsedRange.Columns.Count
                For Each icell As Range In .Range(.Cells(cstRowID, cstColFirstID), .Cells(cstRowID, ColumnCount))
                    '如果单元格是位于合并单元格内
                    If icell.MergeCells Then

                        ' ----------------- 确定工作表中第一行的基坑名称所在的单元格
                        If String.Compare(icell.MergeArea.Address, icell.Offset(0, -1).MergeArea.Address) <> 0 Then
                            '说明此单元格不是位于两个单元格所形成的合并单元格中的第二个，所以它是位于第一个

                            '合并单元格中的第一个单元格，其所在的列即为此基坑ID下的结构构件数据所在的列，
                            '而其后面一列即为此基坑ID下的结构构件的标高数据所在的列。
                            Dim firstInMerge As Range = icell.MergeArea.Cells(1, 1)
                            '基坑ID
                            Dim ID As String = firstInMerge.Value
                            '基坑底部标高
                            Dim ExcavationBottom As Single
                            '底板顶部标高
                            Dim BottomFloor As Single

                            '将基坑名称与这个基坑的结构和标高相对应，并放在一个集合中
                            Dim firstCellColNum = firstInMerge.Column
                            Dim ArrItemAndData As Component() = ColNumToComponents(firstCellColNum, ExcavationBottom, BottomFloor)
                            '
                            Dim ExcavationID As clsData_ExcavationID = New clsData_ExcavationID(ID, ExcavationBottom, BottomFloor, ArrItemAndData)
                            If ExcavationID Is Nothing Then
                                Return Nothing
                            End If
                            IDtoElevationAndData.Add(key:=ID, value:=ExcavationID)
                        End If
                    End If
                Next
            End With
            Return IDtoElevationAndData
        End Function
        ''' <summary>
        ''' 由基坑ID所在的列号返回此基坑ID对应的每一个构件项目的名称与标高
        ''' </summary>
        ''' <param name="firstCellColNum">第一个单元格的列号，即为此基坑ID下的结构构件数据所在的列</param>
        ''' <param name="ExcavationBottom">作为输出变量，表示基坑底的标高</param>
        ''' <param name="BottomFloor">作为输出变量，表示底板顶的标高</param>
        ''' <returns>数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值</returns>
        ''' <remarks></remarks>
        Private Function ColNumToComponents(ByVal firstCellColNum As Integer, ByRef ExcavationBottom As Single, ByRef BottomFloor As Single) As Component()
            '提取数据
            Dim rg As Range
            With Me.F_shtElevation
                rg = .Range(.Cells(3, firstCellColNum), .Cells(DB_Sectional.RowNum_EndRowInElevation, firstCellColNum + 1))
            End With
            Dim AllData As Object(,) = rg.Value
            '
            Try
                Dim index As UShort = 0
                Dim lb As Byte = LBound(AllData, 1)
                Dim Components(0 To UBound(AllData, 1) - lb) As Component
                For i = lb To UBound(AllData, 1)
                    If AllData(i, 2) IsNot Nothing Then
                        Dim tag As String = AllData(i, 1)
                        Dim Elevation As Single = AllData(i, 2)

                        ' ----------------- 判断每一个构件的类型 -------------------------------------
                        '"Contains" method performs an ordinal (case-sensitive and culture-insensitive) comparison.
                        Dim CompontType As ComponentType = ComponentType.Others
                        If tag.IndexOf(DB_Sectional.identifier_TopOfBottomSlab, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                            BottomFloor = Elevation
                            CompontType = ComponentType.TopOfBottomSlab

                        ElseIf tag.IndexOf(DB_Sectional.identifier_ExcavationBottom, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                            ExcavationBottom = Elevation
                            CompontType = ComponentType.ExcavationBottom

                        ElseIf tag.IndexOf(DB_Sectional.identifier_Floor, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                            CompontType = ComponentType.Floor

                        ElseIf tag.IndexOf(DB_Sectional.identifier_struts, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                            CompontType = ComponentType.Strut

                        ElseIf tag.IndexOf(DB_Sectional.identifier_Ground, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                            CompontType = ComponentType.Ground
                        End If
                        ' -------------------------------------------------
                        Components(index) = New Component(CompontType, tag, Elevation)
                        index = index + 1
                    End If
                Next
                '剔除数据库中为空的构件
                Dim arrComponentsWithoutEmpty(0 To index - 1) As Component
                For i2 = 0 To index - 1
                    arrComponentsWithoutEmpty(i2) = Components(i2)
                Next
                Return arrComponentsWithoutEmpty
            Catch ex As Exception
                MessageBox.Show("提取数据库中的基坑ID及其构件的信息出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function

        '开挖分块工作表的数据提取
        ''' <summary>
        ''' 获取工作表“开挖分块”中的“形状名”与“完成日期”的组合。 
        ''' </summary>
        ''' <param name="shtRegion">开挖分块工作表</param>
        ''' <returns></returns>
        ''' <remarks>数组的第一个元素的下标值为0</remarks>
        Private Function GetShapeID_FinishedDatePair(ByVal shtRegion As Worksheet) As clsData_ShapeID_FinishedDate

            Const cstColNum_ShapeID As Byte = DB_ExcavRegionForVisio.ColNum_ShapeID '工作表“开挖分块”中的“形状名”所在的列号
            Const cstColNum_FinishedDate As Byte = DB_ExcavRegionForVisio.ColNum_FinishedDate '工作表“开挖分块”中的“完成日期”所在的列号
            Const cstRowNum_FirstShape As Byte = DB_ExcavRegionForVisio.RowNum_FirstShape '第一个形状数据所在的行号

            '从工作表的range中提取数据
            Dim rgFinishedDate As Range
            Dim rgShapeID As Range
            With shtRegion
                Dim endrow As Integer = .UsedRange.Rows.Count
                rgFinishedDate = .Columns(cstColNum_FinishedDate).range(.Cells(cstRowNum_FirstShape, 1), .Cells(endrow, 1))
                rgShapeID = .Columns(cstColNum_ShapeID).range(.Cells(cstRowNum_FirstShape, 1), .Cells(endrow, 1))
            End With
            '将Object类型的二维数组转换为Integer类型与Date类型的一维数组，而且让数组的第一个元素的下标值为0
            Dim intShapeID() As Integer = ExcelFunction.ConvertRangeDataToVector(Of Integer)(rgShapeID)
            Dim dateFinishedDate() As Date = ExcelFunction.ConvertRangeDataToVector(Of Date)(rgFinishedDate)
            Return New clsData_ShapeID_FinishedDate(intShapeID, dateFinishedDate)
        End Function

        '开挖工况工作表的数据提取
        ''' <summary>
        ''' 获取工作表“开挖工况”中的基坑区域名称，以及每一个基坑区域中的开挖工况信息
        ''' </summary>
        ''' <param name="shtWorkingStage"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetWorkingStage(ByVal shtWorkingStage As Worksheet) As Dictionary(Of String, List(Of clsData_WorkingStage))
            '
            Dim allData As Object(,) = shtWorkingStage.UsedRange.Value
            Dim lb As Byte = LBound(allData, 2)         '对于Excel中返回的数组，第一个元素的下标值为1，而不是0
            Dim ColumnsCount As UShort = UBound(allData, 2) - lb + 1
            Dim RowsCount As UShort = UBound(allData, 1) - lb + 1
            '如果数据列的列数是3的整数，说明工作表的列的数据排布是规范的。
            If (UBound(allData, 2) - lb + 1) Mod DB_WorkingStage.ColCount_EachRegion <> 0 Then
                MessageBox.Show("开挖工况信息工作表中的数据列的排版不规范，请检查。", _
                                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If
            '基坑区域的个数
            Dim RegionsCount As UShort = ColumnsCount / DB_WorkingStage.ColCount_EachRegion
            '
            Dim dicWS As New Dictionary(Of String, List(Of clsData_WorkingStage))
            '
            For index_Region As UShort = 0 To RegionsCount - 1      '第几个基坑区域
                Dim strRegionName As String = allData(DB_WorkingStage.RowNum_RegionName - 1 + lb, index_Region * DB_WorkingStage.ColCount_EachRegion + lb)
                Dim listWS As New List(Of clsData_WorkingStage)
                '
                For RN_WS As UShort = DB_WorkingStage.RowNum_FirstStage - 1 + lb To RowsCount - 1 + lb  '一个基坑区域中每一个工况在数组中的行号
                    Dim CN_FirstReference As UShort = index_Region * DB_WorkingStage.ColCount_EachRegion + lb
                    Try
                        '提取数组中的每一行开挖工况，并将其转换为指定的类型。
                        '如果数据提取或者转换出错，说明此基坑区域的施工工况信息已经提取完毕。
                        Dim strDescription As String = allData(RN_WS, CN_FirstReference + DB_WorkingStage.Index_Description - lb)
                        Dim Elevation As Single = allData(RN_WS, CN_FirstReference + DB_WorkingStage.Index_Elevation - lb)
                        Dim ConstructionDate As Date = allData(RN_WS, CN_FirstReference + DB_WorkingStage.Index_ConstructionDate - lb)
                        If ConstructionDate.Ticks = 0 Then
                            Exit For
                        End If
                        '
                        listWS.Add(New clsData_WorkingStage(strDescription, ConstructionDate, Elevation))
                    Catch ex As Exception
                        Exit For
                    End Try
                Next
                dicWS.Add(strRegionName, listWS)
            Next
            Return dicWS
        End Function

#Region "  ---  施工进度工作表"

        '施工进度工作表的操作
        ''' <summary>
        ''' 施工进度工作表的操作：得到所有基坑区域的施工进度信息
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetDataForProcessRegion(ByVal lst_Sheets_Progress As List(Of Worksheet)) As List(Of clsData_ProcessRegionData)
            '--------------------------------------------
            Const cstColNum_FirstRegionInProgress As Byte = DB_Progress.ColNum_theFirstRegion
            Const cstColNum_DateList As Byte = DB_Progress.ColNum_DateList
            Const cstRowNum_FirstDate As Byte = DB_Progress.RowNum_TheFirstDay
            Const cstRowNum_ExcavTag As Byte = DB_Progress.RowNum_ExcavTag
            Const cstRowNum_ExcavPosition As Byte = DB_Progress.RowNum_ExcavPosition
            '--------------------------------------------
            Dim lstRegion As New List(Of clsData_ProcessRegionData)
            'Dim dicExcavTagToColRange As New Dictionary(Of String, Range)(StringComparer.OrdinalIgnoreCase)
            For Each sht As Worksheet In lst_Sheets_Progress
                With sht
                    Dim UsedRg As Excel.Range = .UsedRange
                    Dim btRegionCount As UShort = UsedRg.Columns.Count
                    Dim btRowsCount As UShort = UsedRg.Rows.Count
                    '以日期所在的数据列作为数据库的主键的列
                    Dim Rg_DateColumn As Excel.Range = UsedRg.Columns(cstColNum_DateList)
                    '日期列中真正保存日期数据的区域
                    Dim rg_Date As Excel.Range = Rg_DateColumn.Range(.Cells(cstRowNum_FirstDate, 1), .Cells(btRowsCount, 1))
                    Dim arrDate As Object(,) = rg_Date.Value
                    '--------------------------------------------
                    For iCol As Byte = cstColNum_FirstRegionInProgress To btRegionCount
                        Try
                            Dim Rg_Process As Range = UsedRg.Columns(iCol)
                            '基坑名：A1/B/C1等
                            Dim mergecell As Range = Rg_Process.Cells(cstRowNum_ExcavTag, 1).MergeArea(1, 1)
                            Dim strExcavName As String = mergecell.Value
                            '基坑分块：普遍区域、东南侧、西侧等
                            Dim strExcavPosition As String = Rg_Process.Cells(cstRowNum_ExcavPosition, 1).Value

                            '构造此基坑区域的开挖标高的索引
                            Dim strExcavID As String = Rg_Process.Cells(Data_Drawing_Format.DB_Progress.RowNum_ExcavID, 1).value
                            Dim ExcavID As clsData_ExcavationID = Nothing
                            If Me.ID_Components.Keys.Contains(strExcavID) Then
                                ExcavID = Me.ID_Components.Item(strExcavID)
                            End If

                            '构造此基坑区域在每一天的施工标高
                            Dim arrElevation As Object(,) = rg_Date.Offset(0, iCol - cstColNum_DateList).Value
                            Dim Date_Process As SortedList(Of Date, Single) = GetDate_Elevatin(sht, arrDate, arrElevation)

                            '
                            Dim strBottomDate As String = Rg_Process.Cells(Data_Drawing_Format.DB_Progress.RowNum_BottomDate, 1).value

                            '--------------------------------------------

                            Dim RD As New clsData_ProcessRegionData(strExcavName, _
                                                               strExcavPosition, _
                                                               ExcavID, strBottomDate, _
                                                               Rg_Process, Rg_DateColumn, Date_Process)

                            '--------------------------------------------
                            lstRegion.Add(RD)
                        Catch ex As Exception

                            Return New List(Of clsData_ProcessRegionData)
                        End Try
                    Next

                End With
            Next
            Return lstRegion
        End Function

        ''' <summary>
        ''' 创建每一天的日期到当天的开挖标高的索引
        ''' </summary>
        ''' <param name="arrDate">在施工进度工作表中，保存日期的那一列数组，数组中只有日期的数据，而没有前几行的表头信息数据。
        ''' 从标高数据的搜索与处理方式上，要求此列日期数据必须是从早到晚进行排列的！</param>
        ''' <param name="arrElevation">在施工进度工作表中，保存日期的那一列数组，数组中只有开挖标高的数据，而没有前几行的表头信息数据</param>
        ''' <returns>此基坑区域的每一天所对应的开挖标高</returns>
        ''' <remarks>如果在搜索的某一天天没有找到对应的开挖标高的数据，那么就开始向更早的日期进行搜索，
        ''' 在数据库的设计中认为，如果今天没有开挖标高的数据，说明今天的开挖标高位于早于今天的最近的那一天的有效标高与晚于今天的最近
        ''' 的那一天的标高之间，在处理时，即认为今天的开挖标高还是位于早于今天的最近的那一天的有效标高。</remarks>
        Private Function GetDate_Elevatin(ByVal sheetData As Excel.Worksheet,
                                          ByVal arrDate As Object(,), ByVal arrElevation As Object(,)) As SortedList(Of Date, Single)
            Dim Sorted_Date_Elevation As New SortedList(Of Date, Single)
            Dim Ground As Single = Project_Expo.Elevation_GroundSurface
            '
            Dim lb As Byte = LBound(arrDate)
            Dim Dt As Date
            Dim Elevation As Object
            Dim ReferenceElevation As Single = Ground
            '
            Dim index_Sortedlist As UShort = 0
            Dim i As Integer = lb
            Try
                For i = lb To UBound(arrDate, 1)
                    Dt = CType(arrDate(i, lb), Date)
                    Elevation = arrElevation(i, lb)
                    '------- 确定具体的开挖标高 "Elevation" 的值 ---------------
                    If Elevation Is Nothing OrElse VarType(Elevation) = VariantType.Empty Then      '说明在今天没有找到对应的开挖标高的数据，
                        '那么就要开始向更早的日期进行搜索，在数据库的设计中认为，如果今天没有开挖标高的数据，
                        '说明今天的开挖标高位于早于今天的最近的那一天的有效标高与晚于今天的最近的那一天的标高之间，
                        '在处理时，即认为今天的开挖标高还是位于早于今天的最近的那一天的有效标高。
                        Elevation = ReferenceElevation
                    End If
                    ReferenceElevation = Elevation
                    '--------------------------------------------
                    Sorted_Date_Elevation.Add(Dt, Elevation)
                Next
            Catch ex As Exception
                MessageBox.Show("施工进度工作表"" " & sheetData.Name & " ""中，第 " & (i - lb + DB_Progress.RowNum_TheFirstDay).ToString &
                                " 行的数据出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return Sorted_Date_Elevation
        End Function

        '此方法可删！！！由监测点所在基坑区域，以及当天的日期，来得到此点当日的开挖标高
        ''' <summary>
        ''' 此方法可删！！！ 由监测点所在基坑区域，以及当天的日期，来得到此点当日的开挖标高
        ''' </summary>
        ''' <param name="ExcavationRegion">在施工进度工作表中，考察点所在的基坑区域Region，其中每一个Range对象，
        ''' 都代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
        ''' <param name="dateThisday">考察的日期</param>
        ''' <returns>此点当天的开挖标高</returns>
        ''' <remarks></remarks>
        Public Shared Function GetElevation(ByVal ExcavationRegion As Range, dateThisday As Date) As String
            Dim Elevation As String
            Dim rowByDate As Integer
            '每一张进度表中，日期的第一天的数据所在的行
            Const cstRowNum_TheFirstDay As Byte = DB_Progress.RowNum_TheFirstDay
            '每一个区域（可能分布于不同的基坑开挖进度的工作表）都有其特有的初始日期
            Dim dateTheFirstDay As Date = ExcavationRegion.Worksheet.Cells(cstRowNum_TheFirstDay, DB_Progress.ColNum_DateList).value
            rowByDate = dateThisday.Subtract(dateTheFirstDay).Days + cstRowNum_TheFirstDay
            If rowByDate >= cstRowNum_TheFirstDay Then   '正常操作
                Elevation = ExcavationRegion.Cells(rowByDate, 1).Value
                Dim i1 = 0
                '任意一天的基坑开挖深度的确定机制： 如果这一天没有挖深的数据，则向上进行索引，一直索引到此基坑区域开挖的第一天。
                Do Until Elevation <> "" Or rowByDate - i1 < cstRowNum_TheFirstDay
                    Elevation = ExcavationRegion.Cells(rowByDate - i1, 1).Value
                    i1 = i1 + 1
                Loop
            Else                '说明这一天的日期小于此基坑区域中记录的最早日期，说明还处于未开挖状态，此时应该将其Depth(i)的值设定为自然地面的标高。
                Elevation = Project_Expo.Elevation_GroundSurface
            End If
            Return Elevation
        End Function

        ''' <summary>
        ''' 从SortedList的日期集合中，搜索距离指定的日期最近的那一天。
        ''' </summary>
        ''' <param name="Keys"></param>
        ''' <param name="WantedDate"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindTheClosestDateInSortedList(ByVal Keys As IList(Of Date), ByVal WantedDate As Date) As Date
            With Keys
                Dim FirstDay As Date = .First
                Dim LastDay As Date = .Last
                ' ------------------------------
                '先考察搜索的日期是否是超出了集合中的界限()
                If WantedDate.CompareTo(FirstDay) <= 0 Then
                    Return FirstDay
                ElseIf WantedDate.CompareTo(LastDay) >= 0 Then
                    Return LastDay
                End If
                ' ------------------------------

                Dim ClosestDate As Date = WantedDate
                Dim SearchingDate As Date = WantedDate
                Dim Date_Earlier As Date
                Dim Date_Later As Date
                Dim blnItemFound As Boolean = False
                Dim Index As Integer
                ' ----------
                '先向上搜索离之最近的那一天DA
                Do Until blnItemFound
                    Index = Keys.IndexOf(SearchingDate)
                    If Index >= 0 Then
                        Date_Earlier = SearchingDate
                        blnItemFound = True
                    Else
                        '将日期提前一天进行搜索
                        SearchingDate = SearchingDate.Subtract(TimeSpan.FromDays(1))
                        blnItemFound = False
                    End If
                Loop
                Date_Later = Keys.Item(Index + 1)
                If Date_Later.Subtract(WantedDate) < WantedDate.Subtract(Date_Earlier) Then
                    ClosestDate = Date_Later
                Else
                    ClosestDate = Date_Earlier
                End If
                Return ClosestDate
            End With
        End Function

#End Region

    End Class

End Namespace
