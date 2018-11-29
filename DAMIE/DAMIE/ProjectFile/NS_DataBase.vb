Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Namespace DataBase

#Region "  ---  Enumerators"

    ''' <summary>
    ''' 基坑中各种构件的类型
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ComponentType As Byte
        ''' <summary>
        ''' 自然地面
        ''' </summary>
        ''' <remarks></remarks>
        Ground
        ''' <summary>
        ''' 支撑
        ''' </summary>
        ''' <remarks></remarks>
        Strut
        ''' <summary>
        ''' 楼板
        ''' </summary>
        ''' <remarks></remarks>
        Floor
        ''' <summary>
        ''' 基坑底板顶部
        ''' </summary>
        ''' <remarks></remarks>
        TopOfBottomSlab
        ''' <summary>
        ''' 基坑底部
        ''' </summary>
        ''' <remarks></remarks>
        ExcavationBottom
        ''' <summary>
        ''' 其他
        ''' </summary>
        ''' <remarks></remarks>
        Others
    End Enum

#End Region

    ''' <summary>
    ''' 在“剖面标高”工作表中，每一个基坑ID所对应的相关信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsData_ExcavationID

        Private _ExcavationBottom As Single
        ''' <summary>
        ''' 基坑底部标高
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExcavationBottom As Single
            Get
                Return _ExcavationBottom
            End Get
        End Property

        Private _BottomFloor As Single
        ''' <summary>
        ''' 底板顶部标高
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property BottomFloor As Single
            Get
                Return _BottomFloor
            End Get
        End Property

        Private _Components As Component()
        ''' <summary>
        ''' 记录基坑ID中对应的每一个构件项目与其对应标高的数组,
        ''' 数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Components As Component()
            Get
                Return _Components
            End Get
        End Property

        Private _name As String
        ''' <summary>
        ''' 此基坑ID的ID名，如“A1-1”
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name As String
            Get
                Return Me._name
            End Get
        End Property

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="name">此基坑ID的ID名，如“A1-1”</param>
        ''' <param name="pitBottom">基坑底部标高</param>
        ''' <param name="bottomfloor">底板顶部标高</param>
        ''' <param name="cmpnts">记录基坑ID下的结构构件及其标高的数组</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal pitBottom As Single, ByVal bottomfloor As Single, ByVal cmpnts As Component())
            Me._ExcavationBottom = pitBottom
            Me._BottomFloor = bottomfloor
            Me._Components = cmpnts
            Me._name = name
        End Sub

    End Class

    ''' <summary>
    ''' 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsData_ProcessRegionData

        ''' <summary>
        ''' 基坑区域的施工进度中，每一个子区域的开挖标高数据对应的Range对象，
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
        ''' </summary>
        ''' <remarks></remarks>
        Private _Range_Process As Excel.Range
        ''' <summary>
        ''' 基坑区域的施工进度中，每一个子区域的开挖标高数据对应的Range对象，
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Range_Process As Excel.Range
            Get
                Return _Range_Process
            End Get
        End Property

        ''' <summary>
        ''' 基坑区域的施工进度中，每一个子区域的开挖时间对应的Range对象，
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
        ''' </summary>
        ''' <remarks></remarks>
        Private _Range_Date As Excel.Range
        ''' <summary>
        ''' 基坑区域的施工进度中，每一个子区域的开挖时间对应的Range对象，
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Range_Date As Excel.Range
            Get
                Return Me._Range_Date
            End Get
        End Property

        ''' <summary>
        ''' 在基坑区域的施工进度中，每一天的日期所对应的开挖标高。
        ''' </summary>
        ''' <remarks></remarks>
        Private _Date_Elevation As SortedList(Of Date, Single)
        ''' <summary>
        ''' 在基坑区域的施工进度中，每一天的日期所对应的开挖标高。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Date_Elevation As SortedList(Of Date, Single)
            Get
                Return Me._Date_Elevation
            End Get
        End Property

        Private _blnHasBottomDate As Boolean
        ''' <summary>
        ''' 指示在数据库文件中，指定的开挖区域是否已经开挖到了基坑底
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property HasBottomDate As Boolean
            Get
                Return _blnHasBottomDate
            End Get
        End Property

        Private _BottomDate As Date
        ''' <summary>
        ''' 在数据库文件中，指定的开挖区域开挖到基坑底的日期
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property BottomDate As Date
            Get
                Return _BottomDate
            End Get
        End Property

        Private _ExcavationID As clsData_ExcavationID
        ''' <summary>
        ''' 基坑区域对应的基坑ID的信息
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExcavationID As clsData_ExcavationID
            Get
                Return _ExcavationID
            End Get
        End Property

        Private _ExcavName As String
        ''' <summary>
        ''' 基坑区域所在的基坑的名称，如A1、B、C1等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExcavName As String
            Get
                Return Me._ExcavName
            End Get
        End Property

        ''' <summary>
        ''' 基坑区域的分块名称，如普遍区域、东南侧、西侧等
        ''' </summary>
        ''' <remarks></remarks>
        Private _ExcavPosition As String
        ''' <summary>
        ''' 基坑区域的分块名称，如普遍区域、东南侧、西侧等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExcavPosition As String
            Get
                Return Me._ExcavPosition
            End Get
        End Property

        Private _description As String
        ''' <summary>
        ''' 关于此基坑区域的描述，如“A1:普遍区域”
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property description As String
            Get
                Return Me._description
            End Get
            Set(value As String)
                Me._description = value
            End Set
        End Property

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="ExcavName">基坑区域所在的基坑的名称，如A1、B、C1等</param>
        ''' <param name="ExcavPosition">基坑区域的分块名称，如普遍区域、东南侧、西侧等</param>
        ''' <param name="ExcavationID">基坑区域对应的基坑ID的信息</param>
        ''' <param name="strBottomDate">在数据库文件中，指定的开挖区域开挖到基坑底的日期，在构造函数中进行数据类型的转换</param>
        ''' <param name="Range_Process">基坑区域的施工进度中，每一个子区域的开挖标高对应的Range对象,
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>  
        ''' <param name="Range_Date">基坑区域的施工进度中，每一个子区域的开挖日期对应的Range对象,
        ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
        ''' <param name="Date_Elevation">在基坑区域的施工进度中，每一天的日期所对应的开挖标高。</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ExcavName As String, ByVal ExcavPosition As String, _
                       ByVal ExcavationID As clsData_ExcavationID, ByVal strBottomDate As String, _
                       ByVal Range_Process As Excel.Range, ByVal Range_Date As Excel.Range, _
                       ByVal Date_Elevation As SortedList(Of Date, Single))
            ' ----------------------------------------
            With Me
                ._ExcavationID = ExcavationID
                ._ExcavName = ExcavName
                ._ExcavPosition = ExcavPosition
                ._description = ExcavName & ":" & ExcavPosition
                ._Range_Date = Range_Date
                ._Range_Process = Range_Process
                ._Date_Elevation = Date_Elevation
            End With

            '将数据库文件中，指定的开挖区域开挖到基坑底的日期的单元格的数据，转换为日期类型
            '如果转换错误，即此单元格为空，或者单元格数据格式不能转换为日期格式，那说明此基坑区域还没有开挖到基坑底部标高
            Try         '对“2014/11/26”形式的日期进行转换
                Me._BottomDate = Date.Parse(strBottomDate)
                Me._blnHasBottomDate = True
            Catch ex1 As Exception
                Try     '对“”形式的日期进行转换，比如“41969”对应于日期的“2014/11/26”
                    Me._BottomDate = Date.FromOADate(strBottomDate)
                Catch ex2 As Exception
                    '如果上面两种转换方法都不能将单元格中的数据转换为日期类型，则认为此基坑区域还没有开挖到基坑底部标高
                    Me._blnHasBottomDate = False
                End Try
            End Try

        End Sub

    End Class

    ''' <summary>
    ''' 记录Visio绘图中所有开挖分块的形状的ID值及每一个分块的完成日期的两个数组，
    ''' 这两个数组的元素个数必须相同。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsData_ShapeID_FinishedDate
        ''' <summary>
        ''' 列举Visio绘图中所有开挖分块的形状的ID值的数组
        ''' </summary>
        ''' <remarks></remarks>
        Public ShapeID As Integer()
        ''' <summary>
        ''' 每一个开挖分块所对应的开挖完成的日期
        ''' </summary>
        ''' <remarks></remarks>
        Public FinishedDate As Date()

        ''' <summary>
        ''' 构造函数，ShapeID与FinishedDate这两个数组中的元素个数必须相同，而且其中第一个元素的下标值为0。
        ''' </summary>
        ''' <param name="ShapeID">列举Visio绘图中所有开挖分块的形状的ID值的数组</param>
        ''' <param name="FinishedDate">每一个开挖分块所对应的开挖完成的日期的数组</param>
        ''' <remarks></remarks>
        Public Sub New(ShapeID As Integer(), FinishedDate As Date())
            If ShapeID.Length <> FinishedDate.Length Then
                MessageBox.Show("形状ID的数组与对应的完成日期的数组的元素个数必须相同！", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Me.ShapeID = ShapeID
                Me.FinishedDate = FinishedDate
            End If
        End Sub
    End Class

    ''' <summary>
    ''' 某一基坑的工况信息，它主要包括此基坑在某一特定的日期开挖到了什么状态，当时的开挖深度值
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsData_WorkingStage

        ''' <summary>
        ''' 工况描述
        ''' </summary>
        ''' <remarks></remarks>
        Public Description As String

        ''' <summary>
        ''' 施工日期
        ''' </summary>
        ''' <remarks></remarks>
        Public ConstructionDate As Date

        ''' <summary>
        ''' 开挖标高，单位为m，书写格式为：4.200，不带单位。
        ''' 注意：是开挖标高值，不是开挖深度值。
        ''' </summary>
        ''' <remarks></remarks>
        Public Elevation As Single

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="Description">工况描述</param>
        ''' <param name="ConstructionDate">施工日期</param>
        ''' <param name="Elevation">开挖标高，单位为m，书写格式为：4.200，不带单位。
        ''' 注意：是开挖标高值，不是开挖深度值。</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Description As String, ByVal ConstructionDate As Date, ByVal Elevation As Single)
            With Me
                .Description = Description
                .ConstructionDate = ConstructionDate
                .Elevation = Elevation
            End With
        End Sub

    End Class

    ''' <summary>
    ''' 在基坑设计中，每一个不同的基坑ID区域，所包含的构件信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure Component
        ''' <summary>
        ''' 此构件的类型
        ''' </summary>
        ''' <remarks></remarks>
        Public Type As ComponentType
        ''' <summary>
        ''' 对于此构件的描述
        ''' </summary>
        ''' <remarks></remarks>
        Public Description As String
        ''' <summary>
        ''' 此构件的特征标高
        ''' </summary>
        ''' <remarks></remarks>
        Public Elevation As Single

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="Type">此构件的类型</param>
        ''' <param name="Description">对于此构件的描述</param>
        ''' <param name="Elevation">此构件的特征标高</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Type As ComponentType, ByVal Description As String, ByVal Elevation As Single)
            With Me
                .Type = Type
                .Description = Description
                .Elevation = Elevation
            End With
        End Sub

    End Structure

    ''' <summary>
    ''' 文件中要保存的内容或对象
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsData_FileContents

        ''' <summary>
        ''' 项目文件中记录的所有工作簿
        ''' </summary>
        ''' <remarks></remarks>
        Public Property lstWkbks As List(Of Workbook)

        ''' <summary>
        ''' 施工进度工作表
        ''' </summary>
        ''' <remarks></remarks>
        Public Property lstSheets_Progress As List(Of Worksheet)
        ''' <summary>
        ''' 开挖平面图工作表
        ''' </summary>
        ''' <remarks></remarks>
        Public Property Sheet_PlanView As Worksheet
        ''' <summary>
        ''' 开挖剖面图工作表
        ''' </summary>
        ''' <remarks></remarks>
        Public Property Sheet_Elevation As Worksheet
        ''' <summary>
        ''' 测点坐标工作表
        ''' </summary>
        ''' <remarks></remarks>
        Public Property Sheet_PointCoordinates As Worksheet
        ''' <summary>
        ''' 开挖工况信息工作表
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Sheet_WorkingStage As Worksheet

        Public Sub New()
            Me.lstSheets_Progress = New List(Of Worksheet)
            Me.lstWkbks = New List(Of Workbook)
        End Sub
    End Class

End Namespace