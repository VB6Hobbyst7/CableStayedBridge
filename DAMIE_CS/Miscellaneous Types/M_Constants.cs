// VBConversions Note: VB project level imports
using System.Data;
using System.Diagnostics;
using System.Xml.Linq;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Collections;
using System;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
// End of VB project level imports


namespace CableStayedBridge
{
	namespace Constants
	{
		
		/// <summary>
		/// 与程序相关的一些参数，与具体的项目没有任何关系
		/// </summary>
		/// <remarks></remarks>
		public class AMEApplication
		{
			
			/// <summary>
			/// 项目文件的后缀名，后缀字符中包含了"."
			/// </summary>
			/// <remarks>用IO.Path.GetExtension返回的文件后缀，后缀字符中包含了"."。</remarks>
			public const string FileExtension = ".dm";
			
			/// <summary>
			/// 在System.Diagnostics.Process.GetProcessesByName()方法中，
			/// 用来获取当前系统中已经打开的所有Visio进程
			/// </summary>
			/// <remarks>用System.Diagnostics.Process.GetProcessesByName("VISIO")，
			/// 来获取当前系统中已经打开的所有Visio进程</remarks>
			public const string ProcessName_Visio = "VISIO";
			
			/// <summary>
			/// 字体名称：Times New Roman
			/// </summary>
			/// <remarks></remarks>
			public const string FontName_TNR = "Times New Roman";
			
			/// <summary>
			/// 长度单位厘米到磅的转换系数，即1cm对应n个磅值
			/// </summary>
			/// <remarks>后面的除1.3表示将图形缩小1.3位</remarks>
			public const float cm2pt = 72 / 2.54 / 1.3;
			
			/// <summary>
			/// 将日期转换为字符时的转换格式：2015/3/28
			/// </summary>
			/// <remarks></remarks>
			public const string DateFormat = "yyyy/M/d";
			
		} //AMEApplication
		
		/// <summary>
		/// 针对于整个项目的一些全局性的常数值
		/// </summary>
		/// <remarks></remarks>
		public class Project_Expo
		{
			
			/// <summary>
			/// 工程项目的自然地面的标高值，单位为m
			/// </summary>
			/// <remarks></remarks>
			public const float Elevation_GroundSurface = 4.2;
			
			/// <summary>
			/// 在“剖面标高”的工作表中，所有标高项目中的最高位置的标高值
			/// </summary>
			/// <remarks></remarks>
			public const float eleTop = 5;
			
			/// <summary>
			/// 测斜管的顶部的绝对标高（只是一般情况下的标高值，每根测斜管的顶部标高值可以不同）
			/// </summary>
			/// <remarks>在水平测斜的监测数据中，深度值是相对于测斜管顶部的深度；
			/// 而项目中的其他构件或者开挖的位置，是按绝对标高给出来的，所以需要此属性的值来进行二者之间的转换。
			/// </remarks>
			public const float InclineTopElevaion = 1.3;
			
		} //Project
		
		/// <summary>
		/// 文件夹或文件的名称(不是指它们的路径)
		/// </summary>
		/// <remarks></remarks>
		public class FolderOrFileName
		{
			
			/// <summary>
			/// 文件夹名称
			/// </summary>
			/// <remarks></remarks>
			public class Folder
			{
				public const string Template = "Templates";
				public const string DataBase = "DataBase";
				public const string Output = "Output";
			}
			
			/// <summary>
			/// “模板”文件存放的文件夹名称
			/// </summary>
			/// <remarks></remarks>
			public class File_Template
			{
				/// <summary>
				/// 测斜动态曲线图
				/// </summary>
				/// <remarks></remarks>
				public const string Chart_Incline = "Mnt_Incline_Vertical.crtx";
				/// <summary>
				/// 水平静态图
				/// </summary>
				/// <remarks></remarks>
				public const string Chart_Horizontal_Static = "Mnt_Static_Horizontal.crtx";
				/// <summary>
				/// 水平动态图
				/// </summary>
				/// <remarks></remarks>
				public const string Chart_Horizontal_Dynamic = "Mnt_Dyn_Horizontal.crtx";
				/// <summary>
				/// 输出到Word
				/// </summary>
				/// <remarks></remarks>
				public const string Word_Output = "Output.dotm";
				/// <summary>
				/// 对于测斜数据，将其某一个测点，在整个施工跨度内，绘制其每一天的测斜最大值以及对应的深度。此曲线图中可能要用到双Y轴
				/// </summary>
				/// <remarks></remarks>
				public const string Chart_Max_Depth = "Mnt_Static_Horizontal_DoubleY.crtx";
				/// <summary>
				/// 对于基坑区域中的开挖工况，描述在不同的施工日期的开挖标高值及其动态变化
				/// </summary>
				/// <remarks></remarks>
				public const string Chart_Elevation = "Elevation.crtx";
			}
			
			/// <summary>
			/// 数据库源文件存放的文件夹名称
			/// </summary>
			/// <remarks></remarks>
			public class SourceFile
			{
				public const string DataBase = "DataBase.xlsb";
			}
			
		} //FolderOrFileName
		
		/// <summary>
		/// Word中的段落样式
		/// </summary>
		/// <remarks></remarks>
		public class ParagraphStyle
		{
			public const string Title = "标题";
			public const string Title_1 = "标题 1";
			public const string Title_2 = "标题 2";
			public const string picture = "图片";
			public const string Content = "正文";
		} //ParagraphStyle
		
		/// <summary>
		/// 各种图形所属的图形种类的名称
		/// </summary>
		/// <remarks></remarks>
		public class DrawingItem
		{
			public const string SectionalView = "开挖剖面图";
			public const string PlanView = "开挖平面图";
			
			public const string Monitor = "监测曲线图";
			public const string Mnt_Incline = "测斜";
			public const string Mnt_Static = "静态曲线图";
			public const string Mnt_Dynamic = "动态曲线图";
			public const string Mnt_Others = "其他监测曲线图";
			
		} //DrawingItem
		
		/// <summary>
		/// 坐标轴的标题值
		/// </summary>
		/// <remarks></remarks>
		public class AxisLabels
		{
			public const string Elevation = "标高";
			public const string Excavation = "基坑";
			public const string ConstructionDate = "日期";
			public const string Displacement_mm = "位移(mm)";
			public const string Points = "测点";
			public const string Depth = "深度(m)";
			public const string AxialForce = "轴力(KN)";
			public const string Displacement_m = "位移（m）";
		} //AxisTitle
		
		/// <summary>
		/// 各种类型的Excel数据库工作表中，数据的排布格式，以及绘图的格式
		/// </summary>
		/// <remarks>这些常数只对于数据库的工作表的排布格式，以及绘图界面的UI显示格式进行定义，
		/// 并不定义与具体的项目相关的任何的信息</remarks>
		public class Data_Drawing_Format
		{
			
			/// <summary>
			/// 开挖剖面图的绘图格式
			/// </summary>
			/// <remarks></remarks>
			public class Drawing_SectionalView
			{
				
			}
			
			/// <summary>
			/// 进行动态滚动的监测数据曲线图的绘图格式
			/// </summary>
			/// <remarks></remarks>
			public class Drawing_Mnt_RollingBase
			{
				
				/// <summary>
				/// 图例形状的宽度，以磅为单位
				/// </summary>
				/// <remarks></remarks>
				public const float Legend_Width = 200;
				
				/// <summary>
				/// 图例形状的高度，以磅为单位
				/// </summary>
				/// <remarks></remarks>
				public const float Legend_Height = 50;
				
			}
			
			/// <summary>
			/// 其他监测数据图的图表界面的尺寸
			/// </summary>
			/// <remarks></remarks>
			public class Drawing_Mnt_Others
			{
				/// <summary>
				/// 图表Chart的高度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartHeight = 250;
				/// <summary>
				/// 图表Chart的宽度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartWidth = 500;
				
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离。
				/// 如果Excel窗口并没有被固定大小，则将其设置为15。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Width = 9;
				
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离。
				/// 如果Excel窗口并没有被固定大小，则将其设置为30。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Height = 26;
				
			}
			
			/// <summary>
			/// 测斜曲线图的绘图格式
			/// </summary>
			/// <remarks></remarks>
			public class Drawing_Incline
			{
				
				/// <summary>
				/// 图表Chart的高度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartHeight = 400; // 400 ' 500
				
				/// <summary>
				/// 图表Chart的宽度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartWidth = 280; // 280 ' 350
				
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离。
				/// 如果Excel窗口并没有被固定大小，则将其设置为15。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Width = 9;
				
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离
				/// 如果Excel窗口并没有被固定大小，则将其设置为30。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Height = 26;
				
				/// <summary>
				/// 图表X轴（位移）划分的区段数
				/// </summary>
				/// <remarks></remarks>
				public const byte AxisParts_X_Displacement = 10;
				
				/// <summary>
				/// 图表Y轴（测斜深度）的最大刻度值
				/// </summary>
				/// <remarks></remarks>
				public const byte AxisMajorUnit_Y = 5;
				
				/// <summary>
				/// 图例形状的宽度，以磅为单位
				/// </summary>
				/// <remarks></remarks>
				public const float Legend_Width = 150;
				
				/// <summary>
				/// 图例形状的高度，以磅为单位
				/// </summary>
				/// <remarks></remarks>
				public const float Legend_Height = 100;
			}
			
			public class Drawing_Incline_DMMD
			{
				/// <summary>
				/// 图表Chart的高度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartHeight = 250;
				/// <summary>
				/// 图表Chart的宽度
				/// </summary>
				/// <remarks></remarks>
				public const int ChartWidth = 500;
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离。
				/// 如果Excel窗口并没有被固定大小，则将其设置为15。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Width = 9;
				/// <summary>
				/// Chart外边缘到Excel界面外边缘的距离。
				/// 如果Excel窗口并没有被固定大小，则将其设置为30。
				/// </summary>
				/// <remarks></remarks>
				public const int MarginOut_Height = 26;
				
				//图表中的四条曲线的系列名称
				public const string SeriesName_Max = "最大位移";
				public const string SeriesName_Min = "最小位移";
				public const string SeriesName_Depth_Max = "最大值深度";
				public const string SeriesName_Depth_Min = "最小值深度";
				
			}
			
			/// <summary>
			/// 监测数据中的测斜数据
			/// </summary>
			/// <remarks></remarks>
			public class Mnt_Incline
			{
				/// <summary>
				/// 记录施工日期的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNumForDate = 1;
				
				/// <summary>
				/// 定义监测数据中的数据区域（包含x轴的深度数据）的起始单元格的位置：一般为“A3”
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_FirstData_WithoutDate = 2;
				/// <summary>
				/// 第一列数据即是数据标签，比如测斜数据工作表中的深度列
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_Depth = 1;
				/// <summary>
				/// 测斜位移值所在的第一列，也对应了第一个监测日期所在的列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_FirstData_Displacement = 2;
			}
			
			/// <summary>
			/// 监测数据中的支撑轴力
			/// </summary>
			/// <remarks></remarks>
			public class Mnt_AxialForce
			{
				
			}
			
			/// <summary>
			/// 监测数据中，除测斜数据与支撑轴力以外的监测数据
			/// </summary>
			/// <remarks></remarks>
			public class Mnt_Others
			{
				/// <summary>
				/// 记录施工日期的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNumForDate = 1;
				
				/// <summary>
				/// 定义监测数据中的数据区域（包含x轴的深度数据）的起始单元格的位置：一般为“A3”
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_FirstData_WithoutDate = 2;
				/// <summary>
				/// 第一列数据即是数据标签，比如测斜数据工作表中的深度列，或者是其他数据类型的测点编号所在列
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_PointsTag = 1;
				/// <summary>
				/// 第一个监测数据所在列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_FirstData_Displacement = 2;
			}
			
			/// <summary>
			/// 数据库文件中，剖面标高项的数据格式
			/// </summary>
			/// <remarks></remarks>
			public class DB_Sectional
			{
				
				/// <summary>
				/// 工作表中记录基坑ID信息的单元格所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_ID = 2;
				
				/// <summary>
				/// 工作表中第一个基坑ID值的单元格的列号，此单元格一般位于合并单元格内，
				/// 返回的列号是合并单元格中的第一个单元格的列号。
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_FirstID = 2;
				
				/// <summary>
				/// 工作表中记录的第一个结构构件所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_FirstItem = 3;
				
				/// <summary>
				/// 工作表“剖面标高”中记录标高值的最后一行的行号，包括这一行
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_EndRowInElevation = 15;
				
				/// <summary>
				/// 在数据库中用于标识支撑构件的字符
				/// </summary>
				/// <remarks></remarks>
				public const string identifier_Ground = "GRD"; // "地面"
				/// <summary>
				/// 在数据库中用于标识支撑构件的字符
				/// </summary>
				/// <remarks></remarks>
				public const string identifier_struts = "S"; // "支撑"
				
				/// <summary>
				/// 在数据库中用于标识楼板构件的字符
				/// </summary>
				/// <remarks></remarks>
				public const string identifier_Floor = "F"; // "楼板"
				
				/// <summary>
				/// 在数据库中用于标识基坑的底部的字符
				/// </summary>
				/// <remarks></remarks>
				public const string identifier_ExcavationBottom = "BTM"; // "基坑底"
				
				/// <summary>
				/// 在数据库中用于标识基坑的底板的顶部的字符
				/// </summary>
				/// <remarks></remarks>
				public const string identifier_TopOfBottomSlab = "DBD"; // "底板顶"
				
			} //DB_Sectional
			
			/// <summary>
			/// 数据库文件中的施工进度表
			/// </summary>
			/// <remarks></remarks>
			public class DB_Progress
			{
				
				/// <summary>
				/// 每个工作表对应的基坑的地块名称（如：A1、B、C1等）
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_ExcavTag = 1;
				
				/// <summary>
				/// 每一个不同方位的基坑区域的标签（比如：普遍区域、东侧）
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_ExcavPosition = 2;
				
				/// <summary>
				/// 记录基坑ID的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_ExcavID = 3;
				
				/// <summary>
				/// 基坑区域开挖到坑底标高时的日期
				/// </summary>
				/// <remarks>用来在绘制剖面标高的矩形时，
				/// 根据当天是否已经开挖到坑底来设置矩形形状的颜色与构件标志的线形等</remarks>
				public const byte RowNum_BottomDate = 4;
				
				/// <summary>
				/// 第一个日期所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_TheFirstDay = 5;
				
				/// <summary>
				/// 记录施工进度日期的数据列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_DateList = 1;
				
				/// <summary>
				/// 记录的第一个基坑区域的列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_theFirstRegion = 4;
				
			} //DB_Progress
			
			/// <summary>
			/// "开挖分块"工作表
			/// </summary>
			/// <remarks>用来记录Visio图形中每一个分块区域所对应的图形的形状ID，以及每一个分块图形相关的信息</remarks>
			public class DB_ExcavRegionForVisio
			{
				/// <summary>
				/// 工作表“开挖分块”中的“形状名”所在的列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_ShapeID = 1;
				/// <summary>
				/// 工作表“开挖分块”中的“完成日期”所在的列号
				/// </summary>
				/// <remarks></remarks>
				public const byte ColNum_FinishedDate = 2;
				/// <summary>
				/// 第一个形状数据所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_FirstShape = 2;
				
			} //DB_ExcavRegionForVisio
			
			/// <summary>
			/// “开挖工况”工作表
			/// </summary>
			/// <remarks></remarks>
			public class DB_WorkingStage
			{
				
				/// <summary>
				/// 基坑区域名称所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_RegionName = 1;
				/// <summary>
				/// 第一个施工工况数据所在的行号
				/// </summary>
				/// <remarks></remarks>
				public const byte RowNum_FirstStage = 3;
				/// <summary>
				/// 每一个基坑区域的施工工况信息所占的列数
				/// </summary>
				/// <remarks></remarks>
				public const byte ColCount_EachRegion = 3;
				// ''' <summary>
				// ''' 记录的有效信息的第一列数据所在的列号
				// ''' </summary>
				// ''' <remarks></remarks>
				//Const ColNum_FirstData As Byte = 1
				/// <summary>
				/// 在每一个基坑区域的工况信息的区域中，工况描述信息所在的相对列号
				/// </summary>
				/// <remarks></remarks>
				public const byte Index_Description = 1;
				/// <summary>
				/// 在每一个基坑区域的工况信息的区域中，施工日期信息所在的相对列号
				/// </summary>
				/// <remarks></remarks>
				public const byte Index_ConstructionDate = 2;
				/// <summary>
				/// 在每一个基坑区域的工况信息的区域中，开挖标高信息所在的相对列号
				/// </summary>
				/// <remarks></remarks>
				public const byte Index_Elevation = 3;
				
			}
		} //DataFormatInExcel
		
		public class xmlNodeNames
		{
			/// <summary>
			/// 记录XML文档中的各种节点的名称
			/// </summary>
			/// <remarks>xmlNode1表示根节点下的第一级节点，xmlNode2表示根节点下的第二级节点，依此类推；
			/// 节点名称中不能包含空字符</remarks>
			public class DataBasePath //DataBasePath
			{
				/// <summary>
				/// 所有项目文件的数据库的元素的父节点
				/// </summary>
				/// <remarks></remarks>
				public const string Nd1_DataBasePaths = "DataBasePaths";
				/// <summary>
				/// DataBase中所包含的每一个Excel工作簿的路径
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_WorkbooksInProject = "WorkbooksInProject";
				/// <summary>
				/// 工作簿文件的绝对路径
				/// </summary>
				/// <remarks></remarks>
				public const string Nd3_FilePath = "FilePath";
				/// <summary>
				/// 工作表名称
				/// </summary>
				/// <remarks></remarks>
				public const string Nd3_SheetName = "SheetName";
				/// <summary>
				/// 施工进度
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_Progress = "ConstructionProgress";
				/// <summary>
				/// 开挖剖面图
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_SectionalView = "SectionalView";
				/// <summary>
				/// 开挖平面图
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_PlanView = "PlanView";
				/// <summary>
				/// 测点坐标
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_PointCoordinates = "PointCoordinates";
				/// <summary>
				/// 开挖工况
				/// </summary>
				/// <remarks></remarks>
				public const string Nd2_WorkingStage = "WorkingStage";
			}
			
			/// <summary>
			/// Visio平面图中的监测点绘制与定位
			/// </summary>
			/// <remarks></remarks>
			public class VisioPlanView_MonitorPoints
			{
				public const string Nd1_MonitorPoints = "Visio平面图中的监测点绘制与定位";
				public const string Nd2_ShapeName_MonitorPointTag = "ShapeName_MonitorPointTag";
				public const string Nd2_pt_Visio_BottomLeft_ShapeID = "pt_Visio_BottomLeft_ShapeID";
				public const string Nd2_pt_Visio_UpRight_ShapeID = "pt_Visio_UpRight_ShapeID";
				public const string Nd2_pt_CAD_BottomLeft_X = "pt_CAD_BottomLeft_X";
				public const string Nd2_pt_CAD_BottomLeft_Y = "pt_CAD_BottomLeft_Y";
				public const string Nd2_pt_CAD_UpRight_X = "pt_CAD_UpRight_X";
				public const string Nd2_pt_CAD_UpRight_Y = "pt_CAD_UpRight_Y";
			}
		}
	} //M_Constants
	
}
