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

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;


namespace CableStayedBridge
{
	namespace DataBase
	{
		
#region   ---  Enumerators
		
		/// <summary>
		/// 基坑中各种构件的类型
		/// </summary>
		/// <remarks></remarks>
		public enum ComponentType
		{
			/// <summary>
			/// 自然地面
			/// </summary>
			/// <remarks></remarks>
			Ground,
			/// <summary>
			/// 支撑
			/// </summary>
			/// <remarks></remarks>
			Strut,
			/// <summary>
			/// 楼板
			/// </summary>
			/// <remarks></remarks>
			Floor,
			/// <summary>
			/// 基坑底板顶部
			/// </summary>
			/// <remarks></remarks>
			TopOfBottomSlab,
			/// <summary>
			/// 基坑底部
			/// </summary>
			/// <remarks></remarks>
			ExcavationBottom,
			/// <summary>
			/// 其他
			/// </summary>
			/// <remarks></remarks>
			Others
		}
		
#endregion
		
		/// <summary>
		/// 在“剖面标高”工作表中，每一个基坑ID所对应的相关信息
		/// </summary>
		/// <remarks></remarks>
		public class clsData_ExcavationID
		{
			
			private float _ExcavationBottom;
			/// <summary>
			/// 基坑底部标高
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public float ExcavationBottom
			{
				get
				{
					return _ExcavationBottom;
				}
			}
			
			private float _BottomFloor;
			/// <summary>
			/// 底板顶部标高
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public float BottomFloor
			{
				get
				{
					return _BottomFloor;
				}
			}
			
			private Component[] _Components;
			/// <summary>
			/// 记录基坑ID中对应的每一个构件项目与其对应标高的数组,
			/// 数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Component[] Components
			{
				get
				{
					return _Components;
				}
			}
			
			private string _name;
			/// <summary>
			/// 此基坑ID的ID名，如“A1-1”
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string Name
			{
				get
				{
					return this._name;
				}
			}
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="name">此基坑ID的ID名，如“A1-1”</param>
			/// <param name="pitBottom">基坑底部标高</param>
			/// <param name="bottomfloor">底板顶部标高</param>
			/// <param name="cmpnts">记录基坑ID下的结构构件及其标高的数组</param>
			/// <remarks></remarks>
			public clsData_ExcavationID(string name, float pitBottom, float bottomfloor, Component[] cmpnts)
			{
				this._ExcavationBottom = pitBottom;
				this._BottomFloor = bottomfloor;
				this._Components = cmpnts;
				this._name = name;
			}
			
		}
		
		/// <summary>
		/// 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
		/// </summary>
		/// <remarks></remarks>
		public class clsData_ProcessRegionData
		{
			
			/// <summary>
			/// 基坑区域的施工进度中，每一个子区域的开挖标高数据对应的Range对象，
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
			/// </summary>
			/// <remarks></remarks>
			private Microsoft.Office.Interop.Excel.Range _Range_Process;
			/// <summary>
			/// 基坑区域的施工进度中，每一个子区域的开挖标高数据对应的Range对象，
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Range Range_Process
			{
				get
				{
					return _Range_Process;
				}
			}
			
			/// <summary>
			/// 基坑区域的施工进度中，每一个子区域的开挖时间对应的Range对象，
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
			/// </summary>
			/// <remarks></remarks>
			private Microsoft.Office.Interop.Excel.Range _Range_Date;
			/// <summary>
			/// 基坑区域的施工进度中，每一个子区域的开挖时间对应的Range对象，
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Range Range_Date
			{
				get
				{
					return this._Range_Date;
				}
			}
			
			/// <summary>
			/// 在基坑区域的施工进度中，每一天的日期所对应的开挖标高。
			/// </summary>
			/// <remarks></remarks>
			private SortedList<DateTime, Single> _Date_Elevation;
			/// <summary>
			/// 在基坑区域的施工进度中，每一天的日期所对应的开挖标高。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public SortedList<DateTime, float> Date_Elevation
			{
				get
				{
					return this._Date_Elevation;
				}
			}
			
			private bool _blnHasBottomDate;
			/// <summary>
			/// 指示在数据库文件中，指定的开挖区域是否已经开挖到了基坑底
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public bool HasBottomDate
			{
				get
				{
					return _blnHasBottomDate;
				}
			}
			
			private DateTime _BottomDate;
			/// <summary>
			/// 在数据库文件中，指定的开挖区域开挖到基坑底的日期
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public DateTime BottomDate
			{
				get
				{
					return _BottomDate;
				}
			}
			
			private clsData_ExcavationID _ExcavationID;
			/// <summary>
			/// 基坑区域对应的基坑ID的信息
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public clsData_ExcavationID ExcavationID
			{
				get
				{
					return _ExcavationID;
				}
			}
			
			private string _ExcavName;
			/// <summary>
			/// 基坑区域所在的基坑的名称，如A1、B、C1等
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string ExcavName
			{
				get
				{
					return this._ExcavName;
				}
			}
			
			/// <summary>
			/// 基坑区域的分块名称，如普遍区域、东南侧、西侧等
			/// </summary>
			/// <remarks></remarks>
			private string _ExcavPosition;
			/// <summary>
			/// 基坑区域的分块名称，如普遍区域、东南侧、西侧等
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string ExcavPosition
			{
				get
				{
					return this._ExcavPosition;
				}
			}
			
			private string _description;
			/// <summary>
			/// 关于此基坑区域的描述，如“A1:普遍区域”
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string description
			{
				get
				{
					return this._description;
				}
				set
				{
					this._description = value;
				}
			}
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="ExcavName">基坑区域所在的基坑的名称，如A1、B、C1等</param>
			/// <param name="ExcavPosition">基坑区域的分块名称，如普遍区域、东南侧、西侧等</param>
			/// <param name="ExcavationID">基坑区域对应的基坑ID的信息</param>
			/// <param name="strBottomDate">在数据库文件中，指定的开挖区域开挖到基坑底的日期，在构造函数中进行数据类型的转换</param>
			/// <param name="Range_Process">基坑区域的施工进度中，每一个子区域的开挖标高对应的Range对象,
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
			/// <param name="Range_Date">基坑区域的施工进度中，每一个子区域的开挖日期对应的Range对象,
			/// 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
			/// <param name="Date_Elevation">在基坑区域的施工进度中，每一天的日期所对应的开挖标高。</param>
			/// <remarks></remarks>
			public clsData_ProcessRegionData(string ExcavName, string ExcavPosition, clsData_ExcavationID ExcavationID, string strBottomDate, Microsoft.Office.Interop.Excel.Range Range_Process, Microsoft.Office.Interop.Excel.Range Range_Date, SortedList<DateTime, Single> Date_Elevation)
			{
				// ----------------------------------------
				clsData_ProcessRegionData with_1 = this;
				with_1._ExcavationID = ExcavationID;
				with_1._ExcavName = ExcavName;
				with_1._ExcavPosition = ExcavPosition;
				with_1._description = ExcavName + ":" + ExcavPosition;
				with_1._Range_Date = Range_Date;
				with_1._Range_Process = Range_Process;
				with_1._Date_Elevation = Date_Elevation;
				
				//将数据库文件中，指定的开挖区域开挖到基坑底的日期的单元格的数据，转换为日期类型
				//如果转换错误，即此单元格为空，或者单元格数据格式不能转换为日期格式，那说明此基坑区域还没有开挖到基坑底部标高
				try //对“2014/11/26”形式的日期进行转换
				{
					this._BottomDate = DateTime.Parse(strBottomDate);
					this._blnHasBottomDate = true;
				}
				catch (Exception)
				{
					try //对“”形式的日期进行转换，比如“41969”对应于日期的“2014/11/26”
					{
						this._BottomDate = DateTime.FromOADate(double.Parse(strBottomDate));
					}
					catch (Exception)
					{
						//如果上面两种转换方法都不能将单元格中的数据转换为日期类型，则认为此基坑区域还没有开挖到基坑底部标高
						this._blnHasBottomDate = false;
					}
				}
				
			}
			
		}
		
		/// <summary>
		/// 记录Visio绘图中所有开挖分块的形状的ID值及每一个分块的完成日期的两个数组，
		/// 这两个数组的元素个数必须相同。
		/// </summary>
		/// <remarks></remarks>
		public class clsData_ShapeID_FinishedDate
		{
			/// <summary>
			/// 列举Visio绘图中所有开挖分块的形状的ID值的数组
			/// </summary>
			/// <remarks></remarks>
			public int[] ShapeID;
			/// <summary>
			/// 每一个开挖分块所对应的开挖完成的日期
			/// </summary>
			/// <remarks></remarks>
			public DateTime[] FinishedDate;
			
			/// <summary>
			/// 构造函数，ShapeID与FinishedDate这两个数组中的元素个数必须相同，而且其中第一个元素的下标值为0。
			/// </summary>
			/// <param name="ShapeID">列举Visio绘图中所有开挖分块的形状的ID值的数组</param>
			/// <param name="FinishedDate">每一个开挖分块所对应的开挖完成的日期的数组</param>
			/// <remarks></remarks>
			public clsData_ShapeID_FinishedDate(int[] ShapeID, DateTime[] FinishedDate)
			{
				if (ShapeID.Length != FinishedDate.Length)
				{
					MessageBox.Show("形状ID的数组与对应的完成日期的数组的元素个数必须相同！", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					this.ShapeID = ShapeID;
					this.FinishedDate = FinishedDate;
				}
			}
		}
		
		/// <summary>
		/// 某一基坑的工况信息，它主要包括此基坑在某一特定的日期开挖到了什么状态，当时的开挖深度值
		/// </summary>
		/// <remarks></remarks>
		public class clsData_WorkingStage
		{
			
			/// <summary>
			/// 工况描述
			/// </summary>
			/// <remarks></remarks>
			public string Description;
			
			/// <summary>
			/// 施工日期
			/// </summary>
			/// <remarks></remarks>
			public DateTime ConstructionDate;
			
			/// <summary>
			/// 开挖标高，单位为m，书写格式为：4.200，不带单位。
			/// 注意：是开挖标高值，不是开挖深度值。
			/// </summary>
			/// <remarks></remarks>
			public float Elevation;
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="Description">工况描述</param>
			/// <param name="ConstructionDate">施工日期</param>
			/// <param name="Elevation">开挖标高，单位为m，书写格式为：4.200，不带单位。
			/// 注意：是开挖标高值，不是开挖深度值。</param>
			/// <remarks></remarks>
			public clsData_WorkingStage(string Description, DateTime ConstructionDate, float Elevation)
			{
				clsData_WorkingStage with_1 = this;
				with_1.Description = Description;
				with_1.ConstructionDate = ConstructionDate;
				with_1.Elevation = Elevation;
			}
			
		}
		
		/// <summary>
		/// 在基坑设计中，每一个不同的基坑ID区域，所包含的构件信息
		/// </summary>
		/// <remarks></remarks>
		public struct Component
		{
			/// <summary>
			/// 此构件的类型
			/// </summary>
			/// <remarks></remarks>
			public ComponentType Type;
			/// <summary>
			/// 对于此构件的描述
			/// </summary>
			/// <remarks></remarks>
			public string Description;
			/// <summary>
			/// 此构件的特征标高
			/// </summary>
			/// <remarks></remarks>
			public float Elevation;
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="Type">此构件的类型</param>
			/// <param name="Description">对于此构件的描述</param>
			/// <param name="Elevation">此构件的特征标高</param>
			/// <remarks></remarks>
			public Component(ComponentType Type, string Description, float Elevation)
			{
				Component with_1 = this;
				with_1.Type = Type;
				with_1.Description = Description;
				with_1.Elevation = Elevation;
			}
			
		}
		
		/// <summary>
		/// 文件中要保存的内容或对象
		/// </summary>
		/// <remarks></remarks>
		public class clsData_FileContents
		{
			
			/// <summary>
			/// 项目文件中记录的所有工作簿
			/// </summary>
			/// <remarks></remarks>
			public List<Workbook> lstWkbks {get; set;}
			
			/// <summary>
			/// 施工进度工作表
			/// </summary>
			/// <remarks></remarks>
			public List<Worksheet> lstSheets_Progress {get; set;}
			/// <summary>
			/// 开挖平面图工作表
			/// </summary>
			/// <remarks></remarks>
			public Worksheet Sheet_PlanView {get; set;}
			/// <summary>
			/// 开挖剖面图工作表
			/// </summary>
			/// <remarks></remarks>
			public Worksheet Sheet_Elevation {get; set;}
			/// <summary>
			/// 测点坐标工作表
			/// </summary>
			/// <remarks></remarks>
			public Worksheet Sheet_PointCoordinates {get; set;}
			/// <summary>
			/// 开挖工况信息工作表
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
			public Worksheet Sheet_WorkingStage {get; set;}
			
			public clsData_FileContents()
			{
				this.lstSheets_Progress = new List<Worksheet>();
				this.lstWkbks = new List<Workbook>();
			}
		}
		
	}
}
