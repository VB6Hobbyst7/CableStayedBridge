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
using CableStayedBridge.Constants;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop.Excel;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.GlobalApp_Form;


namespace CableStayedBridge
{
	namespace All_Drawings_In_Application
	{
		/// <summary>
		/// 在测斜数据中，某一测点在整个施工跨度内，每一天的位移最值，以及对应的深度
		/// </summary>
		/// <remarks></remarks>
		public class ClsDrawing_Mnt_MaxMinDepth : clsDrawing_Mnt_StaticBase
		{
			
#region   ---  Types
			
			
			/// <summary>
			/// 在位移极值走势图中，所需要的所有数据
			/// </summary>
			/// <remarks></remarks>
			public class DateMaxMinDepth
			{
				/// <summary>
				/// 这里的Date数据不能用Date来保存，而应该用函数ToOADate将Date转换为等效的Double类型。
				/// 这是因为：在Excel的Chart中，如果要设置一个坐标轴标签以日期型显示，
				/// 则不能将其以Date型数组进行赋值，而应该以对应的Double类型的数组进行赋值。
				/// </summary>
				/// <remarks></remarks>
				public double[] ConstructionDate;
				public object[] Max;
				public object[] Min;
				public object[] Depth_Max;
				public object[] Depth_Min;
				
				/// <summary>
				/// 输入的数组的元素个数必须相同。元素类型必须为Object，以避免出现因为没有数据而被误写为0.0的错误。
				/// </summary>
				/// <param name="ConstructionDate">
				/// 这里的Date数据不能用Date来保存，而应该用函数ToOADate将Date转换为等效的Double类型。
				/// 这是因为：在Excel的Chart中，如果要设置一个坐标轴标签以日期型显示，
				/// 则不能将其以Date型数组进行赋值，而应该以对应的Double类型的数组进行赋值。</param>
				/// <param name="max"></param>
				/// <param name="min"></param>
				/// <param name="depth_max"></param>
				/// <param name="depth_min"></param>
				/// <remarks></remarks>
				public DateMaxMinDepth(double[] ConstructionDate, object[] max, object[] min, object[] 
					depth_max, object[] depth_min)
				{
					this.ConstructionDate = ConstructionDate;
					this.Max = max;
					this.Min = min;
					this.Depth_Max = depth_max;
					this.Depth_Min = depth_min;
					
				}
			}
			
			
#endregion
			
#region   ---  Properties
			
			/// <summary>
			/// 绘图界面与画布的尺寸
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
protected override ChartSize ChartSize_sugested
			{
				get
				{
					return new ChartSize(Data_Drawing_Format.Drawing_Incline_DMMD.ChartHeight, 
						Data_Drawing_Format.Drawing_Incline_DMMD.ChartWidth, 
						Data_Drawing_Format.Drawing_Incline_DMMD.MarginOut_Height, 
						Data_Drawing_Format.Drawing_Incline_DMMD.MarginOut_Width);
				}
				set
				{
					ExcelFunction.SetLocation_Size(this.ChartSize_sugested, this.Chart, this.Application);
				}
			}
			
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="DataSheet">图表对应的数据工作表</param>
			/// <param name="DrawingChart">Excel图形所在的Chart对象</param>
			/// <param name="ParentApp">此图表所在的Excel类的实例对象</param>
			/// <param name="type">此图表所属的类型，由枚举drawingtype提供</param>
			/// <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
			/// <param name="Info">图表中用来显示相关信息的那个文本框对象</param>
			/// <param name="DrawingTag">每一个监测曲线图的相关信息</param>
			/// <param name="MonitorType">监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等</param>
			/// <remarks></remarks>
			public ClsDrawing_Mnt_MaxMinDepth(Worksheet DataSheet, Chart 
				DrawingChart, Cls_ExcelForMonitorDrawing ParentApp, 
				DrawingType type, bool CanRoll, TextFrame2 Info, 
				MonitorInfo DrawingTag, MntType MonitorType, double[] AllDate, DateMaxMinDepth Data) : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, AllDate)
			{
				//  -----------------------------------
				//
				this.F_dicSeries = new Dictionary<Series, object[]>();
				SeriesCollection Sc = DrawingChart.SeriesCollection();
				this.F_dicSeries.Add(Sc.Item(1), Data.Max);
				this.F_dicSeries.Add(Sc.Item(2), Data.Min);
				this.F_dicSeries.Add(Sc.Item(3), Data.Depth_Max);
				this.F_dicSeries.Add(Sc.Item(4), Data.Depth_Min);
			}
			
		}
	}
}
