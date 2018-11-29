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
using CableStayedBridge.Miscellaneous.GeneralMethods;
using CableStayedBridge.All_Drawings_In_Application;


namespace CableStayedBridge
{
	namespace All_Drawings_In_Application
	{
		public class ClsDrawing_Mnt_Static : clsDrawing_Mnt_StaticBase
		{
			
#region   ---  Constants
			//图表网格与坐标值划分
			public const byte cstChartParts_Y = 10; //图表Y轴（位移）划分的区段数
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
					return new ChartSize(Data_Drawing_Format.Drawing_Mnt_Others.ChartHeight, 
						Data_Drawing_Format.Drawing_Mnt_Others.ChartWidth, 
						Data_Drawing_Format.Drawing_Mnt_Others.MarginOut_Height, 
						Data_Drawing_Format.Drawing_Mnt_Others.MarginOut_Width);
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
			public ClsDrawing_Mnt_Static(Worksheet DataSheet, Chart DrawingChart, 
				Cls_ExcelForMonitorDrawing ParentApp, 
				DrawingType type, bool CanRoll, 
				TextFrame2 Info, MonitorInfo DrawingTag, MntType MonitorType, 
				Dictionary<Series, object[]> AllselectedData, double[] arrAllDate) : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, arrAllDate)
			{
				//  -----------------------------------
				//
				this.F_dicSeries = AllselectedData;
			}
			
		}
	}
}
