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
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.Constants;
using CableStayedBridge.DataBase;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

//using DAMIE.Constants.Data_Drawing_Format;

namespace CableStayedBridge
{
	
	public class ClsDrawing_ExcavationElevation : IAllDrawings, IRolling
	{
		
#region   ---  Declarations and Definitions
		
#region   ---  Properties
		
#region   ---  Excel中的对象
		
		private Application P_ExcelApp;
		/// <summary>
		/// 进行绘图的Excel程序
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Application Application
		{
			get
			{
				if (this.P_ExcelApp == null)
				{
					this.P_ExcelApp = new Application();
					this.P_ExcelApp.WorkbookBeforeClose += this.Application_Quit;
				}
				return this.P_ExcelApp;
			}
			set
			{
				//不弹出警告对话框
				value.DisplayAlerts = false;
				//获取Excel的进程
				this.P_ExcelApp = value;
				this.P_ExcelApp.WorkbookBeforeClose += this.Application_Quit;
			}
		}
		
		/// <summary>
		/// 进行绘图的工作表
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Worksheet P_Sheet_Drawing;
		/// <summary>
		/// 进行绘图的工作表
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Worksheet Sheet_Drawing
		{
			get
			{
				return this.P_Sheet_Drawing;
			}
		}
		
		/// <summary>
		/// 进行绘图的Chart对象
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Chart P_Chart;
		/// <summary>
		/// 进行绘图的Chart对象
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Chart Chart
		{
			get
			{
				return this.P_Chart;
			}
		}
		
		/// <summary>
		/// 记录数据信息的文本框
		/// </summary>
		/// <remarks></remarks>
		private TextFrame2 F_textbox_Info;
		/// <summary>
		/// 记录数据信息的文本框
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
public TextFrame2 Information
		{
			get
			{
				return F_textbox_Info;
			}
		}
		
#endregion
		
		private DrawingType P_Type;
public DrawingType Type
		{
			get
			{
				return this.P_Type;
			}
		}
		
		private long P_UniqueID;
public long UniqueID
		{
			get
			{
				return P_UniqueID;
			}
		}
		
		private DateSpan P_DateSpan;
public DateSpan DateSpan
		{
			get
			{
				return this.P_DateSpan;
			}
			set
			{
				this.P_DateSpan = value;
				// 扩展MainForm.TimeSpan的区间
				GlobalApplication.Application.refreshGlobalDateSpan(value);
			}
		}
		
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 此图表中选择的所有的基坑区域对象
		/// </summary>
		/// <remarks></remarks>
		private List<clsData_ProcessRegionData> F_Regions;
		
		/// <summary>
		/// 记录每一个基坑区域的坑底标高的数据系列
		/// </summary>
		/// <remarks></remarks>
		private Series F_Series_Static;
		
		/// <summary>
		/// 记录每一个基坑区域的即时开挖标高的数据系列
		/// </summary>
		/// <remarks></remarks>
		private Series F_Series_Depth;
		
#endregion
		
#endregion
		
		/// <summary>
		/// 构造函数
		/// </summary>
		/// <param name="Series_Static">表示基坑区域的坑底深度的数据系列</param>
		/// <param name="Series_Depth">表示基坑区域的即时开挖标高的数据系列</param>
		/// <param name="ChosenRegion">此绘图中所包含的矩形方块与对应的数据范围</param>
		/// <param name="textbox">记录信息的文本框</param>
		/// <param name="type">此图表所属的类型，由枚举DrawingType提供</param>
		public ClsDrawing_ExcavationElevation(Microsoft.Office.Interop.Excel.Series Series_Static, Series Series_Depth, 
			List<clsData_ProcessRegionData> ChosenRegion, DateSpan DateSpan, 
			Microsoft.Office.Interop.Excel.TextFrame2 textbox, DrawingType type)
		{
			ClsDrawing_ExcavationElevation with_1 = this;
			with_1.F_Series_Static = Series_Static;
			with_1.F_Series_Depth = Series_Depth;
			with_1.F_textbox_Info = textbox;
			with_1.P_Chart = Series_Static.Parent.Parent;
			with_1.P_Sheet_Drawing = this.P_Chart.Parent.parent;
			with_1.Application = this.P_Sheet_Drawing.Application;
			with_1.F_Regions = ChosenRegion;
			with_1.DateSpan = DateSpan;
			with_1.P_Type = type;
			with_1.Application.Caption = "";
			with_1.P_UniqueID = GeneralMethods.GetUniqueID();
			
			// -------------------------------------------------------------
			GlobalApplication.Application.ElevationDrawing = this;
		}
		
		public void Rolling(DateTime dateThisDay)
		{
			object lockobject = new object();
			lock(lockobject)
			{
				UInt16 RegionsCount = this.F_Regions.Count;
				if (RegionsCount > 0)
				{
					this.Application.ScreenUpdating = false; //禁用excel界面
					Series series_Depth = this.F_Series_Depth;
					clsData_ProcessRegionData Region = default(clsData_ProcessRegionData);
					Microsoft.Office.Interop.Excel.Point Pt = default(Microsoft.Office.Interop.Excel.Point);
					float[] Depths = new float[RegionsCount - 1 + 1];
					for (UInt16 i = 0; i <= RegionsCount - 1; i++)
					{
						Region = this.F_Regions.Item(i);
						Pt = series_Depth.Points().item(i + 1);
						if (Region.HasBottomDate)
						{
							if (dateThisDay.CompareTo(Region.BottomDate) > 0) //说明已经开挖到基坑底，并在向上进行结构物的施工
							{
								
								Pt.Format.Fill.ForeColor.RGB = Information.RGB(255, 0, 0);
							}
							else //说明还未开挖到基坑底，并在向下开挖
							{
								
								Pt.Format.Fill.ForeColor.RGB = Information.RGB(0, 0, 255);
							}
						}
						try
						{
							Depths[i] = System.Convert.ToSingle(Region.Date_Elevation[dateThisDay]);
						}
						catch (KeyNotFoundException)
						{
							DateTime ClosestDate = ClsData_DataBase.FindTheClosestDateInSortedList(Region.Date_Elevation.Keys, dateThisDay);
							Depths[i] = System.Convert.ToSingle(Region.Date_Elevation[ClosestDate]);
						}
					}
					series_Depth.Values = Depths;
					//刷新日期放置在最后，以免由于耗时过长而出现误判
					this.F_textbox_Info.TextRange.Text = dateThisDay.ToString(AMEApplication.DateFormat);
				}
				this.Application.ScreenUpdating = true; //刷新excel界面
			}
		}
		
#region   ---  通用子方法
		
		public void Close(bool SaveChanges = false)
		{
			try
			{
				ClsDrawing_ExcavationElevation with_1 = this;
				Microsoft.Office.Interop.Excel.Workbook wkbk = with_1.P_Sheet_Drawing.Parent;
				wkbk.Close(false);
				with_1.Application.Quit();
			}
			catch (Exception ex)
			{
				MessageBox.Show("关闭开挖剖面图出错！" + "\r\n" + ex.Message, 
					"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}
		
		/// <summary>
		/// 绘图界面被删除时引发的事件
		/// </summary>
		/// <remarks></remarks>
		private void Application_Quit(Workbook Wb, ref bool Cancel)
		{
			GlobalApplication.Application.ElevationDrawing = null;
		}
		
#endregion
		
	}
	
}
