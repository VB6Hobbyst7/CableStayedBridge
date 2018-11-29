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
using CableStayedBridge.GlobalApp_Form;


namespace CableStayedBridge
{
	namespace All_Drawings_In_Application
	{
		/// <summary>
		/// 监测曲线图中，除了测斜曲线图外，其他的可以动态滚动的曲线图。
		/// </summary>
		/// <remarks></remarks>
		public class ClsDrawing_Mnt_OtherDynamics : clsDrawing_Mnt_RollingBase
		{
			
#region   ---  常数值定义
			
			//图表网格与坐标值划分
			public const byte cstChartParts_Y = 10; //图表Y轴（位移）划分的区段数
			public const string cstAxisTitle_X_Dynamic = Constants.AxisLabels.Points;
			public const string cstAxisTitle_Y = Constants.AxisLabels.Displacement_mm;
			
#endregion
			
#region   ---  属性值的定义
			
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
			
			private Dictionary<DateTime, object[]> P_dicDate_ChosenDatum = new Dictionary<DateTime, object[]>();
			/// <summary>
			/// 以每一天的日期来索引这一天的监测数据
			/// </summary>
			/// <value></value>
			/// <returns>返回一个字典，其关键字为监测数据表中有数据的每一天的日期，
			/// 对应的值为当天每一个被选择的监测点的监测数据</returns>
			/// <remarks>监测数据只包含列表中选择了的监测点</remarks>
public Dictionary<DateTime, object[]> DateAndDatum
			{
				get
				{
					return P_dicDate_ChosenDatum;
				}
				set
				{
					P_dicDate_ChosenDatum = value;
				}
			}
			
			
			
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="DataSheet">图表对应的数据工作表</param>
			/// <param name="DrawingChart">Excel图形所在的Chart对象</param>
			/// <param name="ParentApp">此图表所在的Excel类的实例对象</param>
			/// <param name="DateSpan">此图表的TimeSpan跨度</param>
			/// <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
			/// <param name="Date_ChosenDatum">一个字典，其关键字为监测数据表中有数据的每一天的日期，
			/// 对应的值为当天每一个被选择的监测点的监测数据，监测数据只包含列表中选择了的监测点</param>
			/// <param name="Info">记录数据信息的文本框</param>
			/// <remarks></remarks>
			public ClsDrawing_Mnt_OtherDynamics(Worksheet DataSheet, Chart DrawingChart, 
				Cls_ExcelForMonitorDrawing ParentApp, DateSpan DateSpan, 
				DrawingType type, bool CanRoll, TextFrame2 Info, 
				MonitorInfo DrawingTag, MntType MonitorType, 
				Dictionary<DateTime, object[]> Date_ChosenDatum, 
				clsDrawing_Mnt_RollingBase.SeriesTag SeriesTag) : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, DateSpan, SeriesTag)
			{
				
				//  ------------------------------------
				//为进行滚动的那条数据曲线添加数据标签
				//在数据点旁边显示数据值
				this.MovingSeries.ApplyDataLabels();
				//设置数据标签的格式
				DataLabels dataLBs = this.MovingSeries.DataLabels();
				dataLBs.NumberFormat = "0.00";
				dataLBs.Format.TextFrame2.TextRange.Font.Size = 8;
				dataLBs.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Information.RGB(0, 0, 0);
				dataLBs.Format.TextFrame2.TextRange.Font.Name = AMEApplication.FontName_TNR;
				
				P_dicDate_ChosenDatum = Date_ChosenDatum; //包括第一列，但是不包括第一行的日期。
				// '' -----对图例进行更新---------
				//Call LegendRefresh(Me.List_HasCurve)
				
			}
			
			/// <summary>
			/// 图形滚动
			/// </summary>
			/// <param name="dateThisday"></param>
			/// <remarks></remarks>
			public override void Rolling(DateTime dateThisday)
			{
				base.F_RollingDate = dateThisday;
				object lockobject = new object();
				lock(lockobject)
				{
					var app = this.Chart.Application;
					app.ScreenUpdating = false;
					
					// ------------------- 绘制监测曲线图
					var Allday = this.P_dicDate_ChosenDatum.Keys;
					//考察选定的日期是否有数据
					TodayState State = default(TodayState);
					DateTime closedDay = default(DateTime);
					//
					if (DateTime.Compare(dateThisday, this.DateSpan.StartedDate) < 0)
					{
						State = TodayState.BeforeStartDay;
					}
					else if (DateTime.Compare(dateThisday, this.DateSpan.FinishedDate) > 0)
					{
						State = TodayState.AfterFinishedDay;
					}
					else if (Allday.Contains(dateThisday)) //如果搜索的那一天有数据
					{
						State = TodayState.DateMatched;
						closedDay = dateThisday;
					}
					else //搜索的那一天没有数据，则查找选定的日期附近最近的一天，并绘制其监测曲线
					{
						State = TodayState.DateNotFound;
						SortedSet<DateTime> sortedlist_AllDays = new SortedSet<DateTime>(Allday);
						closedDay = base.GetClosestDay(sortedlist_AllDays, dateThisday);
					}
					//
					this.CurveRolling(dateThisday, State, closedDay);
					app.ScreenUpdating = true;
				}
			}
			
			/// <summary>
			/// 绘制距选定的日期（包括选择的日期）最近的日期的监测数据
			/// </summary>
			/// <param name="dateThisDay">选定的施工日期</param>
			/// <param name="State">指示选定的日期所处的状态</param>
			/// <param name="Closestday">距离选定日期最近的日期，如果选定的日期在工作表中有数据，则等效于dateThisDay</param>
			/// <remarks></remarks>
			private void CurveRolling(DateTime dateThisDay, TodayState State, DateTime Closestday)
			{
				Series series = this.MovingSeries; // Me.Chart.SeriesCollection(1)
				const string strDateFormat = AMEApplication.DateFormat;
				F_DicSeries_Tag.Item(1).ConstructionDate = dateThisDay; //刷新滚动的曲线所代表的日期
				//绘制曲线图
				//.XValues = dicDate_ChosenDatum.Item(dateThisday)           'X轴的数据
				//
				switch (State)
				{
					
				case TodayState.BeforeStartDay:
					series.Values = new[] {null}; //不能设置为Series.Value=vbEmpty，因为这会将x轴标签中的某一个值设置为0.0。
					series.Name = dateThisDay.ToString(strDateFormat) + " :早于" +
						this.DateSpan.StartedDate.ToString(strDateFormat);
					break;
					
				case TodayState.DateNotFound:
					series.Values = P_dicDate_ChosenDatum.Item(Closestday);
					series.Name = Closestday.ToString(strDateFormat) + "(" + dateThisDay.ToString(strDateFormat) + ":No Data" + ")";
					break;
					
				case TodayState.AfterFinishedDay:
					series.Values = new[] {null};
					series.Name = dateThisDay.ToString(strDateFormat) + " :晚于" +
						this.DateSpan.FinishedDate.ToString(strDateFormat);
					break;
				case TodayState.DateMatched:
					series.Values = P_dicDate_ChosenDatum.Item(dateThisDay);
					series.Name = dateThisDay.ToString(strDateFormat);
					break;
					
			}
		}
	}
	
}
}
