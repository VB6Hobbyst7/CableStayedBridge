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
// End of VB project level imports

using Microsoft.Office.Interop.Excel;
using CableStayedBridge.Miscellaneous;
using CableStayedBridge.Constants;
using CableStayedBridge.GlobalApp_Form;

namespace CableStayedBridge
{
	public abstract class clsDrawing_Mnt_StaticBase : ClsDrawing_Mnt_Base
	{
		
#region   ---  Declarations & Definitions
		
#region   ---  Properties
		
		protected abstract override ChartSize ChartSize_sugested {get; set;}
		
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
		/// </summary>
		/// <remarks></remarks>
		protected double[] F_arrAllDate;
		
		/// <summary>
		/// 以每一条Series对象来索引此数据系列中的Y轴的数据，
		/// 在表示Y轴数据的Object()中，其元素的个数必须要与F_arrAllDate中的元素个数相等。
		/// </summary>
		/// <remarks></remarks>
		protected Dictionary<Series, object[]> F_dicSeries;
		
#endregion
		
#endregion
		
#region   ---  构造函数与窗体的加载、打开与关闭
		
		/// <summary>
		/// 构造函数，构造时一定要设置好字典F_dicFourSeries的值。
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
		public clsDrawing_Mnt_StaticBase(Worksheet DataSheet, Chart DrawingChart, Cls_ExcelForMonitorDrawing ParentApp, 
			DrawingType type, bool CanRoll, TextFrame2 Info, 
			MonitorInfo DrawingTag, MntType MonitorType, 
			double[] Alldate) : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType)
		{
			// VBConversions Note: Non-static class variable initialization is below.  Class variables cannot be initially assigned non-static values in C#.
			myChart = this.Chart;
			
			//
			this.F_arrAllDate = Alldate;
			this.currentPointsCount = Alldate.Length;
		}
		
#endregion
		
		/// <summary>
		/// Excel图表中，静态曲线图的数据系列中，每条曲线中所显示的数据点个数
		/// </summary>
		/// <remarks></remarks>
		private int currentPointsCount;
		private Chart myChart; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
		private void Chart_DoubleClick(int elementID, int arg1, int _arg2, ref 
			bool cancel)
		{
			//控制界面显示
			cancel = true; //表示此事件屏蔽默认的双击事件
			SpeedMode(myChart, F_arrAllDate, this.F_dicSeries);
		}
		
		
		protected void SpeedMode(Chart Chart, System.Double[] arrAllDate, Dictionary<Series, object[]> dicSeries)
		{
			Worksheet sht = Chart.Parent.Parent;
			sht.Range("A1").Activate(); //取消图表上的对象的选择
			//执行SpeedMode相关的事件
			//
			double startday = System.Convert.ToDouble(arrAllDate.First);
			double endday = System.Convert.ToDouble(arrAllDate.Last);
			int AllDateCount = arrAllDate.Length;
			int pointCount = 0;
			//
			string strPointsCount = "";
			string strInputBoxTitle = "Speed Mode";
			
			strPointsCount = System.Convert.ToString(this.Application.InputBox(Prompt: "设置曲线中显示的测点个数" + "\r\n" 
				+ "当前记录天数为" + System.Convert.ToString(currentPointsCount) + "天" + "\r\n" 
				+ "最大记录天数为" + System.Convert.ToString((arrAllDate.Length - 1) + 1) + "天", Title: ref 
				strInputBoxTitle, Type: 1));
			try
			{
				pointCount = int.Parse(strPointsCount); //可能会出现数据类型转换异常
				if (pointCount >= AllDateCount)
				{
					pointCount = AllDateCount; //开始执行操作
				}
				currentPointsCount = pointCount;
				//获取按指定点数进行划分的时间区段长度
				float unit = (float) ((double) AllDateCount / pointCount);
				
				//记录要绘制的日期的数组，数组中记录这些天对应的列号
				int[] slctCol = new int[pointCount - 1 + 1];
				double[] slctDate = new double[pointCount - 1 + 1];
				
				//-------------------- 按指定的日期间隔得到对应的日期与日期在数组中的列号
				double referenceColumn = 0;
				int iselected = 0;
				for (int icol = 0; icol <= AllDateCount - 1; icol++)
				{
					//此处默认所有施工日期的数据中的日期是按从小到大的顺序进行排列的。
					//不果不是按此方法排列，则得到的结果中监测数据还是能与日期对应上，但是选择的日期可能会比较混乱
					//不具有均匀分布的特征
					if (icol >= referenceColumn)
					{
						slctCol[iselected] = icol;
						slctDate[iselected] = arrAllDate[icol];
						iselected++; //满足条件的记录结果位置+1
						referenceColumn = referenceColumn + unit;
					}
					
				} //下一天
				
				//按上面得到的列号来构造新的数组，以得到新的日期排列下的每一条曲线的数据
				byte irow = (byte) 0;
				object[] DataInSelectedDays = new object[(slctCol.Length - 1) + 1];
				foreach (Series curve in dicSeries.Keys)
				{
					int idata = 0;
					foreach (int selectedColumn in slctCol)
					{
						DataInSelectedDays[idata] = dicSeries.Item(curve)[selectedColumn];
						idata++;
					}
					curve.XValues = slctDate;
					curve.Values = DataInSelectedDays;
					irow++;
				}
			}
			catch (Exception)
			{
				//MessageBox.Show("输入的格式不是合法的数值格式，请重新输入", "tip", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			}
		}
		
	}
	
}
