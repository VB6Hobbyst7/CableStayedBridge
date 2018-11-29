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

using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
//using DAMIE.Constants.Data_Drawing_Format;
//using DAMIE.All_Drawings_In_Application.ClsDrawing_Mnt_MaxMinDepth;

//using DAMIE.Miscellaneous.ExcelFunction;

namespace CableStayedBridge
{
	/// <summary>
	/// 绘制测斜数据的曲线图
	/// </summary>
	/// <remarks></remarks>
	public partial class frmDrawing_Mnt_Incline
	{
		
#region   ---  Declarations and Definitions
		
		/// <summary>
		/// 当前进行绘图的数据工作簿发生变化时触发
		/// </summary>
		/// <remarks></remarks>
		private delegate void DataWorkbookChangedEventHandler(Workbook WorkingDataWorkbook);
		private DataWorkbookChangedEventHandler DataWorkbookChangedEvent;
		
		private event DataWorkbookChangedEventHandler DataWorkbookChanged
		{
			add
			{
				DataWorkbookChangedEvent = (DataWorkbookChangedEventHandler) System.Delegate.Combine(DataWorkbookChangedEvent, value);
			}
			remove
			{
				DataWorkbookChangedEvent = (DataWorkbookChangedEventHandler) System.Delegate.Remove(DataWorkbookChangedEvent, value);
			}
		}
		
		
#region   ---  Fields
		//
		Application F_ExcelAppDrawing;
		/// <summary>
		/// 当前用于绘图的数据工作簿
		/// </summary>
		/// <remarks></remarks>
		Workbook F_wkbkData;
		//
		Worksheet F_shtMonitorData;
		Worksheet F_shtDrawing;
		
		/// <summary>
		/// 在工作表中要进行处理的所有数据的范围。
		/// 包括第一列，但是不包括第一行的日期。
		/// </summary>
		/// <remarks></remarks>
		Range F_Data_UsedRange;
		
		/// <summary>
		///工作表中的每一天对应的数据列,
		/// 以日期索引当天数据在工作表中的列号
		/// </summary>
		/// <remarks></remarks>
		Dictionary<DateTime, int> F_dicDate_ColNum = new Dictionary<DateTime, int>();
		
		/// <summary>
		/// 绘图的Chart
		/// </summary>
		/// <remarks></remarks>
		Chart F_myChart;
		
		TextFrame2 F_textbox_Info;
		
		/// <summary>
		/// Chart中第一条监测曲线所对应的相关信息
		/// </summary>
		/// <remarks></remarks>
		private ClsDrawing_Mnt_Incline.SeriesTag_Incline F_TheFirstseriesTag;
		
		/// <summary>
		/// 所绘图的监测数据类型
		/// </summary>
		/// <remarks></remarks>
		private MntType F_MonitorType;
		
		/// <summary>
		/// 是否是要绘制测斜数据的位移最值及对应的深度的曲线图，而不是测斜的动态图
		/// </summary>
		/// <remarks></remarks>
		private bool F_blnMax_Depth;
		//
		private GlobalApplication F_GlobalApp; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
		private APPLICATION_MAINFORM F_MainForm; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
#endregion
		
#endregion
		
#region   ---  窗口的加载与关闭
		
		public frmDrawing_Mnt_Incline()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			ClsData_DataBase.dic_IDtoComponentsChanged += this.RefreshComobox_ExcavationID;
			ClsData_DataBase.ProcessRangeChanged += this.RefreshComobox_ProcessRange;
			ClsData_DataBase.WorkingStageChanged += this.RefreshCombox_WorkingStage;
			// ------------ 设置控件的默认属性
			btnGenerate.Enabled = false;
			chkBoxOpenNewExcel.Checked = true;
			
			// ------------ 设置默认的监测数据类型
			GeneralMethods.SetMonitorType(this.ComboBox_MntType);
			this.F_MonitorType = MntType.Incline;
			this.ComboBox_MntType.Enabled = false;
			this.Label_MntType.Enabled = false;
			// ------------
			this.ComboBoxOpenedWorkbook.DisplayMember = LstbxDisplayAndItem.DisplayMember;
			this.ComboBoxOpenedWorkbook.ValueMember = LstbxDisplayAndItem.ValueMember;
			
		}
		
		
		
		//在关闭窗口时将其隐藏
		public void frmDrawing_Mnt_Incline_FormClosing(object sender, FormClosingEventArgs e)
		{
			//如果是子窗口自己要关闭，则将其隐藏
			//如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
			if (!(e.CloseReason == CloseReason.MdiFormClosing))
			{
				this.Hide();
			}
			e.Cancel = true;
		}
		
		public void frmDrawing_Mnt_Incline_Disposed(object sender, EventArgs e)
		{
			ClsData_DataBase.dic_IDtoComponentsChanged -= this.RefreshComobox_ExcavationID;
			ClsData_DataBase.ProcessRangeChanged -= this.RefreshComobox_ProcessRange;
			ClsData_DataBase.WorkingStageChanged -= this.RefreshCombox_WorkingStage;
		}
		
#endregion
		
		/// <summary>
		/// 生成监测曲线图
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnGenerate_Click(object sender, EventArgs e)
		{
			if (this.F_shtMonitorData != null)
			{
				
				//开始绘图
				if (!this.BGWK_NewDrawing.IsBusy)
				{
					bool blnNewExcelApp = false;
					//用来判断是否要创建新的Excel，以及是否要对新画布所在的Excel进行美化。
					blnNewExcelApp = F_GlobalApp.MntDrawing_ExcelApps.Count == 0 || chkBoxOpenNewExcel.Checked;
					//
					this.BGWK_NewDrawing.RunWorkerAsync(new[] {blnNewExcelApp});
				}
			}
			else
			{
				MessageBox.Show("请选择一个监测数据的工作表", "Warning", MessageBoxButtons.OK);
			}
		}
		
#region   ---  后台线程进行操作
		
		/// <summary>
		/// 生成绘图，此方法是在后台的工作者线程中执行的。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BGW_Generate_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			bool blnNewExcelApp = System.Convert.ToBoolean(e.Argument(0));
			//主程序界面的进度条的UI显示
			F_MainForm.ShowProgressBar_Marquee();
			//执行具体的绘图操作
			if (F_blnMax_Depth)
			{
				try
				{
					Generate_Max_Depth(sheetMonitorData: ref this.F_shtMonitorData, NewExcelApp: ref blnNewExcelApp);
				}
				catch (Exception ex)
				{
					MessageBox.Show("绘制最大值走势图失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.StackTrace, "Error", MessageBoxButtons.OK);
				}
			}
			else
			{
				try
				{
					Generate_DynamicDrawing(sheetMonitorData: ref this.F_shtMonitorData, Components: ref this.F_Components, ProcessRegionData: ref this.F_ProcessRegionData, NewExcelApp: ref blnNewExcelApp);
				}
				catch (Exception ex)
				{
					MessageBox.Show("绘制监测曲线图失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.StackTrace, "Error", MessageBoxButtons.OK);
				}
			}
		}
		
		/// <summary>
		/// 当后台的工作者线程结束（即BGW_Generate_DoWork方法执行完毕）时触发，注意，此方法是在UI线程中执行的。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BGW_Generate_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			
			//在绘图完成后，隐藏进度条
			F_MainForm.HideProgress("Done");
		}
		
#endregion
		
		/// <summary>
		/// 1.绘制动态滚动图
		/// </summary>
		/// <param name="sheetMonitorData">测斜曲线的数据所在的工作表</param>
		/// <param name="Components">与基坑ID相关的信息：结构项目与对应标高</param>
		/// <param name="ProcessRegionData">与此监测数据所对应的基坑的施工进度</param>
		/// <param name="NewExcelApp">用来判断是否要创建新的Excel，以及是否要对新画布所在的Excel进行美化。</param>
		/// <remarks></remarks>
		private void Generate_DynamicDrawing(Worksheet sheetMonitorData, Component[] Components, clsData_ProcessRegionData ProcessRegionData, bool NewExcelApp)
			{
			
			Cls_ExcelForMonitorDrawing myExcelForMonitorDrawing = null;
			//   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
			PrePare(sheetMonitorData, NewExcelApp, myExcelForMonitorDrawing);
			//   ---------------
			//对第一条监测曲线的施工日期进行初始同仁，用来绘制第一条监测曲线
			DateTime Date_theFirstCurve = System.Convert.ToDateTime(F_dicDate_ColNum.Keys(0));
			//Data_UsedRange包括第一列，但是不包括第一行的日期。
			this.F_myChart = DrawDynamicChart(this.F_shtDrawing, Date_theFirstCurve, F_Data_UsedRange);
			
			//设置监测曲线的时间跨度
			DateSpan DtSp = GetDateSpan(this.F_dicDate_ColNum);
			
			
			
			//------------------------------------ 将执行结果传递给mainform的属性中
			//设置图表的Tag属性
			MonitorInfo Tags = GetChartTags(this.F_shtMonitorData);
			ClsDrawing_Mnt_Incline moniSheet = new ClsDrawing_Mnt_Incline(F_shtMonitorData, F_myChart,
				myExcelForMonitorDrawing, DtSp,
				DrawingType.Monitor_Incline_Dynamic, true, F_textbox_Info, Tags, this.F_MonitorType,
				F_dicDate_ColNum, F_Data_UsedRange,
				F_TheFirstseriesTag, ProcessRegionData);
			
			//--------------标高箭头-------------------- 在监测曲线图中绘制标高箭头
			if (Components != null)
			{
				ComponentsAndElevatins(moniSheet.InclineTopElevaion, Components);
			}
			
			// ------------------------------------------------------------------------------------------------
			// ----------动态开挖深度的直线与文本框-------- 根据选择确定是否要绘制动态开挖深度的直线与文本框
			if (ProcessRegionData != null)
			{
				//将表示挖深的直线与文本框赋值给moniSheet的ExcavationDepth属性，以供后面在滚动时进行移动
				RollingDepth_lineAndtextbox(ProcessRegionData, Date_theFirstCurve, moniSheet.ShowLabelsWhileRolling, moniSheet.InclineTopElevaion);
			}
			
			//扩展mainForm.TimeSpan的区间
			GlobalApplication.Application.refreshGlobalDateSpan(moniSheet.DateSpan);
			//-------- 界面显示与美化
			DrawingFinished(NewExcelApp);
			
		}
		
		/// <summary>
		/// 2.绘制最大值的走势图
		/// </summary>
		/// <param name="sheetMonitorData"></param>
		/// <param name="NewExcelApp"></param>
		/// <remarks></remarks>
		private void Generate_Max_Depth(Worksheet sheetMonitorData, bool NewExcelApp)
		{
			Cls_ExcelForMonitorDrawing myExcelForMonitorDrawing = null;
			//   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
			PrePare(sheetMonitorData, NewExcelApp, myExcelForMonitorDrawing);
			//   ---------------
			//以每一天的日期索引当天的测斜位移的极值与对应的深度
			DateMaxMinDepth DMMD = default(DateMaxMinDepth);
			
			DMMD = getMaxMinDepth(this.F_dicDate_ColNum, F_Data_UsedRange);
			if (DMMD == null)
			{
				return;
			}
			//
			Chart cht = DrawDMMDChart(this.F_shtDrawing, DMMD);
			// ------------------------------------------------------------------------------------------------
			//设置图表的Tag属性
			MonitorInfo Tags = GetChartTags(this.F_shtMonitorData);
			ClsDrawing_Mnt_MaxMinDepth Drawing_MMD = new ClsDrawing_Mnt_MaxMinDepth(this.F_shtMonitorData, cht, myExcelForMonitorDrawing,
				DrawingType.Monitor_Incline_MaxMinDepth, false, this.F_textbox_Info, Tags, this.F_MonitorType,
				DMMD.ConstructionDate, DMMD);
			if (this.F_WorkingStage != null)
			{
				DrawWorkingStage(cht, this.F_WorkingStage);
			}
			//-------- 界面显示与美化
			DrawingFinished(NewExcelApp);
		}
		
#region   ---  动态图
		
		/// <summary>
		/// 工作表中的每一天对应的数据列，以日期索引当天数据在工作表中的列号。
		/// 另外，在此函数中，对于每一个测点的工作表中的数据格式进行判断，其中第一行数据应该为日期类型，而第一列数据应该为表示开挖深度的Double类型。
		/// 如果表头的数据格式不对，则会弹出警告，并将有效的数据限制在未出错前的范围内。
		/// </summary>
		/// <param name="shtMonitorData"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private Dictionary<DateTime, int> getdicDate_ColNum(Worksheet shtMonitorData)
		{
			Dictionary<DateTime, int> dic_Date_ColNum = new Dictionary<DateTime, int>();
			Range UsedRange_shtData = shtMonitorData.UsedRange;
			int rowsCount = UsedRange_shtData.Rows.Count;
			int colsCount = UsedRange_shtData.Columns.Count;
			//
			int endRow = 0;
			int endCol = 0;
			//--------------------------------------------------------------- 找出测斜数据DataRange的范围
			//数据区域（包含x轴的深度数据）的起始单元格为“A3”
			Worksheet with_1 = shtMonitorData;
			int iRowNum = System.Convert.ToInt32(Mnt_Incline.RowNum_FirstData_WithoutDate);
			//表示测点深度的ID列
			//最末尾一行取用rowsCount是基于工作表的第一行有数据的情况。
			Range rgDepth = with_1.Range(with_1.Cells[Mnt_Incline.RowNum_FirstData_WithoutDate, Mnt_Incline.ColNum_Depth],
				with_1.Cells[rowsCount, Mnt_Incline.ColNum_Depth]);
			object[,] vDepth = rgDepth.Value;
			//ReDim F_arrDepth(0 To rowsCount - Mnt_Incline.RowNum_FirstData_WithoutDate)
			
			foreach (object v in vDepth)
			{
				if (v != null)
				{
					try
					{
						//尝试将第一列的深度数据转换为Double
						float Depth = System.Convert.ToSingle(v);
						//如果成功转换，说明此深度数据有效，否则，说明到了最后一行的深度值
						//Me.F_arrDepth(iRowNum - Mnt_Incline.RowNum_FirstData_WithoutDate) = Depth
						iRowNum++;
					}
					catch (Exception)
					{
						MessageBox.Show("第" + Mnt_Incline.ColNum_Depth + "列的深度数据无法转换为数值，请检查第" + iRowNum.ToString() + "行。",
							"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						goto endOfForLoop;
					}
				}
				else
				{
					MessageBox.Show("第" + Mnt_Incline.ColNum_Depth + "列的深度数据为空，请检查第" + iRowNum.ToString() + "行。",
						"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					break;
				}
			}
endOfForLoop:
			endRow = iRowNum - 1;
			//------------------
			int iColNum = System.Convert.ToInt32(Mnt_Incline.ColNum_FirstData_Displacement);
			//表示施工日期的表头行
			//最末尾一列取用colsCount是基于工作表的第一列有数据的情况。
			Range rgDate = with_1.Range(with_1.Cells[Mnt_Incline.RowNumForDate, Mnt_Incline.ColNum_FirstData_Displacement],
				with_1.Cells[Mnt_Incline.RowNumForDate, colsCount]);
			object[,] vDate = rgDate.Value;
			foreach (object v in vDate)
			{
				if (v != null)
				{
					try
					{
						//尝试将第一行的日期数据转换为Date
						DateTime ConstructionDate = System.Convert.ToDateTime(v);
						//创建字典，以监测数据的日期key索引数据所在列号item
						dic_Date_ColNum.Add(ConstructionDate, iColNum);
						//如果成功转换，说明此深度数据有效，否则，说明到了最后一行的深度值
						iColNum++;
					}
					catch (Exception)
					{
						MessageBox.Show("第" + Mnt_Incline.RowNumForDate + "行的日期字段的格式不正确，请检查第" + ConvertColumnNumberToString(iColNum) + "列。",
							"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						goto endOfForLoop1;
					}
				}
				else
				{
					MessageBox.Show("第" + Mnt_Incline.RowNumForDate + "行的日期数据为空，请检查第" + ConvertColumnNumberToString(iColNum) + "列。",
						"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					break;
				}
			}
endOfForLoop1:
			endCol = iColNum - 1;
			//
			UsedRange_shtData = with_1.Range(with_1.Cells[Mnt_Incline.RowNum_FirstData_WithoutDate,
				Mnt_Incline.ColNum_Depth], with_1.Cells[endRow, endCol]);
			//Data_UsedRange包括第一列，但是不包括第一行的日期。
			F_Data_UsedRange = UsedRange_shtData;
			//
			//-------------------------------------
			return dic_Date_ColNum;
		}
		
		/// <summary>
		/// 绘制测斜曲线图
		/// </summary>
		/// <param name="DateOfCurve">绘制的曲线所对应的施工日期</param>
		/// <param name="dataUsedRange">包括第一列，但是不包括第一行的日期。</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private Chart DrawDynamicChart(Worksheet DrawingSheet, DateTime DateOfCurve, Range dataUsedRange)
		{
			string TemplatePath = System.IO.Path.Combine(System.Convert.ToString(My.Settings.Default.Path_Template), Constants.FolderOrFileName.File_Template.Chart_Incline);
			Chart cht = DrawChart(DrawingSheet: ref DrawingSheet, UserDefinedTemplate: true, TemplatePath: ref TemplatePath);
			//------------------------ 设置集合中的第一条曲线的数据
			Series series = cht.SeriesCollection(LowIndexOfObjectsInExcel.SeriesInSeriesCollection);
			
			//设置数据列的数据值
			series.Name = System.Convert.ToString(F_shtMonitorData.Cells[Mnt_Incline.RowNumForDate, Mnt_Incline.ColNum_FirstData_Displacement].Value); //系列名称
			series.XValues = dataUsedRange.Columns[Mnt_Incline.ColNum_FirstData_Displacement].Value; //X轴的数据
			series.Values = dataUsedRange.Columns[Mnt_Incline.ColNum_Depth].value; //Y轴的数据
			
			//初始赋值
			this.F_TheFirstseriesTag = new ClsDrawing_Mnt_Incline.SeriesTag_Incline(series, DateOfCurve);
			
			//测斜数据区的数组，即不包含dataRange中的第1列以外的区域的大数组。
			object arrDataDisplacement = null;
			arrDataDisplacement = F_shtMonitorData.Range(F_shtMonitorData.Cells[Mnt_Incline.RowNum_FirstData_WithoutDate, Mnt_Incline.ColNum_FirstData_Displacement],  //.Value
				F_shtMonitorData.Cells[dataUsedRange.Rows.Count, dataUsedRange.Columns.Count]);
			
			//------------------------ 设置X轴的格式
			dynamic with_3 = cht.Axes(XlAxisType.xlCategory);
			
			//由数据的最小与最大值来划分表格区间
			var imax = F_ExcelAppDrawing.WorksheetFunction.Max(arrDataDisplacement);
			var iMin = F_ExcelAppDrawing.WorksheetFunction.Min(arrDataDisplacement);
			
			//主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
			float unit = float.Parse(Strings.Format((imax - iMin) / Drawing_Incline.AxisMajorUnit_Y, "0.0E+00")); //这里涉及到有效数字的处理的问题
			with_3.MajorUnit = unit;
			with_3.MinorUnitIsAuto = true;
			
			//坐标轴上显示的总区间
			with_3.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(iMin, with_3.MajorUnit);
			with_3.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(imax, with_3.MajorUnit);
			
			//坐标轴标题
			with_3.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_Dynamic, this.F_MonitorType, XlAxisType.xlCategory);
			
			//------------------------- 设置Y轴的格式
			dynamic with_4 = cht.Axes(XlAxisType.xlValue);
			with_4.MajorUnit = Drawing_Incline.AxisMajorUnit_Y;
			with_4.MinorUnitIsAuto = true;
			var arr = dataUsedRange.Columns[1].value;
			var imin = F_ExcelAppDrawing.WorksheetFunction.Min(dataUsedRange.Columns[Mnt_Incline.ColNum_Depth]);
			with_4.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(imin, with_4.MajorUnit);
			//
			var imax = F_ExcelAppDrawing.WorksheetFunction.Max(dataUsedRange.Columns[Mnt_Incline.ColNum_Depth]);
			with_4.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(imax, with_4.MajorUnit);
			//
			with_4.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_Dynamic, this.F_MonitorType, XlAxisType.xlValue);
			return cht;
		}
		
		/// <summary>
		/// 由选择的基坑区域绘制对应的挖深直线与文本框
		/// </summary>
		///  <param name="ProcessRegion">此基坑区域所在的列的Range对象(包括前面几行的表头数据)</param>
		///  <param name="ShowLabelsWhileRolling">指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。</param>
		///  <param name="inclineTopElev"> 测斜管顶部的标高值 </param>
		/// <remarks>此挖深直线与文本框用来进行后期滚动时显示每一天的开挖情况之用的。</remarks>
		private void RollingDepth_lineAndtextbox(clsData_ProcessRegionData ProcessRegion, DateTime Date_theFirstCurve, bool ShowLabelsWhileRolling, float inclineTopElev)
			{
			// ---------------- 绘制表示挖深的直线与文本框
			//'对第一条监测曲线的施工日期进行初始同仁，用来绘制第一条监测曲线
			//Me.F_Date_theFirstCurve = F_dicDate_ColNum.Keys(0)
			float excavElev = 0;
			try
			{
				excavElev = System.Convert.ToSingle(ProcessRegion.Date_Elevation[Date_theFirstCurve]);
			}
			catch (KeyNotFoundException)
			{
				DateTime ClosestDate = ClsData_DataBase.FindTheClosestDateInSortedList(ProcessRegion.Date_Elevation.Keys, Date_theFirstCurve);
				excavElev = System.Convert.ToSingle(ProcessRegion.Date_Elevation[ClosestDate]);
			}
			//
			//根据每一个标高值，画出相应的深度线及文本框
			// ---------------------- 绘制直线 与 设置直线格式 ----------------------------------------------------
			
			//直线的基本几何参数
			var linetop = ExcelFunction.GetPositionInChartByValue(F_myChart.Axes(XlAxisType.xlValue), inclineTopElev - excavElev);
			//
			float lineLeft = 0;
			float lineWidth = 0;
			PlotArea plotA = F_myChart.PlotArea;
			lineLeft = (float) (plotA.InsideLeft + 0.2 * plotA.InsideWidth);
			lineWidth = (float) (0.12 * plotA.InsideWidth); //水平线条的长度
			//
			//绘制直线与文本框
			Shape Line = F_myChart.Shapes.AddLine(BeginX: ref lineLeft, BeginY: ref linetop, EndX: lineLeft + lineWidth, EndY: ref linetop);
			// -- 设置直线格式 --
			Line.Line.ForeColor.RGB = Information.RGB(255, 0, 0);
			Line.Line.Weight = (float) (1.5F);
			Line.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadStealth;
			Line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadLong;
			Line.Line.EndArrowheadWidth = Office.MsoArrowheadWidth.msoArrowheadWidthMedium;
			// ---------------------- 绘制文本框 与 设置文本框格式 ----------------------------------------------------
			Shape Textbox = null;
			if (ShowLabelsWhileRolling)
			{
				
				Textbox = F_myChart.Shapes.AddTextbox(Orientation: ref Office.MsoTextOrientation.msoTextOrientationHorizontal, Left: plotA.InsideLeft, Top: linetop - 10, Width: lineLeft - plotA.InsideLeft + 100, Height: 20);
				
				Textbox.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText;
				Textbox.Left = Line.Left;
				Textbox.Top = Line.Top - Textbox.Height;
				//文本框中的文本
				Textbox.TextFrame2.TextRange.Text = "挖深"; //& vbTab & ClassData_DataBase.cstDepthRefer
				Textbox.TextFrame2.TextRange.Font.Size = 12;
				Textbox.TextFrame2.TextRange.Font.Name = AMEApplication.FontName_TNR;
				Textbox.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;
				Textbox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Information.RGB(255, 0, 0);
			}
			// ----------------   为数据列添加信息
			F_TheFirstseriesTag.DepthLine = Line;
			F_TheFirstseriesTag.DepthTextbox = Textbox;
		}
		
		/// <summary>
		/// 标注每个构件的标高
		/// </summary>
		/// <param name="inclineTopElevation"> 测斜管顶端的绝对标高值 </param>
		/// <param name="ComponentsOfExcationID">记录基坑ID中对应的每一个构件项目与其对应标高的数组,
		/// 数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值</param>
		/// ''' <remarks></remarks>
		private void ComponentsAndElevatins(float inclineTopElevation, Component[] ComponentsOfExcationID) //, ByVal ExcavID As String, ByVal IDtoElevationAndData As Dictionary(Of String, ClsData_DataBase.clsExcavationID)
		{
			//ComponentsOfExcationID
			//Dim arrexcavation(,) As String = IDtoElevationAndData.Item(ExcavID).Components
			//定义标高线的长度范围
			PlotArea plotA;
			
			List<string> lst_Name_LinesAndTextBox = new List<string>();
			plotA = F_myChart.PlotArea;
			
			//开始对此基坑ID对应的[标高项，标高值]中的每一项，用直线标注其位置，用文本框显示其标高值或其他信息
			for (byte i = 0; i <= (ComponentsOfExcationID.Length - 1); i++)
			{
				Component Component = ComponentsOfExcationID[i];
				float Elevation = Component.Elevation;
				//基准标高，即将标高数据换算为相对于地面标高的深度数据
				float relativedepth = inclineTopElevation - Elevation;
				if (relativedepth >= 0)
				{
					if (Component.Type == ComponentType.Strut || Component.Type == ComponentType.TopOfBottomSlab)
					{
						//------------------------------------ 绘制对应标高项的标高线与文本框
						Shape Line = null;
						Shape Textbox = null;
						//-------------------------------
						// 根据每一个标高值，画出相应的深度线及文本框
						
						DrawDepthLineAndTextBox(myChart: ref F_myChart, relativeDepth: ref relativedepth, Line: ref Line, ShowLabelsWhileRolling: true, Textbox: ref Textbox);
						
						//文本框中的文本
						Textbox.TextFrame2.TextRange.Text = Component.Description + " (" + Elevation.ToString() + ")";
						//-------------------------------
						lst_Name_LinesAndTextBox.Add(Line.Name);
						lst_Name_LinesAndTextBox.Add(Textbox.Name);
					}
				}
			} //下一个标高项
			
			try //将这些构件与标高的图形对象组合为一个组。
			{
				//可能会由于数组中的图形小于两个，而出现不能执行Group的错误。
				F_myChart.Shapes.Range(lst_Name_LinesAndTextBox.ToArray).Group();
			}
			catch (Exception ex)
			{
				Debug.Print("将形状进行组合时出错。异常原因为:{0}" + "\r\n" +
					"，报错位置：为 {1}。", ex.Message, ex.TargetSite.Name);
			}
		}
		
		/// <summary>
		/// 根据每一个构件的标高值，画出相应的深度线及文本框
		/// </summary>
		/// <param name="myChart">进行绘图的图表对象</param>
		/// <param name="relativeDepth"> 构件相对于测斜管顶部的深度值</param>
		///  <param name="ShowLabelsWhileRolling">指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。</param>
		///  <param name="Line">要返回的Line对象</param>
		///  <param name="Textbox">要返回的文本框Textbox对象，如果ShowLabelsWhileRolline的值为False，则其返回Nothing</param>
		/// <remarks></remarks>
		private void DrawDepthLineAndTextBox(Chart myChart, float relativeDepth, ref Shape Line, bool ShowLabelsWhileRolling, ref Shape Textbox)
			{
			//直线与文本框的基本几何参数
			
			//
			var linetop = ExcelFunction.GetPositionInChartByValue(myChart.Axes(XlAxisType.xlValue), relativeDepth);
			
			//---------------------------- 绘制直线 设置直线格式 ----------------------------------
			Line = myChart.Shapes.AddLine(BeginX: myChart.ChartArea.Left + myChart.ChartArea.Width, BeginY: ref linetop, EndX: myChart.ChartArea.Left + myChart.ChartArea.Width - 50, EndY: ref linetop); //EndX中的 -50 即为水平直线的长度
			
			// -- 设置直线格式 --
			Line.Line.ForeColor.RGB = Information.RGB(0, 0, 255);
			Line.Line.Weight = (float) (1.5F);
			Line.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadStealth;
			Line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadLong;
			Line.Line.EndArrowheadWidth = Office.MsoArrowheadWidth.msoArrowheadWidthMedium;
			//--------------------------- 绘制文本框 设置文本框格式----------------------------
			if (ShowLabelsWhileRolling)
			{
				Textbox = myChart.Shapes.AddTextbox(Orientation: ref Office.MsoTextOrientation.msoTextOrientationHorizontal, Left: myChart.ChartArea.Left + myChart.ChartArea.Width - 100, Top: linetop - 15, Width: 100, Height: 20); //文本框的宽度为100像素
				// ------ 设置文本框格式
				Textbox.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText;
				//文字位于文本框的右上角
				Textbox.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
				Textbox.TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight;
				Textbox.TextFrame2.MarginRight = 2; //文本距离文本框右边界的长度，单位为像素
				//文本框中的文本
				Textbox.TextFrame2.TextRange.Font.Size = 9;
				Textbox.TextFrame2.TextRange.Font.Name = AMEApplication.FontName_TNR;
				Textbox.TextFrame2.TextRange.Font.Bold = true;
				Textbox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Information.RGB(0, 0, 255);
			}
		}
		
#endregion
		
#region   ---  最值图
		
		/// <summary>
		/// 根据工作表中有效的日期范围以及对应的数据范围，得到每一天中，某测点的位移极值和对应的深度
		/// </summary>
		/// <param name="dic">以日期索引当天数据在工作表中的列号</param>
		/// <param name="UR">包括第一列的深度，但是不包括第一行是施工日期</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private DateMaxMinDepth getMaxMinDepth(Dictionary<DateTime, int> dic, Range UR)
		{
			int count = System.Convert.ToInt32(dic.Count);
			//
			double[] arrDate = new double[count - 1 + 1];
			object[] arrMax = new object[count - 1 + 1];
			object[] arrMin = new object[count - 1 + 1];
			object[] arrDepth_Max = new object[count - 1 + 1];
			object[] arrDepth_Min = new object[count - 1 + 1];
			//
			Worksheet shtData = UR.Worksheet;
			
			float[] arrDepth = ExcelFunction.ConvertRangeDataToVector<Single>(UR.Columns[1]);
			//
			int index = 0;
			DateTime Day_Data = default(DateTime);
			try
			{
				foreach (DateTime tempLoopVar_Day_Data in dic.Keys)
				{
					Day_Data = tempLoopVar_Day_Data;
					object[] arrData = ExcelFunction.ConvertRangeDataToVector<object>(UR.Columns[dic.Item(Day_Data)]);
					//
					object max = shtData.Application.WorksheetFunction.Max(arrData);
					object min = shtData.Application.WorksheetFunction.Min(arrData);
					
					
					//
					int Row_Max = (int) (shtData.Application.WorksheetFunction.Match(max, arrData, 0));
					int Row_Min = (int) (shtData.Application.WorksheetFunction.Match(min, arrData, 0));
					//
					
					object depth_Max = arrDepth[Row_Max - 1];
					object depth_Min = arrDepth[Row_Min - 1];
					// --------------- 赋值 ---------------
					arrDate[index] = Day_Data.ToOADate();
					arrMax[index] = max;
					arrMin[index] = min;
					arrDepth_Max[index] = depth_Max;
					arrDepth_Min[index] = depth_Min;
					// -----------------------------
					index++;
				}
				return new DateMaxMinDepth(arrDate, arrMax, arrMin, arrDepth_Max, arrDepth_Min);
			}
			catch (Exception ex)
			{
				MessageBox.Show("提取监测数据工作表中的位移最值及其对应的深度值出错。" +
					"\r\n" + "出错的日期为：" + Day_Data.ToShortDateString() +
					"\r\n" + ex.Message +
					"\r\n" + "报错位置：" + ex.TargetSite.Name,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
		}
		
		/// <summary>
		/// 绘制图表
		/// </summary>
		/// <param name="DrawingSheet"></param>
		/// <param name="DMMD"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private Chart DrawDMMDChart(Worksheet DrawingSheet, DateMaxMinDepth DMMD)
		{
			string TemplatePath = System.IO.Path.Combine(System.Convert.ToString(My.Settings.Default.Path_Template), Constants.FolderOrFileName.File_Template.Chart_Max_Depth);
			Chart cht = DrawChart(DrawingSheet: ref DrawingSheet, UserDefinedTemplate: true, TemplatePath: ref TemplatePath);
			//下面这一句激活语句非常重要，如果不对ChartObject进行激活，那么下面的Chart.SetElement就会失效，
			//从而导致图表中相应的元素不存在，那么在对元素的格式进行设置的时候就会报错。
			cht.Parent.Activate();
			//
			double[] arrDate = null;
			object[] arrMax = null;
			object[] arrMin = null;
			object[] arrDepth_Max = null;
			object[] arrDepth_Min = null;
			DateMaxMinDepth with_1 = DMMD;
			arrDate = with_1.ConstructionDate;
			arrMax = with_1.Max;
			arrMin = with_1.Min;
			arrDepth_Max = with_1.Depth_Max;
			arrDepth_Min = with_1.Depth_Min;
			//
			//-------------------------------------------------------------------------
			SeriesCollection SC = cht.SeriesCollection();
			//保证图表中有四条数据系列
			if (SC.Count < 4)
			{
				for (var i = 0; i <= 4 - SC.Count - 1; i++)
				{
					SC.NewSeries();
				}
			}
			//-------------------------------------------------------------------------
			cht.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryCategoryAxisShow);
			cht.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementSecondaryValueAxisTitleRotated);
			cht.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisTitleRotated);
			//-------------------------------------------------------------------------
			//四条曲线的数据及所在的Y轴类型
			Microsoft.Office.Interop.Excel.Series with_3 = SC.Item(1);
			with_3.Name = Drawing_Incline_DMMD.SeriesName_Max;
			with_3.XValues = arrDate;
			with_3.Values = arrMax;
			//Dim b = .XValues     'Date类型的数组传递给XValue后，便成了String类型的数组
			with_3.AxisGroup = XlAxisGroup.xlPrimary;
			Microsoft.Office.Interop.Excel.Series with_4 = SC.Item(2);
			with_4.Name = Drawing_Incline_DMMD.SeriesName_Min;
			with_4.XValues = arrDate;
			with_4.Values = arrMin;
			with_4.AxisGroup = XlAxisGroup.xlPrimary;
			Microsoft.Office.Interop.Excel.Series with_5 = SC.Item(3);
			with_5.Name = Drawing_Incline_DMMD.SeriesName_Depth_Max;
			with_5.XValues = arrDate;
			with_5.Values = arrDepth_Max;
			with_5.AxisGroup = XlAxisGroup.xlSecondary;
			Microsoft.Office.Interop.Excel.Series with_6 = SC.Item(4);
			with_6.Name = Drawing_Incline_DMMD.SeriesName_Depth_Min;
			with_6.XValues = arrDate;
			with_6.Values = arrDepth_Min;
			with_6.AxisGroup = XlAxisGroup.xlSecondary;
			//-------------------------------------------------------------------------
			
			//------------------------ 设置X轴的格式：整个日期跨度
			Axis axisX = cht.Axes(XlAxisType.xlCategory);
			axisX.CategoryType = XlCategoryType.xlTimeScale;
			//设置X轴的时间跨度
			double maxScale = System.Convert.ToDouble(Max_Array<double>(arrDate));
			double minScale = System.Convert.ToDouble(Min_Array<double>(arrDate));
			axisX.MinimumScale = minScale;
			axisX.MaximumScale = maxScale;
			//.MaximumScaleIsAuto = True
			//.MinimumScaleIsAuto = True
			
			//设置竖向网格间距
			axisX.MajorUnitIsAuto = true;
			axisX.MinorUnitIsAuto = true;
			
			//设置坐标轴标签的位置
			axisX.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow;
			axisX.TickLabels.NumberFormatLocal = "yy/m/d";
			axisX.TickLabels.Orientation = (Microsoft.Office.Interop.Excel.XlTickLabelOrientation) 0; // XlTickLabelOrientation.xlTickLabelOrientationHorizontal
			//
			axisX.TickMarkSpacing = 10;
			axisX.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, this.F_MonitorType, XlAxisType.xlCategory);
			
			//------------------------- 设置Y主轴的格式：测斜位移的最值
			Axis axisY_Prime = cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
			axisY_Prime.MinorUnitIsAuto = true;
			axisY_Prime.MajorUnitIsAuto = true;
			axisY_Prime.MaximumScaleIsAuto = true;
			axisY_Prime.MinimumScaleIsAuto = true;
			//
			//Dim imax1 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMax)
			//Dim imax2 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMin)
			//.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(Math.Max(imax1, imax2), .MajorUnit)
			//'
			//Dim imin1 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMax)
			//Dim imin2 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMin)
			//.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(Math.Min(imin1, imin2), .MajorUnit)
			
			axisY_Prime.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, this.F_MonitorType,
				XlAxisType.xlValue, XlAxisGroup.xlPrimary);
			
			//------------------------- 设置Y次轴的格式：测斜位移的最值所对应的深度
			Axis ax = cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary);
			ax.MinorUnitIsAuto = true;
			ax.MajorUnitIsAuto = true;
			ax.MaximumScaleIsAuto = true;
			ax.MinimumScaleIsAuto = true;
			//
			//Dim imax1 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMax)
			//Dim imax2 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMin)
			//.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(Math.Max(imax1, imax2), .MajorUnit)
			//'
			//Dim imin1 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMax)
			//Dim imin2 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMin)
			//.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(Math.Min(imin1, imin2), .MajorUnit)
			ax.ReversePlotOrder = true;
			ax.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, this.F_MonitorType,
				XlAxisType.xlValue, XlAxisGroup.xlSecondary);
			
			//-------------------------------------------------------------------------
			DrawingSheet.Cells[1, 1].Activate(); //以免图形显示时Chart中的对象被选中。
			return cht;
		}
		
		/// <summary>
		/// 绘制开挖工况的位置线
		/// </summary>
		/// <param name="Cht"></param>
		/// <param name="WorkingStage"></param>
		/// <remarks></remarks>
		private void DrawWorkingStage(Chart Cht, List<clsData_WorkingStage> WorkingStage)
		{
			Axis AX = Cht.Axes(XlAxisType.xlCategory);
			string[] arrLineName = new string[WorkingStage.Count - 1 + 1];
			string[] arrTextName = new string[WorkingStage.Count - 1 + 1];
			Chart with_1 = Cht;
			try
			{
				int i = 0;
				float Depth = 0;
				foreach (clsData_WorkingStage WS in WorkingStage)
				{
					// -------------------------------------------------------------------------------------------
					Depth = Project_Expo.Elevation_GroundSurface - WS.Elevation;
					float EndY = (float) (ExcelFunction.GetPositionInChartByValue(Cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary), Depth));
					//
					Shape shpLine = with_1.Shapes.AddLine(BeginX: 0, BeginY: Cht.PlotArea.InsideTop, EndX: 0, EndY: ref EndY);
					
					shpLine.Line.Weight = (float) (1.5F);
					shpLine.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth;
					shpLine.Line.EndArrowheadLength = Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLengthMedium;
					shpLine.Line.EndArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWidthMedium;
					shpLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDashDot;
					shpLine.Line.ForeColor.RGB = Information.RGB(0, 0, 0);
					//
					ExcelFunction.setPositionInChart(shpLine, AX, WS.ConstructionDate.ToOADate());
					// -------------------------------------------------------------------------------------------
					float TextWidth = 25;
					float textHeight = 10;
					Shape shpText = default(Shape);
					shpText = Cht.Shapes.AddTextbox(Orientation: ref Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left: shpLine.Left - TextWidth / 2, Top: shpLine.Top - textHeight, Height: ref textHeight, Width: ref TextWidth);
					ExcelFunction.FormatTextbox_Tag(TextFrame: shpText.TextFrame2, Text: ref WS.Description, HorizontalAlignment: ref Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter);
					// -------------------------------------------------------------------------------------------
					arrLineName[i] = shpLine.Name;
					arrTextName[i] = shpText.Name;
					i++;
				}
				try //可能会由于数组中的图形小于两个，而出现不能执行Group的错误。
				{
					Shape shp1 = Cht.Shapes.Range(arrLineName).Group();
					Shape shp2 = Cht.Shapes.Range(arrTextName).Group();
					Cht.Shapes.Range(new[] {shp1.Name, shp2.Name}).Group();
				}
				catch (Exception ex)
				{
					Debug.Print("将形状进行组合时出错。异常原因为:{0}" + "\r\n" + "，报错位置：为 {1}。", ex.Message, ex.TargetSite.Name);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("设置开挖工况位置出现异常。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
					"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}
		
#endregion
		
#region   ---  通用子方法
		
		/// <summary>
		/// 在此方法中，分别获得了进行绘图所需要的ExcelApplicaion程序、进行绘图的工作表，进行绘图所需要的监测数据的Range，
		/// 以及参与绘图的监测数据中，每一个有效的施工日期在数据工作表中的列号。
		/// </summary>
		/// <param name="sheetMonitorData"></param>
		/// <param name="NewExcelApp"></param>
		/// <param name="ExcelForMntDrawing"></param>
		/// <remarks></remarks>
		private void PrePare(Worksheet sheetMonitorData, bool NewExcelApp, Cls_ExcelForMonitorDrawing ExcelForMntDrawing)
			{
			//   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
			F_ExcelAppDrawing = GetApplication(NewExcelApp: ref NewExcelApp, ExcelForMntDrawing: ref ExcelForMntDrawing, MntDrawingExcelApps: F_GlobalApp.MntDrawing_ExcelApps);
			//   ----------------------------------
			F_ExcelAppDrawing.ScreenUpdating = false;
			
			//打开工作簿以画图
			Workbook wkbkDrawing = default(Workbook);
			if (F_ExcelAppDrawing.Workbooks.Count == 0)
			{
				wkbkDrawing = F_ExcelAppDrawing.Workbooks.Add();
			}
			else
			{
				wkbkDrawing = F_ExcelAppDrawing.Workbooks[1]; //总是定义为第一个，因为就只开了一个
			}
			//新开一个工作表以画图
			//wkbkDrawing.Worksheets(wkbkDrawing.Worksheets.Count).Delete()    '删除工作簿中的最后一个工作表。如果是直接在原界面上新开工作表以绘图，那么此时便可以删除上次绘图的工作表，以节省资源。
			F_shtDrawing = wkbkDrawing.Worksheets.Add();
			
			//   ---------------------------------
			//-- 开始画图 ,并返回关键参数：dicDate_ColNum
			F_dicDate_ColNum = getdicDate_ColNum(sheetMonitorData);
			//   ---------------------------------
		}
		
		/// <summary>
		/// 当进行绘图的数据工作簿发生变化时触发
		/// </summary>
		/// <param name="WorkingDataWorkbook">要进行绘图的数据工作簿</param>
		/// <remarks></remarks>
		public void frmDrawing_Mnt_Incline_DataWorkbookChanged(Workbook WorkingDataWorkbook)
		{
			//在列表中显示出监测数据工作簿中的所有工作表
			byte sheetsCount = (byte) WorkingDataWorkbook.Worksheets.Count;
			if (sheetsCount > 0)
			{
				LstbxDisplayAndItem[] arrSheetsName = new LstbxDisplayAndItem[sheetsCount - 1 + 1];
				byte i = (byte) 0;
				foreach (Worksheet sht in WorkingDataWorkbook.Worksheets)
				{
					arrSheetsName[i] = new LstbxDisplayAndItem(sht.Name, sht);
					i++;
				}
				ListBoxWorksheetsName.DisplayMember = LstbxDisplayAndItem.DisplayMember;
				ListBoxWorksheetsName.ValueMember = LstbxDisplayAndItem.ValueMember;
				ListBoxWorksheetsName.DataSource = arrSheetsName;
				//.Items.Clear()
				//.Items.AddRange(arrSheetsName)
				//.SelectedItem = .Items(0)
			}
			//
			btnGenerate.Enabled = true;
		}
		
		/// <summary>
		/// 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
		/// </summary>
		/// <param name="NewExcelApp">按情况看是否要打开新的Application</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private Application GetApplication(bool NewExcelApp, Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> MntDrawingExcelApps, ref Cls_ExcelForMonitorDrawing ExcelForMntDrawing)
			{
			Application app = default(Application);
			if (NewExcelApp) //打开新的Excel程序
			{
				
				app = new Application();
				ExcelForMntDrawing = new Cls_ExcelForMonitorDrawing(app);
				
			}
			else //在原有的Excel程序上作图
			{
				ExcelForMntDrawing = MntDrawingExcelApps.Last.Value;
				ExcelForMntDrawing.ActiveMntDrawingSheet.RemoveFormCollection();
				app = ExcelForMntDrawing.Application;
			}
			return app;
		}
		
		private DateSpan GetDateSpan(Dictionary<DateTime, int> dic)
		{
			DateSpan DtSp = new DateSpan();
			DateTime[] arrTimeRange = new DateTime[F_dicDate_ColNum.Count - 1 + 1];
			F_dicDate_ColNum.Keys.CopyTo(arrTimeRange, 0);
			Array.Sort(arrTimeRange);
			DtSp.StartedDate = arrTimeRange[0];
			DtSp.FinishedDate = arrTimeRange[arrTimeRange.Length - 1];
			return DtSp;
		}
		
		/// <summary>
		/// 在工作表中绘制出一个Chart，但是不设置其格式
		/// </summary>
		/// <param name="DrawingSheet">Chart所在的工作表</param>
		/// <param name="UserDefinedTemplate">是否使用用户自定义的Chart模板</param>
		/// <param name="ChartType">使用Excel自带的模板时的模板类型</param>
		/// <param name="TemplatePath">用户自定义的Chart模板文件的路径</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private Chart DrawChart(Worksheet DrawingSheet, bool UserDefinedTemplate, XlChartType ChartType = XlChartType.xlXYScatterLines, string TemplatePath = null)
			{
			Chart cht = default(Chart);
			DrawingSheet.Activate();
			//---------- 添加图表并选定模板
			if (!UserDefinedTemplate)
			{
				cht = DrawingSheet.Shapes.AddChart(ChartType).Chart;
			}
			else
			{
				cht = DrawingSheet.Shapes.AddChart().Chart;
				cht.ApplyChartTemplate(TemplatePath);
			}
			
			
			//---------- 设置图表尺寸及标题
			//获取图表中的信息文本框
			F_textbox_Info = cht.Shapes[0].TextFrame2; //Chart中的Shapes集合的第一个元素的下标值为0
			//textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText
			return cht;
		}
		
		/// <summary>
		/// 设置测斜曲线图的Tags属性
		/// </summary>
		/// <param name="MntDataSheet">监测数据所在的工作表</param>
		/// <remarks></remarks>
		private MonitorInfo GetChartTags(Worksheet MntDataSheet)
		{
			string MonitorItem = DrawingItem.Mnt_Incline;
			string PointName = MntDataSheet.Name;
			//
			string ExcavationRegion = "";
			Workbook t_wkbk = MntDataSheet.Parent;
			string filepathwithoutextension = System.IO.Path.GetFileNameWithoutExtension(t_wkbk.FullName);
			//如果没有找到"-"，则会返回-1。
			short Ind = (short) (filepathwithoutextension.IndexOf("-"));
			ExcavationRegion = filepathwithoutextension.Substring(Ind + 1, filepathwithoutextension.Length - Ind - 1);
			//
			MonitorInfo Tags = new MonitorInfo(MonitorItem, ExcavationRegion, PointName);
			return Tags;
		}
		
		/// <summary>
		/// 绘图完成后的收尾工作
		/// </summary>
		/// <param name="blnNewExcelApp">如果是新创建的Excel界面，则进行美化，否则保持原样</param>
		/// <remarks></remarks>
		private void DrawingFinished(bool blnNewExcelApp)
		{
			//-------- 界面显示与美化
			if (blnNewExcelApp) //如果是新创建的Excel界面，则进行美化，否则保持原样
			{
				ExcelAppBeauty(F_ExcelAppDrawing);
			}
			ActiveWindowBeauty(F_ExcelAppDrawing.ActiveWindow);
			
			//刷新窗口
			F_ExcelAppDrawing.ScreenUpdating = true;
			//启用主界面的程序滚动按钮
			APPLICATION_MAINFORM.MainForm.MainUI_RollingObjectCreated();
		}
		
		/// <summary>
		/// 程序界面美化
		/// </summary>
		/// <param name="app"></param>
		/// <remarks></remarks>
		private void ExcelAppBeauty(Application app)
		{
			
			Application with_1 = app;
			with_1.DisplayStatusBar = false;
			with_1.DisplayFormulaBar = false;
			with_1.Visible = true;
			with_1.WindowState = XlWindowState.xlNormal;
			with_1.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Ribbon\",false)");
		}
		private void ActiveWindowBeauty(Window win)
		{
			Window with_1 = win;
			with_1.DisplayGridlines = false;
			with_1.DisplayHeadings = false;
			with_1.DisplayWorkbookTabs = false;
			with_1.Zoom = 100;
			with_1.DisplayHorizontalScrollBar = false;
			with_1.DisplayVerticalScrollBar = false;
			with_1.WindowState = XlWindowState.xlMaximized;
		}
		
#endregion
		
#region   ---  界面操作
		
		/// <summary>
		/// 绘制测点位置图
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnDrawMonitorPoints_Click(object sender, EventArgs e)
		{
			GlobalApplication.Application.DrawingPointsInVisio();
		}
		
		/// <summary>
		/// 选择文件对话框：选择监测数据的文件
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnChooseMonitorData_Click(object sender, EventArgs e)
		{
			string FilePath = "";
			APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Title = "选择测斜数据文件";
			APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb";
			APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FilterIndex = 2;
			if (APPLICATION_MAINFORM.MainForm.OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				FilePath = APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FileName;
			}
			else
			{
				return;
			}
			if (FilePath.Length > 0)
			{
				//将监测数据文件在DataBase的Excel程序中打开
				try
				{
					//有可能会出现选择了同样的监测数据文档
					bool fileHasOpened = false;
					foreach (LstbxDisplayAndItem item in this.ComboBoxOpenedWorkbook.Items)
					{
						Workbook wkbk = (Workbook) item.Value;
						if (string.Compare(wkbk.FullName, FilePath, true) == 0)
						{
							this.F_wkbkData = wkbk;
							fileHasOpened = true;
							break;
						}
					}
					// ----------------------------
					if (fileHasOpened)
					{
						MessageBox.Show("选择的工作簿已经打开", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						this.F_wkbkData = GlobalApplication.Application.ExcelApplication_DB.Workbooks.Open(Filename: ref FilePath, UpdateLinks: false, ReadOnly: true);
						LstbxDisplayAndItem lstItem = new LstbxDisplayAndItem(this.F_wkbkData.Name, this.F_wkbkData);
						this.ComboBoxOpenedWorkbook.Items.Add(lstItem);
						this.ComboBoxOpenedWorkbook.SelectedItem = lstItem;
						if (DataWorkbookChangedEvent != null)
							DataWorkbookChangedEvent(this.F_wkbkData);
					}
				}
				catch (Exception)
				{
					Debug.Print("打开新的数据工作簿出错！");
					return;
				}
			}
		}
		
		public void CheckBox1_CheckedChanged(object sender, EventArgs e)
		{
			if (CheckBox1.Checked)
			{
				ComboBox1.Enabled = true;
			}
			else
			{
				ComboBox1.Enabled = false;
			}
		}
		
#region   ---  选择列表框内容时进行赋值
		
		private List<clsData_WorkingStage> F_WorkingStage;
		public void ComboBox_WorkingStage_SelectedIndexChanged(object sender, EventArgs e)
		{
			this.F_WorkingStage = null;
			LstbxDisplayAndItem lstItem = ComboBox_WorkingStage.SelectedItem;
			if (lstItem != null)
			{
				if (!lstItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
				{
					F_WorkingStage = (List<clsData_WorkingStage>) lstItem.Value;
				}
			}
		}
		
		public void ListBoxWorksheetsName_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				this.F_shtMonitorData = (Worksheet) this.ListBoxWorksheetsName.SelectedValue;
			}
			catch (Exception)
			{
				this.F_shtMonitorData = null;
			}
		}
		
		/// <summary>
		/// 所选择的监测数据所对应的基坑ID，以及此基坑ID中的相关信息。
		/// </summary>
		/// <remarks></remarks>
		private Component[] F_Components;
		public void CbBoxExcavID_SelectedIndexChanged(object sender, EventArgs e)
		{
			this.F_Components = null;
			LstbxDisplayAndItem lstItem_ID = ComboBox_ExcavID.SelectedItem;
			if (lstItem_ID != null)
			{
				if (!lstItem_ID.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
				{
					F_Components = (Component[]) lstItem_ID.Value;
				}
			}
		}
		
		/// <summary>
		/// 设置监测数据的类型
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ComboBox_MntType_SelectedValueChanged(object sender, EventArgs e)
		{
			this.F_MonitorType = MntType.Incline;
			//下面的为程序保留方案，以应对程序中用此窗口来绘制其他类型的监测数据的图形。
			//Dim item As LstbxDisplayAndItem = DirectCast(Me.ComboBox_MntType.SelectedItem, LstbxDisplayAndItem)
			//Me.F_MonitorType = DirectCast(item.Value, MntType)
		}
		
		/// <summary>
		/// 由根据测斜数据表选择的基坑区域的标签，来得到对应的数据列的Range对象。
		/// 作用是根据选择确定是否要绘制动态开挖深度的直线与文本框
		/// </summary>
		/// <remarks></remarks>
		private clsData_ProcessRegionData F_ProcessRegionData;
		public void CbBoxExcavRegion_SelectedIndexChanged(object sender, EventArgs e)
		{
			F_ProcessRegionData = null;
			LstbxDisplayAndItem lstItem_PR = ComboBox_ExcavRegion.SelectedItem;
			if (lstItem_PR != null)
			{
				if (!lstItem_PR.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
				{
					F_ProcessRegionData = (clsData_ProcessRegionData) lstItem_PR.Value;
				}
			}
		}
		
		public void RadioButton_Dynamic_CheckedChanged(object sender, EventArgs e)
		{
			if (RadioButton_Dynamic.Checked)
			{
				this.F_blnMax_Depth = false;
				this.Panel_Dynamic.Visible = true;
				this.Panel_Static.Visible = false;
			}
			else
			{
				this.F_blnMax_Depth = true;
				this.Panel_Dynamic.Visible = false;
				this.Panel_Static.Visible = true;
			}
		}
		
		/// <summary>
		/// 选择进行绘图的数据工作簿
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ComboBoxOpenedWorkbook_SelectedIndexChanged(object sender, EventArgs e)
		{
			LstbxDisplayAndItem lst = this.ComboBoxOpenedWorkbook.SelectedItem;
			try
			{
				Workbook Wkbk = (Workbook) lst.Value;
				this.F_wkbkData = Wkbk;
				APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = true;
				APPLICATION_MAINFORM.MainForm.StatusLabel1.Text = Wkbk.FullName;
				if (DataWorkbookChangedEvent != null)
					DataWorkbookChangedEvent(this.F_wkbkData);
			}
			catch (Exception)
			{
				Debug.Print("选择数据工作簿出错");
			}
		}
		
#endregion
		
#region   ---  关联组合列表框中的数据
		
		/// <summary>
		/// 在列表中列出所有基坑数据的ID值
		/// </summary>
		/// <param name="dic_Data"></param>
		/// <remarks></remarks>
		private void RefreshComobox_ExcavationID(Dictionary<string, clsData_ExcavationID> dic_Data)
		{
			if (dic_Data != null)
			{
				try
				{
					var IDKeys = dic_Data.Keys;
					var IDData = dic_Data.Values;
					int IDcount = System.Convert.ToInt32(IDKeys.Count);
					LstbxDisplayAndItem[] IDList = new LstbxDisplayAndItem[IDcount + 1];
					
					IDList[0] = new LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None);
					//
					for (int i = 0; i <= IDcount - 1; i++)
					{
						string id = System.Convert.ToString(IDKeys(i));
						Component[] Components = IDData(i).Components;
						IDList[i + 1] = new LstbxDisplayAndItem(id, Components);
					}
					GeneralMethods.RefreshCombobox(this.ComboBox_ExcavID, IDList);
					//
				}
				catch (Exception)
				{
				}
				
			}
		}
		
		/// <summary>
		/// 在列表中列出所有基坑区域的Tag值
		/// </summary>
		/// <param name="ProcessRange"></param>
		/// <remarks></remarks>
		private void RefreshComobox_ProcessRange(List<clsData_ProcessRegionData> ProcessRange)
		{
			if (ProcessRange != null)
			{
				try
				{
					int TagsCount = System.Convert.ToInt32(ProcessRange.Count);
					LstbxDisplayAndItem[] TagsList = new LstbxDisplayAndItem[TagsCount + 1];
					
					TagsList[0] = new LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None);
					//
					for (int i1 = 0; i1 <= TagsCount - 1; i1++)
					{
						clsData_ProcessRegionData PR = ProcessRange(i1);
						TagsList[i1 + 1] = new LstbxDisplayAndItem(PR.description, PR);
					}
					GeneralMethods.RefreshCombobox(this.ComboBox_ExcavRegion, TagsList);
				}
				catch (Exception)
				{
				}
			}
		}
		
		private void RefreshCombox_WorkingStage(Dictionary<string, List<clsData_WorkingStage>> NewWorkingStage)
		{
			if (NewWorkingStage != null)
			{
				Dictionary<,> with_1 = NewWorkingStage;
				try
				{
					var RegionNames = with_1.Keys;
					var WorkingStages = with_1.Values;
					int TagsCount = System.Convert.ToInt32(with_1.Count);
					LstbxDisplayAndItem[] TagsList = new LstbxDisplayAndItem[TagsCount + 1];
					TagsList[0] = new LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None);
					//
					for (int i1 = 0; i1 <= TagsCount - 1; i1++)
					{
						TagsList[i1 + 1] = new LstbxDisplayAndItem(System.Convert.ToString(RegionNames(i1)), WorkingStages(i1));
					}
					GeneralMethods.RefreshCombobox(this.ComboBox_WorkingStage, TagsList);
				}
				catch (Exception)
				{
					
				}
			}
		}
		
#endregion
		
#region   ---  窗口的激活去取消激活
		
		public void frmDrawing_Mnt_Incline_Activated(object sender, EventArgs e)
		{
			if (this.F_wkbkData != null)
			{
				APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = true;
				APPLICATION_MAINFORM.MainForm.StatusLabel1.Text = this.F_wkbkData.FullName;
			}
			if (this.WindowState == FormWindowState.Maximized)
			{
				this.MdiParent.MaximumSize = this.MaximumSize;
			}
		}
		
		public void frmDrawing_Mnt_Incline_Deactivate(object sender, EventArgs e)
		{
			APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = false;
		}
#endregion
		
#endregion
		
	}
}
