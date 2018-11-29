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
using Microsoft.Office.Interop;
using Microsoft.Office.Core;

namespace CableStayedBridge
{
	
	/// <summary>
	/// 绘制剖面标高图的窗口界面
	/// </summary>
	/// <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
	/// 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
	/// 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值，
	/// 其转换关系式为：构件A在绘图中的Y坐标 = (eleTop - 构件A的标高值) * cm2pt + F_sngTopRef</remarks>
	public partial class frmDrawElevation
	{
		
#region   ---  声明与定义
		
#region   ---  Fields
		
#region   ---  与Excel相关的对象
		
		/// <summary>
		/// 进行绘图的Chart对象
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Chart F_DrawingChart;
		
		/// <summary>
		/// 绘图工作表中的文本框，用以记录施工当天的日期
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.TextFrame2 F_Textbox_Info;
		
#endregion
		
		/// <summary>
		/// 程序的主程序对象，这里要单独再用一个变量，是为了解决在多线程调用时APPLICATION_MAINFORM可能会不能被正常调用（它的有些属性会错误地返回Nothing）
		/// </summary>
		/// <remarks></remarks>
		private APPLICATION_MAINFORM F_MainForm; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
		private GlobalApplication GlobalApp; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
		
#endregion
		
#endregion
		
#region   ---  窗体的加载与关闭
		
		public frmDrawElevation()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			ClsData_DataBase.ProcessRangeChanged += this.RefreshComobox_ProcessRange;
		}
		
		//在关闭窗口时将其隐藏
		public void frmDrawSectionalView_FormClosing(object sender, FormClosingEventArgs e)
		{
			//如果是子窗口自己要关闭，则将其隐藏
			//如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
			if (!(e.CloseReason == CloseReason.MdiFormClosing))
			{
				this.Hide();
			}
			e.Cancel = true;
		}
		
		public void frmDrawElevation_Disposed(object sender, EventArgs e)
		{
			ClsData_DataBase.ProcessRangeChanged -= this.RefreshComobox_ProcessRange;
		}
		
#endregion
		
		/// <summary>
		/// 点击生成按钮
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnGenerate_Click(object sender, EventArgs e)
		{
			short count = System.Convert.ToInt16(this.F_SelectedRegions.Count);
			if (count > 0)
			{
				//开始绘图
				if (!this.BGW_Generate.IsBusy)
				{
					//在工作线程中执行绘图操作
					this.BGW_Generate.RunWorkerAsync(this.F_SelectedRegions);
				}
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
			
			//主程序界面的进度条的UI显示
			F_MainForm.ShowProgressBar_Marquee();
			//执行具体的绘图操作
			GenerateChart(e.Argument);
			//在绘图完成后，隐藏进度条
			F_MainForm.HideProgress("Done");
		}
		
		/// <summary>
		/// 当后台的工作者线程结束（即BGW_Generate_DoWork方法执行完毕）时触发，注意，此方法是在UI线程中执行的。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BGW_Generate_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			
		}
		
#endregion
		
		/// <summary>
		/// 开始生成整个绘图图表
		/// </summary>
		/// <param name="SelectedRegion">窗口中所有选择的要进行绘图的基坑区域</param>
		/// <remarks></remarks>
		private void GenerateChart(List<clsData_ProcessRegionData> SelectedRegion)
		{
			//列表中选择的基坑区域
			int count = System.Convert.ToInt32(SelectedRegion.Count);
			if (count > 0)
			{
				//---------------- 打开用于绘图的Excel程序，并进行界面设计
				Microsoft.Office.Interop.Excel.Worksheet DrawingSheet = this.GetDrawingSheet();
				Microsoft.Office.Interop.Excel.Application DrawingApp = DrawingSheet.Application;
				try
				{
					
					//------------------- 在绘图工作表中进行绘图
					this.F_DrawingChart = DrawChart(DrawingSheet, SelectedRegion);
					
					//  ----------- 绘制数据系列图 ---------------------------
					Microsoft.Office.Interop.Excel.SeriesCollection src = this.F_DrawingChart.SeriesCollection();
					Series series_DeepestExca = src.Item(1);
					Series series_Depth = src.Item(2);
					Series[] DataSeries = new Series[2];
					DataSeries = SetDataSeries(this.F_DrawingChart, series_DeepestExca, series_Depth, SelectedRegion);
					
					//-------------------------------------------------------
					DateSpan date_Span = GetDateSpan(SelectedRegion);
					//-------------------------------------------------------
					ClsDrawing_ExcavationElevation shtEle = 
						new ClsDrawing_ExcavationElevation(series_DeepestExca, series_Depth, SelectedRegion, date_Span, 
						this.F_Textbox_Info, DrawingType.Xls_SectionalView);
					
				}
				catch (Exception ex)
				{
					MessageBox.Show("绘制基坑区域开挖标高图失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					//------- Excel的界面美化 --------------------
					ExcelAppBeauty(DrawingApp);
				}
			}
		}
		
#region   ---  绘制标高图
		
		/// <summary>
		/// 打开用于绘图的Excel程序，并进行界面设计
		/// </summary>
		/// <returns></returns>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Worksheet GetDrawingSheet()
		{
			Microsoft.Office.Interop.Excel.Application app = default(Microsoft.Office.Interop.Excel.Application);
			Microsoft.Office.Interop.Excel.Workbook wkbk = default(Microsoft.Office.Interop.Excel.Workbook);
			Microsoft.Office.Interop.Excel.Worksheet sht = default(Microsoft.Office.Interop.Excel.Worksheet);
			//获取绘图的Application对象
			ClsDrawing_ExcavationElevation ElevationDrawing = GlobalApp.ElevationDrawing;
			if (ElevationDrawing == null)
			{
				app = new Microsoft.Office.Interop.Excel.Application();
			}
			else
			{
				app = ElevationDrawing.Application;
			}
			app.Visible = false;
			
			//获取绘图的工作簿
			Microsoft.Office.Interop.Excel.Workbooks wkbks = app.Workbooks;
			if (wkbks.Count == 0)
			{
				wkbk = wkbks.Add();
			}
			else
			{
				wkbk = wkbks.Item(1);
			}
			
			//获取绘图的工作表
			sht = wkbk.Worksheets.Add();
			
			sht.Activate();
			//'绘图的标题
			//Static DrawingNum As Integer = 0
			//DrawingNum += 1
			//app.Caption = "绘图" & DrawingNum.ToString
			//sht.Name = "绘图" & DrawingNum.ToString
			//'
			//
			return sht;
		}
		
		/// <summary>
		/// 开始绘制开挖剖面图的Chart对象
		/// </summary>
		/// <returns>进行绘图的Chart对象的高度值，以磅为单位，可以用来确定Excel Application的高度值</returns>
		/// <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
		/// 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
		/// 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值</remarks>
		private Microsoft.Office.Interop.Excel.Chart DrawChart(Microsoft.Office.Interop.Excel.Worksheet DrawingSheet, 
			List<clsData_ProcessRegionData> SelectedRegion)
		{
			Microsoft.Office.Interop.Excel.Application DrawingApp = DrawingSheet.Application;
			DrawingApp.ScreenUpdating = false;
			DrawingApp.Caption = "开挖标高图";
			
			//---------------- 创建一个新的，进行绘图的Chart对象 -------------------------------
			Excel.Chart DrawingChart = default(Excel.Chart);
			DrawingChart = DrawingSheet.Shapes.AddChart(Top: 0, Left: 0).Chart;
			string TemplatePath = System.IO.Path.Combine(System.Convert.ToString(My.Settings.Default.Path_Template), 
				Constants.FolderOrFileName.File_Template.Chart_Elevation);
			DrawingChart.Parent.Activate();
			DrawingChart.ApplyChartTemplate(TemplatePath);
			this.F_Textbox_Info = DrawingChart.Shapes.Item(1).TextFrame2;
			DrawingChart.ChartTitle.Text = "开挖标高图";
			//
			Microsoft.Office.Interop.Excel.SeriesCollection src = DrawingChart.SeriesCollection();
			for (short i = 0; i <= 1 - src.Count; i++) //确保Chart中至少有两个数据系列
			{
				src.NewSeries();
			}
			// ----------------------- 设置绘图及Excel窗口的尺寸 ----------------------------
			double ChartHeight = 400;
			double InsideLeft = 60;
			double InsideRight = 20;
			double LeastWidth_Chart = 500;
			double LeastWidth_Column = 100;
			//
			double ChartWidth = LeastWidth_Chart;
			double insideWidth = LeastWidth_Chart - InsideLeft - InsideRight;
			//
			ChartWidth = GetChartWidth(DrawingChart, SelectedRegion.Count, 
				LeastWidth_Chart, LeastWidth_Column, 
				InsideLeft + InsideRight, ref insideWidth);
			ChartSize Size_Chart_App = new ChartSize(ChartHeight, 
				ChartWidth, 
				26, 
				9);
			ExcelFunction.SetLocation_Size(Size_Chart_App, DrawingChart, DrawingChart.Application, true);
			//With DrawingChart.PlotArea
			//    .InsideLeft = InsideLeft
			//    .InsideWidth = insideWidth
			//End With
			// --------------------------------------------------
			return DrawingChart;
		}
		
		/// <summary>
		/// 根据数据库中的数据信息，绘制两条数据系列图，并在表示基坑深度的那一条数据系列上绘制此区域所在的基坑ID的构件图
		/// </summary>
		/// <param name="DrawingChart"></param>
		/// <param name="SelectedRegion"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public Series[] SetDataSeries(Chart DrawingChart, Series 
			series_DeepestExca, Series Series_Depth, 
			List<clsData_ProcessRegionData> SelectedRegion)
		{
			int RegionsCount = System.Convert.ToInt32(SelectedRegion.Count);
			string[] arrDescrip = new string[RegionsCount - 1 + 1]; //每一个区域的描述，作为坐标轴中的X轴数据
			float[] arrDeepest = new float[RegionsCount - 1 + 1]; //每一个区域的坑底标高
			float[] arrDepth = new float[RegionsCount - 1 + 1]; //每一个区域在当天的开挖标高，对于初始绘图，先设定这个值为地面标高
			clsData_ExcavationID[] arrExcavID = new clsData_ExcavationID[RegionsCount - 1 + 1]; //每一个区域所对应的基坑ID对象
			float Elevation_Ground = Project_Expo.Elevation_GroundSurface; //项目的自然地面的标高
			//所有选择的区域中的最深的标高位置，以米为单位
			float DeepestElevation = Elevation_Ground;
			for (UInt16 i = 0; i <= RegionsCount - 1; i++)
			{
				clsData_ProcessRegionData Region = SelectedRegion.Item(i);
				arrDescrip[i] = Region.description;
				float BottomElevation = Region.ExcavationID.ExcavationBottom;
				arrDeepest[i] = BottomElevation;
				arrDepth[i] = Elevation_Ground; // ClsData_DataBase.GetElevation(Region.Range, )
				DeepestElevation = Math.Min((short) DeepestElevation, (short) BottomElevation);
				arrExcavID[i] = Region.ExcavationID;
			}
			//  ------------------------  设置Chart数据  ---------------------------
			try
			{
				Series with_2 = series_DeepestExca;
				with_2.Name = "";
				with_2.XValues = arrDescrip;
				with_2.Values = arrDeepest;
				Series with_3 = Series_Depth;
				with_3.Name = "";
				with_3.XValues = arrDescrip;
				with_3.Values = arrDepth;
			}
			catch (Exception ex)
			{
				MessageBox.Show("设置基坑区域开挖图中的开挖标高数据出错！" + "\r\n" + ex.Message +
					"\r\n" + "报错位置：" + ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			
			//  ------------------------  设置坐标轴格式  ---------------------------
			try
			{
				double max = Elevation_Ground;
				
				double min = 0;
				min = System.Convert.ToDouble(Min_Array<Single>(arrDeepest));
				if (min > 0)
				{
					min = Math.Ceiling(min);
				}
				else //注意Math.Ceiling(-3.2)=-3
				{
					min = Math.Floor(min);
				}
				Microsoft.Office.Interop.Excel.Axis axY = DrawingChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
				axY.MaximumScale = max;
				axY.MinimumScale = min;
				axY.AxisTitle.Text = "标高（m）";
				Microsoft.Office.Interop.Excel.Axis axX = DrawingChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
				
			}
			catch (Exception ex)
			{
				MessageBox.Show("设置基坑区域开挖图中的坐标轴格式出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			//  --------------  绘制每一个选定区域所属的基坑ID的支撑位置  -----------
			DrawComponents(arrExcavID, DrawingChart, series_DeepestExca);
			//
			return new[] {series_DeepestExca, Series_Depth};
		}
		
		/// <summary>
		/// 在Chart中绘制每一个选定区域所属的基坑ID的支撑位置
		/// </summary>
		/// <param name="arrExcavID">所有选择的基坑区域所属的基坑ID</param>
		/// <param name="series_DeepestExca">构件图形的位置的参考</param>
		/// <remarks></remarks>
		private void DrawComponents(clsData_ExcavationID[] arrExcavID, Chart DrawingChart, Series series_DeepestExca)
		{
			UInt16 RegionsCount = arrExcavID.Length;
			try
			{
				Chart with_1 = DrawingChart;
				UInt16 Index = 0;
				string[] arrRegionName = new string[RegionsCount - 1 + 1];
				foreach (Point pt in series_DeepestExca.Points())
				{
					Component[] Components = arrExcavID[Index].Components;
					double dblLeft = pt.Left;
					double dblWidth = pt.Width;
					List<string> listComponentName = new List<string>();
					foreach (Component Component in Components)
					{
						if (Component.Type == ComponentType.Strut)
						{
							double dblTop = ExcelFunction.GetPositionInChartByValue(with_1.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue), 
								Component.Elevation);
							//添加标高线
							Microsoft.Office.Interop.Excel.Shape shpLine = with_1.Shapes.AddLine(BeginX: ref dblLeft, BeginY: ref dblTop, EndX: 
								dblLeft + dblWidth, EndY: ref dblTop);
							shpLine.Line.Weight = (float) (1.5F);
							shpLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid;
							shpLine.Line.ForeColor.RGB = Information.RGB(255, 0, 0);
							//添加文本框
							double dblTextHeight = 40;
							Microsoft.Office.Interop.Excel.Shape shpText = with_1.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, (float) 
								dblLeft, (float) (dblTop - dblTextHeight), (float) dblWidth, (float) dblTextHeight);
							ExcelFunction.FormatTextbox_Tag(TextFrame: shpText.TextFrame2, TextSize: 8, Text: ref Component.Description, VerticalAnchor: ref 
								MsoVerticalAnchor.msoAnchorBottom);
							
							Microsoft.Office.Interop.Excel.Shape shpRg_Components = with_1.Shapes.Range(new[] {shpLine.Name, shpText.Name}).Group();
							listComponentName.Add(shpRg_Components.Name);
						}
					} //下一组构件
					UInt16 ComponentCount = listComponentName.Count;
					if (ComponentCount > 2)
					{
						arrRegionName[Index] = DrawingChart.Shapes.Range(listComponentName.ToArray).Group().Name;
					}
					else if (ComponentCount == 1) //只有一组构件，那么不用进行Group
					{
						arrRegionName[Index] = System.Convert.ToString(listComponentName.Item(0));
					}
					else //说明此基坑区域中一个构件也没有。
					{
						arrRegionName[Index] = null;
					}
					Index++;
				} //下一个基坑区域
				if (arrRegionName.Count() >= 2)
				{
					with_1.Shapes.Range(arrRegionName).Group();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("绘制基坑构件出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
		}
		
		/// <summary>
		/// 确定Chart的宽度值，以磅为单位
		/// </summary>
		/// <param name="DrawingChart">进行绘图的Chart</param>
		/// <param name="PointsCount">选择的基坑区域的数量</param>
		/// <param name="LeastWidth_Chart">Chart对象的宽度的最小值，即使图表中只有很少的点</param>
		/// <param name="LeastWidth_Column">柱形图中每一个柱形的最小宽度，用来进行基本的文本书写</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private double GetChartWidth(Microsoft.Office.Interop.Excel.Chart DrawingChart, UInt16 PointsCount, double 
			LeastWidth_Chart, double LeastWidth_Column, double Margin, ref 
			double InsideWidth)
		{
			double ChartWidth = LeastWidth_Chart;
			InsideWidth = LeastWidth_Chart - Margin;
			//
			Microsoft.Office.Interop.Excel.Chart with_1 = DrawingChart;
			Microsoft.Office.Interop.Excel.SeriesCollection src = with_1.SeriesCollection();
			UInt16 n = src.Count;
			double hh = LeastWidth_Chart / PointsCount;
			ChartGroup ChtGroup = with_1.ChartGroups(1);
			// ------------------------------------------------------------------------------------------------
			//1、在已知最小的chart宽度(即最小的PlotArea.InsideWidth)的情况下，验算柱形的宽度是否比指定的最小柱形宽度要大
			double H = InsideWidth;
			// Dim O As Single = ChtGroup.Overlap / 100
			ChtGroup.Overlap = 0;
			float G = (float) ((double) ChtGroup.GapWidth / 100);
			double ColumnWidth = hh / (1 + G + n);
			if (ColumnWidth < LeastWidth_Column)
			{
				// ------------------------------------------------------------------------------------------------
				//2、在已知柱体的最小宽度的情况下，去推算整个PlotArea.InsideWidth的值
				
				
			}
			return ChartWidth;
		}
		
#endregion
		
#region   ---  子方法
		
		/// <summary>
		/// 获取时间跨度
		/// </summary>
		/// <param name="SelectedRegion">窗口中所有选择的要进行绘图的基坑区域</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private DateSpan GetDateSpan(List<clsData_ProcessRegionData> SelectedRegion)
		{
			UInt16 RegionCount = SelectedRegion.Count;
			if (RegionCount > 0)
			{
				//
				const byte cstColNum_DateList = Data_Drawing_Format.DB_Progress.ColNum_DateList;
				const byte cstRowNum_TheFirstDay = Data_Drawing_Format.DB_Progress.RowNum_TheFirstDay;
				//
				Range[] arrRangesChosen = new Range[RegionCount - 1 + 1];
				UInt16 i = 0;
				foreach (clsData_ProcessRegionData Region in SelectedRegion)
				{
					arrRangesChosen[i] = Region.Range_Process;
					i++;
				}
				DateSpan DateSpan_Old = new DateSpan();
				
				DateTime dtStartDay_new = default(DateTime);
				DateTime dtEndDay_new = default(DateTime);
				try
				{
					//以arrRangesChosen的第一个值来进行TimeSpan_SectionalView的初始化
					Worksheet sht = arrRangesChosen[0].Worksheet;
					dtStartDay_new = System.Convert.ToDateTime(sht.Cells[cstRowNum_TheFirstDay, cstColNum_DateList].value);
					dtEndDay_new = System.Convert.ToDateTime(sht.Cells[sht.UsedRange.Rows.Count, cstColNum_DateList].value);
					DateSpan_Old.StartedDate = dtStartDay_new;
					DateSpan_Old.FinishedDate = dtEndDay_new;
					//从第二项开始，对TimeSpan_SectionalView进行扩展
					DateSpan DateSpan_new = new DateSpan();
					for (byte iRg = 1; iRg <= (arrRangesChosen.Length - 1); iRg++)
					{
						sht = arrRangesChosen[iRg].Worksheet;
						dtStartDay_new = System.Convert.ToDateTime(sht.Cells[cstRowNum_TheFirstDay, cstColNum_DateList].value);
						dtEndDay_new = System.Convert.ToDateTime(sht.Cells[sht.UsedRange.Rows.Count, cstColNum_DateList].value);
						DateSpan_new.StartedDate = dtStartDay_new;
						DateSpan_new.FinishedDate = dtEndDay_new;
						DateSpan_Old = GeneralMethods.ExpandDateSpan(DateSpan_Old, DateSpan_new);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("提取基坑区域中的施工日期跨度DateSpan出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				
				return DateSpan_Old;
			}
		}
		
		/// <summary>
		/// 程序界面美化
		/// </summary>
		/// <param name="ExcelApp"></param>
		/// <remarks></remarks>
		private void ExcelAppBeauty(Application ExcelApp)
		{
			Application with_1 = ExcelApp;
			with_1.Visible = false;
			with_1.DisplayStatusBar = false;
			with_1.DisplayFormulaBar = false;
			with_1.ActiveWindow.DisplayGridlines = false;
			with_1.ActiveWindow.DisplayHeadings = false;
			with_1.ActiveWindow.DisplayWorkbookTabs = false;
			with_1.ActiveWindow.Zoom = 100;
			with_1.ActiveWindow.DisplayHorizontalScrollBar = false;
			with_1.ActiveWindow.DisplayVerticalScrollBar = false;
			with_1.WindowState = XlWindowState.xlNormal;
			FixWindow(with_1.Hwnd);
			with_1.Visible = true;
			//隐藏Excel的功能区。注意，在进行隐藏前，一定要确保Application.Visible属性为True，否则显示出来界面会没有“标题栏”。
			//在此情况下，只要用鼠标右击窗口中的任何区域（单元格或形状等）即可将标题栏显示出来。
			with_1.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Ribbon\",false)");
			with_1.ScreenUpdating = true;
		}
		
#endregion
		
#region   ---  一般的界面操作
		
		/// <summary>
		/// 全选
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnChooseAll_Click(object sender, EventArgs e)
		{
			bool blnSelected = true;
			byte btItemsCount = (byte) lstbxChooseRegion.Items.Count;
			for (byte i = 0; i <= btItemsCount - 1; i++)
			{
				lstbxChooseRegion.SetSelected(i, blnSelected);
			}
		}
		
		/// <summary>
		/// 全不选
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnChooseNone_Click(object sender, EventArgs e)
		{
			lstbxChooseRegion.ClearSelected();
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
					// list all the tags of each excavation region in the listbox
					byte count = System.Convert.ToByte(ProcessRange.Count);
					LstbxDisplayAndItem[] arrItems = new LstbxDisplayAndItem[count - 1 + 1];
					byte i = (byte) 0;
					foreach (clsData_ProcessRegionData PR in ProcessRange)
					{
						arrItems[i] = new LstbxDisplayAndItem(DisplayedText: PR.description, Value: 
							PR);
						i++;
					}
					GeneralMethods.RefreshCombobox(this.lstbxChooseRegion, arrItems);
				}
				catch (Exception)
				{
					
				}
			}
		}
		
		/// <summary>
		/// 窗口中所有选择的要进行绘图的基坑区域
		/// </summary>
		/// <remarks></remarks>
		private List<clsData_ProcessRegionData> F_SelectedRegions;
		public void RefreshSelectedRegion(object sender, EventArgs e)
		{
			var items = this.lstbxChooseRegion.SelectedItems;
			List<clsData_ProcessRegionData> SelectedRegion = new List<clsData_ProcessRegionData>();
			foreach (LstbxDisplayAndItem item in items)
			{
				SelectedRegion.Add((clsData_ProcessRegionData) item.Value);
			}
			this.F_SelectedRegions = SelectedRegion;
		}
		
#endregion
		
	}
}
