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

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Text.RegularExpressions;

namespace CableStayedBridge
{
	public partial class frmDeriveData_Excel : frmDeriveData
	{
		
#region   ---  types
		
		/// <summary>
		/// 在用后台线程提取所有的工作表的数据时，进行传递的参数
		/// </summary>
		/// <remarks>此结构中包含了要进行数据提取的所有文档以及工作表和Range信息</remarks>
		private struct ExportToWorksheet
		{
			/// <summary>
			/// 放置提取后的数据的工作簿
			/// </summary>
			/// <remarks></remarks>
			public Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo;
			/// <summary>
			/// 要进行提取的工作簿
			/// </summary>
			/// <remarks></remarks>
			public string[] arrWkbk;
			/// <summary>
			/// 每一个工作簿中要进行提取的工作表，并用来索引此工作表中的Range范围
			/// </summary>
			/// <remarks>集合中的Worksheet对象对应的是保存数据的工作簿中的工作表对象。</remarks>
			public List<RangeInfoForExport> listRangeInfo;
			
			/// <summary>
			/// 是否要分析出提取数据的工作簿中的日期数据
			/// </summary>
			/// <remarks></remarks>
			public bool ParseDateFromFilePath;
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="WorkBook_ExportedTo">放置提取后的数据的工作簿</param>
			/// <param name="arrWkbk">要进行提取的工作簿</param>
			/// <param name="listRangeInfo">每一个工作簿中要进行提取的工作表，并用来索引此工作表中的Range范围</param>
			/// <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
			/// <remarks></remarks>
			public ExportToWorksheet(Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo, 
				string[] arrWkbk, 
				List<RangeInfoForExport> listRangeInfo, 
				bool ParseDateFromFilePath)
			{
				this.WorkBook_ExportedTo = WorkBook_ExportedTo;
				this.arrWkbk = arrWkbk;
				this.listRangeInfo = listRangeInfo;
				this.ParseDateFromFilePath = ParseDateFromFilePath;
			}
			
			/// <summary>
			/// 每一个工作簿中要提取的Range对象的信息
			/// </summary>
			/// <remarks></remarks>
			public struct RangeInfoForExport
			{
				/// <summary>
				/// 要保存到的工作表对象，也是每一个数据工作簿中要进行检索的工作表对象
				/// </summary>
				/// <remarks></remarks>
				public Worksheet sheet;
				/// <summary>
				/// 工作表中进行提取的数据范围
				/// </summary>
				/// <remarks></remarks>
				public string strRange;
				/// <summary>
				/// 每一种要提取的数据范围的列数，即Range.Areas集合中每一个小Area中的Columns.Count之和
				/// </summary>
				/// <remarks></remarks>
				public int ColumnsCount;
				
				/// <summary>
				/// 构造函数
				/// </summary>
				/// <param name="sheet">要保存到的工作表对象，也是每一个数据工作簿中要进行检索的工作表对象</param>
				/// <param name="strRange">工作表中进行提取的数据范围</param>
				/// <param name="ColumnsCount">每一种要提取的数据范围的列数，即Range.Columns.Count的值</param>
				/// <remarks></remarks>
				public RangeInfoForExport(Worksheet sheet, string strRange, int ColumnsCount)
				{
					this.sheet = sheet;
					this.strRange = strRange;
					this.ColumnsCount = ColumnsCount;
				}
			}
			
		}
		
#endregion
		
		public frmDeriveData_Excel()
		{
			InitializeComponent();
			InitializeComponent_ActivateAtRuntime();
			// Add any initialization after the InitializeComponent() call.
			this.F_ChildType = frmDeriveData.ChildType.Excel;
		}
		
		/// <summary>
		/// 开始输出数据
		/// </summary>
		/// <remarks></remarks>
		protected override void StartExportData()
		{
			
			
			// ------------- 提取每一个工作表与Range范围的格式 -------------
			List<ExportToWorksheet.RangeInfoForExport> listRangeInfo = new List<ExportToWorksheet.RangeInfoForExport>();
			//
			string strTestRange = "";
			//
			//记录DataGridView控件中所有数据的数组
			try
			{
				int RowsCount = MyDataGridView1.Rows.Count;
				for (int RowIndex = 0; RowIndex <= RowsCount - 2; RowIndex++)
				{
					DataGridViewRow RowObject = MyDataGridView1.Rows[RowIndex];
					
					//获取对应的Worksheet对象
					string strSheetName = RowObject.Cells[0].Value.ToString();
					Microsoft.Office.Interop.Excel.Worksheet ExportedSheet = GetExactWorksheet(F_WorkBook_ExportedTo, listSheetNameInWkbk, strSheetName);
					
					//检查Range对象的格式是否正确()
					strTestRange = RowObject.Cells[1].Value.ToString();
					Range testRange = ExportedSheet.Range(strTestRange); //这一步可能出错：Range的格式不规范
					//
					int columnsCount = 0;
					foreach (Range a in testRange.Areas)
					{
						//如果想引用相交区域（公共区域)，可以在多个区域间添加空格“ ”：  如Range("B1:B10 A4:D6 ").Select()  '选中多个单元格区域的交集
						columnsCount += a.Columns.Count;
					}
					ExportToWorksheet.RangeInfoForExport RangeInfo = new ExportToWorksheet.RangeInfoForExport(ExportedSheet, strTestRange, columnsCount);
					listRangeInfo.Add(RangeInfo);
				}
			}
			catch (Exception)
			{
				MessageBox.Show("定义区域范围的格式出错，出错的格式为 : " + "\r\n" 
					+ strTestRange + "，请重新输入", "Error", 
					MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
				return;
			}
			
			
			
			// -----------------------------------
			//是否要分析出提取数据的工作簿中的日期数据
			bool blnParseDateFromFilePath = false;
			if (this.ChkboxParseDate.Checked)
			{
				blnParseDateFromFilePath = true;
			}
			//不允许再更改提取日期的正则表达式
			this.ChkboxParseDate.Checked = false;
			//开始提取数据
			ExportToWorksheet Export = new ExportToWorksheet(F_WorkBook_ExportedTo, arrDocPaths, listRangeInfo, blnParseDateFromFilePath);
			this.BackgroundWorker1.RunWorkerAsync(Export);
			
		}
		
		//在后台线程中执行操作
		public void BackgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			
			//定义初始变量
			ExportToWorksheet ExportToWorksheet = (ExportToWorksheet) e.Argument;
			string[] arrWkbk = ExportToWorksheet.arrWkbk;
			List<ExportToWorksheet.RangeInfoForExport> listRangeInfo = ExportToWorksheet.listRangeInfo;
			Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo = ExportToWorksheet.WorkBook_ExportedTo;
			bool blnParseDateFromFilePath = ExportToWorksheet.ParseDateFromFilePath;
			
			//一共要处理的工作表数(工作簿个数*每个工作簿中提取的工作表数)，用来显示进度条的长度
			int Count_Workbooks = this.ListBoxDocuments.Items.Count;
			int Count_RangesInOneWkbk = System.Convert.ToInt32(listRangeInfo.Count);
			int Count_AllRanges = Count_Workbooks * Count_RangesInOneWkbk;
			
			//
			int percent = 0;
			float unit = 0;
			unit = (float) ((double) (this.ProgressBar1.Maximum - this.ProgressBar1.Minimum) / Count_AllRanges);
			//报告进度
			this.BackgroundWorker1.ReportProgress(percent, "");
			//开始提取数据
			bool blnRangeFormatValidated = false;
			
			for (short iWkbk = 0; iWkbk <= Count_Workbooks - 1; iWkbk++)
			{
				string strWkbkPath = arrWkbk[iWkbk];
				Workbook wkbk = null;
				try
				{
					//下面有可能会出现工作簿打开出错
					wkbk = this.F_ExcelApp.Workbooks.Open(Filename: ref strWkbkPath, UpdateLinks: false,[ReadOnly]: true);
					
					//获取工作簿中的所有工作表的名称，以备后面的与要进行提取的工作表名称的比较之用
					string[] arrExistedSheetsName = new string[wkbk.Worksheets.Count - 1 + 1];
					short i = (short) 0;
					foreach (Worksheet sht in wkbk.Worksheets)
					{
						arrExistedSheetsName[i] = sht.Name;
						i++;
					}
					
					//此工作簿中的每一个要提取的Range对象在列表中的行号。
					int iRow_Sheet_Range = 0;
					//此工作簿所对应的表头的数据：工作簿的名称或者是工作簿中包含的日期信息
					string ColumnTitle = GetColumnTitle(strWkbkPath, blnParseDateFromFilePath);
					//
					for (iRow_Sheet_Range = 0; iRow_Sheet_Range <= Count_RangesInOneWkbk - 1; iRow_Sheet_Range++)
					{
						Worksheet sheetExportTo = listRangeInfo.Item(iRow_Sheet_Range).sheet;
						string strRange = "";
						try
						{
							//有可能会出现工作表提取出错：此工作表不存在
							Worksheet sheetExtractFrom = GetContainedWorksheet(wkbk, arrExistedSheetsName, sheetExportTo.Name);
							if (sheetExtractFrom != null)
							{
								//
								strRange = System.Convert.ToString(listRangeInfo.Item(iRow_Sheet_Range).strRange);
								
								//----------------------------------------
								
								//更新这一组数据所放置的列号,初始的放置数据的列号
								int ColumnsCount = System.Convert.ToInt32(listRangeInfo.Item(iRow_Sheet_Range).ColumnsCount);
								int ColNumToBeAdded = cstColNum_FirstData + ColumnsCount * iWkbk;
								
								// 提取数据
								ExportData(sheetExtractFrom, sheetExportTo, strRange, ColNumToBeAdded, ColumnTitle);
								
								//----------------------------------------
							}
							else
							{
								throw (new NullReferenceException());
							}
						}
						catch (Exception)
						{
							//工作表提取出错：此工作表不存在
							string strError = "工作表：" + wkbk.FullName + " ： " + sheetExportTo.Name + " 无法找到。";
							Debug.Print(strError);
							this.F_ErrorList.Add(strError);
						}
						finally
						{
							this.BackgroundWorker1.ReportProgress(System.Convert.ToInt32((iWkbk * Count_RangesInOneWkbk + iRow_Sheet_Range + 1) * unit), strWkbkPath + ":" + sheetExportTo.Name + ":" + strRange);
						}
					}
				}
				catch (Exception)
				{
					//工作簿打开出错
					string strError = "工作簿：" + wkbk.FullName + " 打开时出错。";
					this.F_ErrorList.Add(strError);
				}
				finally
				{
					if (wkbk != null) //说明工作簿顺利打开
					{
						wkbk.Close(SaveChanges: false);
					}
					this.BackgroundWorker1.ReportProgress(System.Convert.ToInt32((iWkbk + 1) * Count_RangesInOneWkbk * unit), strWkbkPath);
				}
			}
		}
		
#region   --- 数据提取
		
		//匹配工作表
		/// <summary>
		/// 获取工作簿中的工作表对象
		/// </summary>
		/// <param name="wkbk">工作表所在的工作簿</param>
		/// <param name="SheetName">工作表的名称</param>
		/// <remarks>如果此工作表已经在工作簿中出现，则返回对应的工作表，否则，创建一个新的工作表，
		/// 并将新工作表名称添加到已经存在的工作表名称列表中</remarks>
		private Worksheet GetExactWorksheet(Workbook wkbk, List<string> ExistedSheetsName, string SheetName)
		{
			bool blnSheetExisted = false;
			Worksheet sheet = null;
			foreach (string ExistedSheet in ExistedSheetsName)
			{
				//下面的比较一定要忽略大小写，因为Excel中大小写不同的工作表名称被认为是同一个工作表
				//如果新添加的工作表的名称与已经存在的工作表名称只是大小写不同，则会报错。
				if (string.Compare(ExistedSheet, SheetName, true) == 0)
				{
					sheet = wkbk.Worksheets[SheetName];
					//如果检索的工作表名称与已有的工作表名称只是大小写不同，则要将工作表名称设置为进行检索的工作表名称。
					sheet.Name = SheetName;
					return sheet;
				}
			}
			if (!blnSheetExisted)
			{
				sheet = wkbk.Worksheets.Add();
				sheet.Name = SheetName;
				//将新工作表名称添加到已经存在的工作表名称列表中，以供下次调用
				ExistedSheetsName.Add(SheetName);
			}
			return sheet;
		}
		/// <summary>
		/// 获取工作簿中的工作表对象，如果此工作表已经在工作簿中出现，则返回对应的工作表，否则，返回Nothing。
		/// </summary>
		/// <param name="wkbk">工作表所在的工作簿</param>
		/// <param name="ExistedSheetsName">工作簿中已经存在的工作表的名称的集合</param>
		/// <param name="SheetName">工作表的名称</param>
		/// <remarks>比较的依据：1、忽略大小写，2、要检索的工作表的名称的字符串是包含于已经存在的工作表名称的字符串的。</remarks>
		private Worksheet GetContainedWorksheet(Workbook wkbk, string[] ExistedSheetsName, string SheetName)
		{
			Worksheet sheet = null;
			foreach (string ExistedSheet in ExistedSheetsName)
			{
				//忽略大小写，因为Excel中大小写不同的工作表名称被认为是同一个工作表 StringComparer.OrdinalIgnoreCase
				if (ExistedSheet.IndexOf(SheetName, System.StringComparison.OrdinalIgnoreCase) >= 0)
				{
					sheet = wkbk.Worksheets[ExistedSheet];
					return sheet;
				}
			}
			return sheet;
		}
		//
		/// <summary>
		/// 正式开始提取数据
		/// </summary>
		/// <param name="shtExtractFrom">要进行数据提取的工作表</param>
		/// <param name="shtExportTo">放置提取的数据的工作表</param>
		/// <param name="strRange">提取的数据区间</param>
		/// <param name="ColNumToBeAdded">要放置的数据Range的第一个列号</param>
		/// <param name="ColumnTitle">每一个Range数据的表头信息</param>
		/// <remarks>在此方法中，可以引用多个不连续的区域，即在各区域间添加逗号“,”。</remarks>
		private void ExportData(Microsoft.Office.Interop.Excel.Worksheet shtExtractFrom, Microsoft.Office.Interop.Excel.Worksheet shtExportTo, string strRange, int ColNumToBeAdded, string ColumnTitle)
		{
			//下一个要放置的数据Range的列号
			//Dim ColNumToBeAdded As Integer = ColNum_FirstData
			
			//添加表头数据
			shtExportTo.Cells[cstRowNum_ColumnTitle, ColNumToBeAdded].Value = ColumnTitle;
			
			//要提取的数据的范围
			Range rgOut = shtExtractFrom.Range(strRange);
			foreach (Range rg in rgOut.Areas)
			{
				int ColsCount = rg.Columns.Count;
				int RowsCount = rg.Rows.Count;
				//要放置数据的区域范围
				Range rgIn = default(Range);
				Microsoft.Office.Interop.Excel.Worksheet with_1 = shtExportTo;
				//这里将每一个Area的第一个单元格都移到指定的第一个数据单元格，其实是有一点问题的。如果在一个Worksheet中，引用了两个不连续的区域，
				//而这两个区域中的最顶部的单元格并不是在同一行，那么在导出到两列数据时，这两列数据也应该不是从同一行开始的。
				rgIn = with_1.Range(with_1.Cells[cstRowNum_FirstData, ColNumToBeAdded], with_1.Cells[cstRowNum_FirstData + RowsCount - 1, ColNumToBeAdded + ColsCount - 1]);
				//提取数据
				rgIn.Value = rg.Value;
				ColNumToBeAdded += ColsCount;
			}
		}
#endregion
		
	}
}
