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
	public partial class frmDeriveDataFromExcel
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
				/// 每一种要提取的数据范围的列数，即Range.Columns.Count的值
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
		
#region   ---  Constants
		
		/// <summary>
		/// 记录异常信息的文本的名称
		/// </summary>
		/// <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
		private const string cstErrorInfoFileName = "ErrorInfo.txt";
		
		/// <summary>
		/// 每一列数据的表头信息所在的行，一般为第一行，一般为数据对应的日期
		/// </summary>
		/// <remarks></remarks>
		const byte cstRowNum_ColumnTitle = 1;
		
		/// <summary>
		/// 提取的数据中的第一行在工作表中所要放置的行号，一般为第3行。第一行一般用来放数据对应的日期，第二行一般为预留行。
		/// </summary>
		/// <remarks></remarks>
		const byte cstRowNum_FirstData = 3;
		
		/// <summary>
		/// 提取的数据中的第一列在工作表中所要放置的列号，一般为第2列。第1列用来放数据的说明
		/// </summary>
		/// <remarks></remarks>
		const byte cstColNum_FirstData = 2;
		
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 用于操作数据的Excel程序
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Application F_ExcelApp;
		
		/// <summary>
		/// 保存提取后的数据的工作簿
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Workbook F_WorkBook_ExportedTo;
		//
		/// <summary>
		/// 搜索日期的正则表达式字符串
		/// </summary>
		/// <remarks></remarks>
		private string F_regexPattern;
		/// <summary>
		/// 利用正则表达式提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
		/// </summary>
		/// <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
		private byte[] F_regexComponents = new byte[4];
		//
		/// <summary>
		/// 异常信息的集合
		/// </summary>
		/// <remarks></remarks>
		private List<string> F_ErrorList;
		
		/// <summary>
		/// 定时触发器
		/// </summary>
		/// <remarks></remarks>
		private System.Windows.Forms.Timer ProgressTimer;
		
#endregion
		
#region   ---  窗体的加载与关闭
		
		/// <summary>
		/// 窗体的加载
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void DeriveDataFromExcel_Load(object sender, EventArgs e)
		{
			//
			this.BackgroundWorker1.WorkerReportsProgress = true;
			this.BackgroundWorker1.WorkerSupportsCancellation = true;
			//
			this.txtbxSavePath.Text = Path.Combine(
				System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), 
				"数据提取.xlsx");
		}
		
		/// <summary>
		/// 鼠标移出控件时隐藏
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void frmDeriveDataFromWord_MouseMove(object sender, MouseEventArgs e)
		{
			if (!this.AddFileOrDirectoryFiles1.Bounds.Contains(e.X, e.Y))
			{
				this.AddFileOrDirectoryFiles1.HideLabel();
			}
		}
		
		/// <summary>
		/// 点击ESC时关闭窗口
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void frmDeriveDataFromExcel_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				this.Close();
			}
		}
		
		/// <summary>
		/// 在窗口关闭前，关闭进行数据处理的Excel与Word程序
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void frmDeriveDataFromWord_FormClosing(object sender, FormClosingEventArgs e)
		{
			//关闭Excel程序
			if (this.F_ExcelApp != null)
			{
				foreach (Microsoft.Office.Interop.Excel.Workbook wkbk in this.F_ExcelApp.Workbooks)
				{
					wkbk.Close(SaveChanges: false);
				}
				this.F_ExcelApp.Quit();
				this.F_ExcelApp = null;
			}
		}
		
#endregion
		
		
#region    ---   界面操作
		/// <summary>
		/// 是否要提取文件名中的日期
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ChkboxParseDate_CheckedChanged(object sender, EventArgs e)
		{
			if (ChkboxParseDate.Checked == true)
			{
				btn_DateFormat.Enabled = true;
				Txtbox_DateFormat.Enabled = true;
			}
			else
			{
				btn_DateFormat.Enabled = false;
				Txtbox_DateFormat.Enabled = false;
			}
		}
		/// <summary>
		/// 构造提取日期的正则表达式
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btn_DateFormat_Click(object sender, EventArgs e)
		{
			frmRegexDate f = new frmRegexDate();
			f.ShowDialog(ref this.F_regexPattern, ref this.F_regexComponents);
			this.Txtbox_DateFormat.Text = this.F_regexPattern;
		}
		/// <summary>
		/// 刷新提取日期的正则表达式
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void Txtbox_DateFormat_TextChanged(object sender, EventArgs e)
		{
			this.F_regexPattern = Txtbox_DateFormat.Text;
		}
		
		// 拖拽操作
		public void APPLICATION_MAINFORM_DragDrop(object sender, DragEventArgs e)
		{
			string[] FilePaths = e.Data.GetData(DataFormats.FileDrop);
			// DoSomething with the Files or Directories that are droped in.
			List<string> arrExcelFilePath = new List<string>();
			foreach (string filepath in FilePaths)
			{
				string ext = Path.GetExtension(filepath);
				string str = ".xlsx.xls.xlsb";
				bool blnExtensionMatched = str.Contains(ext);
				if (blnExtensionMatched)
				{
					this.ListBoxWorksheets.Items.Add(filepath);
				}
			}
		}
		public void APPLICATION_MAINFORM_DragEnter(object sender, DragEventArgs e)
		{
			// See if the data includes text.
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				// There is text. Allow copy.
				e.Effect = DragDropEffects.Copy;
			}
			else
			{
				// There is no text. Prohibit drop.
				e.Effect = DragDropEffects.None;
			}
			
		}
		
#endregion
		
#region   ---  获取文件或文件夹路径
		
		//添加单个文件
		/// <summary>
		/// 以选择文件的形式向列表中添加文件
		/// </summary>
		/// <remarks></remarks>
		public void AddFile(object sender, EventArgs e)
		{
			string[] FilePaths = null;
			this.OpenFileDialog1.Title = "选择要进行数据提取的Excel文件";
			this.OpenFileDialog1.Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb";
			this.OpenFileDialog1.FilterIndex = 2;
			this.OpenFileDialog1.Multiselect = true;
			if (this.OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				FilePaths = this.OpenFileDialog1.FileNames;
			}
			else
			{
				return;
			}
			if (FilePaths.Length > 0)
			{
				this.ListBoxWorksheets.Items.AddRange(FilePaths);
			}
		}
		
		//添加文件夹中的所有文件
		/// <summary>
		/// 以选择文件夹的形式向列表中添加文件
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void lbAddDir_Click(object sender, EventArgs e)
		{
			string strPath = "";
			this.FolderBrowserDialog1.ShowNewFolderButton = true;
			this.FolderBrowserDialog1.Description = "添加文件夹中的全部Excel文件";
			if (this.FolderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				strPath = this.FolderBrowserDialog1.SelectedPath;
			}
			else
			{
				return;
			}
			if (strPath.Length > 0)
			{
				string[] files = Directory.GetFiles(strPath);
				foreach (string strFile in files)
				{
					string ext = Path.GetExtension(strFile);
					if (string.Compare(ext, ".xls", true) == 0 || 
						string.Compare(ext, ".xlsx", true) == 0 || 
						string.Compare(ext, ".xlsb", true) == 0)
					{
						this.ListBoxWorksheets.Items.Add(strFile);
					}
				}
			}
		}
		
		//保存数据的工作簿路径
		public void BtnChoosePath_Click(object sender, EventArgs e)
		{
			string FilePath = "";
			this.SaveFileDialog1.Title = "选择将数据保存到的Excel工作簿路径";
			this.SaveFileDialog1.Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb";
			this.SaveFileDialog1.CreatePrompt = false;
			this.SaveFileDialog1.OverwritePrompt = true;
			this.SaveFileDialog1.AddExtension = true;
			this.SaveFileDialog1.DefaultExt = ".xlsx";
			this.SaveFileDialog1.FilterIndex = 2;
			if (this.SaveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				FilePath = this.SaveFileDialog1.FileName;
			}
			else
			{
				return;
			}
			if (FilePath.Length > 0)
			{
				this.txtbxSavePath.Text = FilePath;
			}
		}
		
		//从列表框中移除选择的工作簿
		/// <summary>
		/// 从列表框中移除选择的工作簿
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnRemove_Click(object sender, EventArgs e)
		{
			short count = (short) this.ListBoxWorksheets.SelectedIndices.Count;
			for (short i = count - 1; i >= 0; i--)
			{
				short index = (short) (this.ListBoxWorksheets.SelectedIndices[i]);
				this.ListBoxWorksheets.Items.RemoveAt(index);
			}
		}
		
#endregion
		
#region   --- 数据提取
		
		/// <summary>
		/// 开始输出数据
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnExport_Click(object sender, EventArgs e)
		{
			if (!this.BackgroundWorker1.IsBusy)
			{
				//打开进行数据操作的Excel程序
				if (this.F_ExcelApp == null)
				{
					this.F_ExcelApp = new Microsoft.Office.Interop.Excel.Application();
				}
				//初始化错误列表
				this.F_ErrorList = new List<string>();
				//
				string strWorkBook_ExportedTo = this.txtbxSavePath.Text;
				//打开保存数据的工作簿，并提取其中的所有工作表
				List<string> listSheetNameInWkbk = new List<string>();
				try
				{
					//If Me.F_WorkBook_ExportedTo Is Nothing Then
					if (File.Exists(strWorkBook_ExportedTo))
					{
						F_WorkBook_ExportedTo = this.F_ExcelApp.Workbooks.Open(Filename: ref strWorkBook_ExportedTo, UpdateLinks: false,[ReadOnly]: false);
					}
					else
					{
						F_WorkBook_ExportedTo = this.F_ExcelApp.Workbooks.Add();
						F_WorkBook_ExportedTo.SaveAs(Filename: ref strWorkBook_ExportedTo, FileFormat: ref 
							Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, CreateBackup: 
							false);
					}
					//End If
					//
					dynamic AllSheets = F_WorkBook_ExportedTo.Worksheets;
					int shtCount = System.Convert.ToInt32(AllSheets.Count);
					foreach (Worksheet shtInWorkbook in (IEnumerable) AllSheets)
					{
						listSheetNameInWkbk.Add(shtInWorkbook.Name);
					}
				}
				catch (Exception)
				{
					MessageBox.Show("保存数据的工作簿打开出错，请检查或者关闭此工作簿。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				
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
						ExportToWorksheet.RangeInfoForExport RangeInfo = new ExportToWorksheet.RangeInfoForExport(ExportedSheet, strTestRange, testRange.Columns.Count);
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
				
				// -----------进行数据提取的工作簿对象数组------------------------
				System.Windows.Forms.ListBox.ObjectCollection WkbkItems = this.ListBoxWorksheets.Items;
				int WorkbooksCount = WkbkItems.Count;
				//记录DataGridView控件中所有数据的数组
				string[] arrWkbk = new string[WorkbooksCount - 1 + 1];
				for (int i = 0; i <= WorkbooksCount - 1; i++)
				{
					string WkbkPath = WkbkItems[i].ToString();
					arrWkbk[i] = WkbkPath;
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
				ExportToWorksheet Export = new ExportToWorksheet(F_WorkBook_ExportedTo, arrWkbk, listRangeInfo, blnParseDateFromFilePath);
				this.BackgroundWorker1.RunWorkerAsync(Export);
			}
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
			int Count_Workbooks = this.ListBoxWorksheets.Items.Count;
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
							this.F_ErrorList.Add(strError);
						}
						finally
						{
							this.BackgroundWorker1.ReportProgress((int) ((iWkbk * Count_RangesInOneWkbk + iRow_Sheet_Range + 1) * unit), strWkbkPath + ":" + sheetExportTo.Name + ":" + strRange);
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
					this.BackgroundWorker1.ReportProgress((int) ((iWkbk + 1) * Count_RangesInOneWkbk * unit), strWkbkPath);
				}
			}
		}
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
		/// 由工作簿的路径返回此组数据的表头信息
		/// </summary>
		/// <param name="FilePath">返回表头数据的依据：工作簿的路径</param>
		/// <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
		/// <returns></returns>
		/// <remarks></remarks>
		private string GetColumnTitle(string FilePath, bool ParseDateFromFilePath)
		{
			string filename = Path.GetFileNameWithoutExtension(FilePath);
			string ColumnTitle = filename;
			//尝试从工作簿文件名分解出其中的日期信息
			if (ParseDateFromFilePath)
			{
				try
				{
					Regex rg = new Regex(this.F_regexPattern, RegexOptions.Singleline, new TimeSpan((long) 10000000.0));
					Match m = rg.Match(filename);
					if (m.Success)
					{
						//按“年/月/日”的格式构造日期字符串
						ColumnTitle = m.Groups[this.F_regexComponents[1]].Value + "/" +
							m.Groups[this.F_regexComponents[2]].Value + "/" +
							m.Groups[this.F_regexComponents[3]].Value;
					}
					else
					{
						string strError = "日期转换异常，异常的工作簿为： " + FilePath;
						this.F_ErrorList.Add(strError);
					}
				}
				catch (Exception)
				{
					string strError = "日期转换异常，异常的工作簿为： " + FilePath;
					this.F_ErrorList.Add(strError);
				}
			}
			return ColumnTitle;
		}
		/// <summary>
		/// 正式开始提取数据
		/// </summary>
		/// <param name="shtExtractFrom">要进行数据提取的工作表</param>
		/// <param name="shtExportTo">放置提取的数据的工作表</param>
		/// <param name="strRange">提取的数据区间</param>
		/// <param name="ColNumToBeAdded">要放置的数据Range的第一个列号</param>
		/// <param name="ColumnTitle">每一个Range数据的表头信息</param>
		/// <remarks></remarks>
		private void ExportData(Microsoft.Office.Interop.Excel.Worksheet shtExtractFrom, Microsoft.Office.Interop.Excel.Worksheet shtExportTo, string strRange, int ColNumToBeAdded, string ColumnTitle)
		{
			//下一个要放置的数据Range的列号
			//Dim ColNumToBeAdded As Integer = ColNum_FirstData
			
			//添加表头数据
			shtExportTo.Cells[cstRowNum_ColumnTitle, ColNumToBeAdded].Value = ColumnTitle;
			
			//要提取的数据的范围
			Range rgOut = shtExtractFrom.Range(strRange);
			int ColsCount = rgOut.Columns.Count;
			int RowsCount = rgOut.Rows.Count;
			//要放置数据的区域范围
			Range rgIn = default(Range);
			Microsoft.Office.Interop.Excel.Worksheet with_1 = shtExportTo;
			rgIn = with_1.Range(with_1.Cells[cstRowNum_FirstData, ColNumToBeAdded], with_1.Cells[cstRowNum_FirstData + RowsCount - 1, ColNumToBeAdded + ColsCount - 1]);
			//提取数据
			rgIn.Value = rgOut.Value;
		}
		
		//报告进度
		public void BackgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
		{
			string strHandlePath = System.Convert.ToString(e.UserState.ToString());
			this.lbSheetName.Text = strHandlePath;
			this.ProgressBar1.Value = e.ProgressPercentage;
		}
		
		// 操作完成
		/// <summary>
		/// 操作完成，关闭Excel程序，写入异常信息，并控件进度条的显示
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BackgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			this.lbSheetName.Text = "Done";
			//激活更改提取日期的正则表达式
			this.ChkboxParseDate.Checked = true;
			//列举所有出错项
			if (this.F_WorkBook_ExportedTo != null)
			{
				//输出异常信息
				string ErrorFilePath = Path.Combine(this.F_WorkBook_ExportedTo.Path, cstErrorInfoFileName);
				Thread thd = new Thread(new System.Threading.ThreadStart(this.ReportError));
				thd.Start(new[] {ErrorFilePath, this.F_ErrorList});
				
				// 保存工作簿中的数据
				this.F_WorkBook_ExportedTo.Save();
				
				//关闭或者显示工作簿
				if (this.ChkBxOpenExcelWhileFinished.Checked)
				{
					this.F_ExcelApp.Visible = true;
					this.F_WorkBook_ExportedTo.Worksheets.Item(1).Activate();
				}
				else
				{
					this.F_WorkBook_ExportedTo.Close(SaveChanges: true);
					this.F_WorkBook_ExportedTo = null;
					//关闭Excel程序
					this.F_ExcelApp.Quit();
					this.F_ExcelApp = null;
				}
			}
			//最后刷新进度条
			if (this.ProgressTimer == null)
			{
				this.ProgressTimer = new System.Windows.Forms.Timer();
				this.ProgressTimer.Tick += this.ProgressTimer_Tick;
			}
			this.ProgressTimer.Interval = 500;
			this.ProgressTimer.Start();
		}
		/// <summary>
		/// 将异常信息的集合写入文本中
		/// </summary>
		/// <param name="Parameters">新线程中的输入参数，它是一个有两个元素的数组，
		/// 其中第1个元素代表异常信息文件的路径，第二个代表收集了异常信息的List集合</param>
		/// <remarks></remarks>
		private void ReportError(object Parameters)
		{
			//ByVal ErrorFilePath As String, ByVal ErrorList As List(Of String)
			string ErrorFilePath = System.Convert.ToString(Parameters(0));
			List<string> ErrorList = Parameters(1);
			if (ErrorList.Count > 0)
			{
				StreamWriter sw = new StreamWriter(ErrorFilePath, true);
				//上面这一步会在指定的路径上生成一个新的文件
				sw.WriteLine(DateTime.Now.ToLongDateString() + DateTime.Now.ToLongTimeString());
				foreach (string strError in ErrorList)
				{
					sw.WriteLine(strError);
				}
				//Close之前，文本文件中只没有数据，Close之后，数据被刷新到文本文件中。
				sw.Close();
			}
			
		}
		/// <summary>
		/// 在定时器触发时将进度条的值设置为0
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		private void ProgressTimer_Tick(object sender, EventArgs e)
		{
			this.ProgressBar1.Value = 0;
			this.ProgressTimer.Stop();
			this.ProgressTimer.Dispose();
			this.ProgressTimer = null;
			this.ProgressTimer.Tick += this.ProgressTimer_Tick;
		}
		
#endregion
		
		
	}
}
