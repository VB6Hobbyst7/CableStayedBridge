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
	public abstract partial class frmDeriveData
	{
		
		/// <summary>
		/// 此界面所处理的数据类型：Excel中的监测数据还是Word中的监测数据
		/// </summary>
		/// <remarks></remarks>
		protected enum ChildType
		{
			Word,
			Excel
		}
		
#region   ---  Constants
		
		/// <summary>
		/// 记录异常信息的文本的名称
		/// </summary>
		/// <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
		protected const string cstErrorInfoFileName = "ErrorInfo.txt";
		
		/// <summary>
		/// 每一列数据的表头信息所在的行，一般为第一行，一般为数据对应的日期
		/// </summary>
		/// <remarks></remarks>
		protected const byte cstRowNum_ColumnTitle = 1;
		
		/// <summary>
		/// 提取的数据中的第一行在工作表中所要放置的行号，一般为第2行。第一行一般用来放数据对应的日期，第二行一般为预留行。
		/// </summary>
		/// <remarks></remarks>
		protected const byte cstRowNum_FirstData = 2;
		
		/// <summary>
		/// 提取的数据中的第一列在工作表中所要放置的列号，一般为第2列。第1列用来放数据的说明
		/// </summary>
		/// <remarks></remarks>
		protected const byte cstColNum_FirstData = 2;
		
#endregion
		
#region   ---  Fields
		/// <summary>
		/// 此界面所处理的数据类型：Excel中的监测数据还是Word中的监测数据
		/// </summary>
		/// <remarks></remarks>
		protected ChildType F_ChildType;
		
		/// <summary>
		/// 用于操作数据的Excel程序
		/// </summary>
		/// <remarks></remarks>
		protected Microsoft.Office.Interop.Excel.Application F_ExcelApp;
		
		/// <summary>
		/// 保存提取后的数据的工作簿
		/// </summary>
		/// <remarks></remarks>
		protected Microsoft.Office.Interop.Excel.Workbook F_WorkBook_ExportedTo;
		//
		/// <summary>
		/// 搜索日期的正则表达式字符串
		/// </summary>
		/// <remarks></remarks>
		protected string F_regexPattern;
		/// <summary>
		/// 利用正则表达式提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
		/// </summary>
		/// <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
		protected byte[] F_regexComponents = new byte[4];
		//
		/// <summary>
		/// 异常信息的集合
		/// </summary>
		/// <remarks></remarks>
		protected List<string> F_ErrorList;
		
		/// <summary>
		/// 定时触发器
		/// </summary>
		/// <remarks></remarks>
		protected System.Windows.Forms.Timer ProgressTimer;
		
		/// <summary>
		///
		/// </summary>
		/// <remarks></remarks>
		protected List<string> listSheetNameInWkbk;
		
		/// <summary>
		/// 列表框中所记录的所有要进行数据提取的Excel或者Word文档的路径
		/// </summary>
		/// <remarks></remarks>
		protected string[] arrDocPaths;
#endregion
		
#region   ---  窗体的加载与关闭
		
		/// <summary>
		/// 构造函数
		/// </summary>
		/// <remarks></remarks>
		public frmDeriveData()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			//
			this.BackgroundWorker1.WorkerReportsProgress = true;
			this.BackgroundWorker1.WorkerSupportsCancellation = true;
			//
			this.txtbxSavePath.Text = Path.Combine(
				System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), 
				"数据提取.xlsx");
			this.F_ChildType = ChildType.Excel;
			this.F_regexComponents = new[] {1, 2, 3, 4};
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
				if (this.F_ExcelApp.Visible == false)
				{
					foreach (Microsoft.Office.Interop.Excel.Workbook wkbk in this.F_ExcelApp.Workbooks)
					{
						wkbk.Close(SaveChanges: false);
					}
					this.F_ExcelApp.Quit();
					this.F_ExcelApp = null;
				}
			}
		}
		
#endregion
		
#region   ---  界面操作
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
			string[] FilePaths = e.Data.GetData(DataFormats.FileDrop) as string[];
            // DoSomething with the Files or Directories that are droped in.
            List<string> arrExcelFilePath = new List<string>();
			foreach (string filepath in FilePaths)
			{
				string ext = Path.GetExtension(filepath);
				string str = ".xlsx.xls.xlsb";
				if (this.F_ChildType == ChildType.Word)
				{
					str = ".docx.doc.docm";
				}
				bool blnExtensionMatched = str.Contains(ext);
				if (blnExtensionMatched)
				{
					this.ListBoxDocuments.Items.Add(filepath);
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
		//
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
			if (this.F_ChildType == ChildType.Word)
			{
				this.OpenFileDialog1.Title = "选择要进行数据提取的Word文件";
				this.OpenFileDialog1.Filter = "Word文档(*.docx, *.doc, *.docm)|*.docx;*.doc;*.docm";
			}
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
				this.ListBoxDocuments.Items.AddRange(FilePaths);
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
			if (this.F_ChildType == ChildType.Excel)
			{
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
							this.ListBoxDocuments.Items.Add(strFile);
						}
					}
				}
			}
			else
			{
				this.FolderBrowserDialog1.ShowNewFolderButton = true;
				this.FolderBrowserDialog1.Description = "添加文件夹中的全部Word文档";
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
						if (string.Compare(ext, ".doc", true) == 0 || 
							string.Compare(ext, ".docx", true) == 0 || 
							string.Compare(ext, ".doxm", true) == 0)
						{
							this.ListBoxDocuments.Items.Add(strFile);
						}
					}
				}
			}
			
		}
		
		//保存数据的Excel工作簿路径
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
				// Exit Sub
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
			short count = (short) this.ListBoxDocuments.SelectedIndices.Count;
			for (short i = (short)(count - 1); i >= 0; i--)
			{
				short index = (short) (this.ListBoxDocuments.SelectedIndices[i]);
				this.ListBoxDocuments.Items.RemoveAt(index);
			}
		}
		
#endregion
		
#region   --- 数据提取
		
		/// <summary>
		/// 开始输出数据
		/// </summary>
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
				this.F_ExcelApp.DisplayAlerts = false;
				//一般情况下，默认是隐藏的，如果原来是打开的，则手动隐藏
				this.F_ExcelApp.Visible = false;
				//初始化错误列表
				this.F_ErrorList = new List<string>();
				
				//---------- 打开保存数据的工作簿，并提取其中的所有工作表 ----------------
				string strWorkBook_ExportedTo = this.txtbxSavePath.Text;
				try
				{
					//If Me.F_WorkBook_ExportedTo Is Nothing Then
					if (File.Exists(strWorkBook_ExportedTo))
					{
						F_WorkBook_ExportedTo = this.F_ExcelApp.Workbooks.Open(Filename:  strWorkBook_ExportedTo, UpdateLinks: false,ReadOnly: false);
					}
					else
					{
						F_WorkBook_ExportedTo = this.F_ExcelApp.Workbooks.Add();
						F_WorkBook_ExportedTo.SaveAs(Filename:  strWorkBook_ExportedTo, FileFormat:  
							Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, CreateBackup: 
							false);
					}
				}
				catch (Exception)
				{
					MessageBox.Show("保存数据的工作簿打开出错，请检查或者关闭此工作簿。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				listSheetNameInWkbk = new List<string>();
				foreach (Worksheet shtInWorkbook in F_WorkBook_ExportedTo.Worksheets)
				{
					listSheetNameInWkbk.Add(shtInWorkbook.Name);
				}
				// -----------进行数据提取的工作簿对象数组------------------------
				System.Windows.Forms.ListBox.ObjectCollection WkbkItems = this.ListBoxDocuments.Items;
				int WorkbooksCount = WkbkItems.Count;
				arrDocPaths = new string[WorkbooksCount - 1 + 1];
				for (int i = 0; i <= WorkbooksCount - 1; i++)
				{
					arrDocPaths[i] = WkbkItems[i].ToString();
				}
				StartExportData();
			}
		}
		
		/// <summary>
		/// 开始输出数据
		/// </summary>
		/// <remarks></remarks>
		protected abstract void StartExportData();
		
		/// <summary>
		/// 由工作簿的路径返回此组数据的表头信息
		/// </summary>
		/// <param name="FilePath">返回表头数据的依据：工作簿的路径</param>
		/// <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
		/// <returns></returns>
		/// <remarks></remarks>
		protected string GetColumnTitle(string FilePath, bool ParseDateFromFilePath)
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
#endregion
#region   --- 操作进度与对应处理
		
		
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
			string ErrorFilePath = System.Convert.ToString(Parameters[0]);
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
