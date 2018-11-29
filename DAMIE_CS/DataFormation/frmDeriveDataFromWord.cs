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
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Text.RegularExpressions;

namespace CableStayedBridge
{
	
	public partial class frmDeriveDataFromWord
	{
		
#region   ---  types
		
		/// <summary>
		/// 每一个文档中要提取的测点及数据的位置信息
		/// </summary>
		/// <remarks></remarks>
		private class PointsInfoForExport
		{
			
			/// <summary>
			/// 文档中要进行提取的测点标签
			/// </summary>
			/// <remarks></remarks>
			public string PointTag;
			
			/// <summary>
			/// 进行搜索的方向：按行或者按列，即下一个搜索单元格是按行还是按列的方向前进。
			/// </summary>
			/// <remarks></remarks>
			public Microsoft.Office.Interop.Excel.XlSearchOrder SearchOrder;
			
			/// <summary>
			/// 在Word文档中，测点所对应的数据距离测点单元格的水平偏移位置。
			/// 如数据单元格是在测点标签单元格的左边且紧靠标签单元格，则Offset的值为+1。
			/// </summary>
			/// <remarks></remarks>
			public byte Offset;
			
			
			/// <summary>
			/// 要在excel最终保存数据的工作表中写入数据的列号，即此列前面的行都已经被写入或者是预留的空行。
			/// </summary>
			/// <remarks></remarks>
			public int ColNumToBeWritten;
			/// <summary>
			/// 要在excel最终保存数据的工作表中写入数据的行号，即此行上面的行都已经被写入或者是预留的空行。
			/// </summary>
			/// <remarks></remarks>
			public int RowNumToBeWritten;
			
			/// <summary>
			/// 每一组提取数据所占据的列数。从数据提取上来看，此字段并没有什么作用，因为一般情况下，它的值都应该是2。
			/// 但是从表格的设计上来看，它的值可以用来腾出空的列以放置其他数据。
			/// </summary>
			/// <remarks></remarks>
			public byte ColumnsCountToBeAdd;
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="PointTag">文档中要进行提取的测点标签</param>
			/// <param name="Offset">测点所对应的数据距离测点单元格的水平偏移位置。
			/// 如数据单元格是在测点标签单元格的左边且紧靠标签单元格，则Offset的值为+1。</param>
			/// <param name="SearchOrder">进行搜索的方向：按行或者按列，即下一个搜索单元格是按行还是按列的方向前进。</param>
			/// <remarks></remarks>
			public PointsInfoForExport(string PointTag, int Offset, Microsoft.Office.Interop.Excel.XlSearchOrder SearchOrder)
			{
				this.SearchOrder = SearchOrder;
				this.PointTag = PointTag;
				this.Offset = (byte) Offset;
				//
				this.ColNumToBeWritten = cstColNum_FirstData;
				this.RowNumToBeWritten = cstRowNum_FirstData;
				this.ColumnsCountToBeAdd = cstColumnsCountToBeAdded;
			}
		}
		
		/// <summary>
		/// 在用后台线程提取所有的工作表的数据时，进行传递的参数
		/// </summary>
		/// <remarks>此结构中包含了要进行数据提取的所有文档，
		/// 以及每个文档中进行提取的测点和对应数据的位置标签信息。</remarks>
		private struct ExportToWorksheet
		{
			
			/// <summary>
			/// 放置提取后的数据的工作簿
			/// </summary>
			/// <remarks></remarks>
			public Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo;
			
			/// <summary>
			/// 要进行提取的Word文档
			/// </summary>
			/// <remarks></remarks>
			public string[] arrDocsPath;
			
			/// <summary>
			/// 是否要分析出提取数据的工作簿中的日期数据
			/// </summary>
			/// <remarks></remarks>
			public bool ParseDateFromFilePath;
			
			/// <summary>
			/// 用来暂时保存数据的Excel工作表对象。在提取每一个文档的数据时，
			/// 先将文档中的表格复制到Excel中的此暂存工作表中，然后对于此工作表中的内容进行搜索。
			/// </summary>
			/// <remarks></remarks>
			public Microsoft.Office.Interop.Excel.Worksheet BufferSheet;
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="WorkBook_ExportedTo">放置提取后的数据的工作簿</param>
			/// <param name="arrDocsPath">要进行提取的所有word文档的绝对路径</param>
			/// <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
			/// <remarks></remarks>
			public ExportToWorksheet(Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo, 
				string[] arrDocsPath, 
				bool ParseDateFromFilePath, 
				Microsoft.Office.Interop.Excel.Worksheet BufferSheet)
			{
				this.WorkBook_ExportedTo = WorkBook_ExportedTo;
				this.arrDocsPath = arrDocsPath;
				this.ParseDateFromFilePath = ParseDateFromFilePath;
				this.BufferSheet = BufferSheet;
			}
			
		}
		
#endregion
		
#region   ---  Constants
		
		/// <summary>
		/// 记录异常信息的文本的名称
		/// </summary>
		/// <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
		const string cstErrorInfoFileName = "ErrorInfo.txt";
		
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
		
		/// <summary>
		/// 每一组提取数据所占据的列数。从数据提取上来看，此字段并没有什么作用，因为一般情况下，它的值都应该是2。
		/// 但是从表格的设计上来看，它的值可以用来腾出空的列以放置其他数据。
		/// </summary>
		/// <remarks></remarks>
		const byte cstColumnsCountToBeAdded = 2;
		
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 保存数据的Excel程序
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Application F_ExcelApp;
		/// <summary>
		/// 从Word文档中提取数据
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Word.Application F_WordApp;
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
		/// 定时触发器，用来将进度滚动条的值设置为0
		/// </summary>
		/// <remarks></remarks>
		private System.Windows.Forms.Timer ProgressTimer;
		
		/// <summary>
		/// 用来暂时保存数据的Excel工作表对象。在提取每一个文档的数据时，
		/// 先将文档中的表格复制到Excel中的此暂存工作表中，然后对于此工作表中的内容进行搜索。
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Worksheet F_BufferSheet;
		
		/// <summary>
		/// 每一个文档中要进行提取的测点标签，和与之对应的数据的相对偏移位置。
		/// </summary>
		/// <remarks>集合中的Worksheet对象对应的是保存数据的工作簿中的工作表对象。</remarks>
		private Dictionary<Worksheet, PointsInfoForExport> F_DicPointsInfo;
#endregion
		
#region   ---  窗体的加载与关闭
		
		/// <summary>
		/// 窗体的加载
		/// </summary>
		/// <remarks></remarks>
		public void DeriveDataFromWord_Load(object sender, EventArgs e)
		{
			//设置后台线程的属性
			this.BkgWk_Extract.WorkerReportsProgress = true;
			this.BkgWk_Extract.WorkerSupportsCancellation = true;
			//
			this.txtbxSavePath.Text = Path.Combine(
				System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), 
				"数据提取.xlsx");
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
				this.F_ExcelApp.WorkbookBeforeClose += this.F_ExcelApp_WorkbookBeforeClose;
			}
			//关闭Word程序
			if (this.F_WordApp != null)
			{
				foreach (Microsoft.Office.Interop.Word.Document doc in this.F_WordApp.Documents)
				{
					object null_object = null;
					object null_object2 = null;
					object null_object3 = null;
					doc.Close(ref null_object, ref null_object2, ref null_object3);
				}
				object null_object4 = null;
				object null_object5 = null;
				object null_object6 = null;
				this.F_WordApp.Quit(ref null_object4, ref null_object5, ref null_object6);
				this.F_WordApp = null;
				this.F_WordApp.DocumentBeforeClose += this.F_WordApp_DocumentBeforeClose;
			}
		}
		
		/// <summary>
		/// 逻辑值，指示此时是否正在进行数据的提取操作。
		/// 这是为了应对在程序数据提取时引发的word文档关闭与用户手动关闭Word文档时的区别对待。
		/// </summary>
		/// <remarks></remarks>
		private bool blnIsBeingExtracting = false;
		private void F_ExcelApp_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
		{
			if (!blnIsBeingExtracting)
			{
				Wb.Application.Quit();
				this.F_ExcelApp = null;
				this.F_ExcelApp.WorkbookBeforeClose += this.F_ExcelApp_WorkbookBeforeClose;
			}
		}
		private void F_WordApp_DocumentBeforeClose(Document Doc, ref bool Cancel)
		{
			if (!blnIsBeingExtracting)
			{
				object null_object = null;
				object null_object2 = null;
				object null_object3 = null;
				Doc.Application.Quit(ref null_object, ref null_object2, ref null_object3);
				this.F_WordApp = null;
				this.F_WordApp.DocumentBeforeClose += this.F_WordApp_DocumentBeforeClose;
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
		//
		// 拖拽操作
		public void APPLICATION_MAINFORM_DragDrop(object sender, DragEventArgs e)
		{
			string[] FilePaths = e.Data.GetData(DataFormats.FileDrop) as string[];
			// DoSomething with the Files or Directories that are droped in.
			List<string> arrExcelFilePath = new List<string>();
			foreach (string filepath in FilePaths)
			{
				string ext = Path.GetExtension(filepath);
				string str = ".docx.doc.docm";
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
			this.OpenFileDialog1.Title = "选择要进行数据提取的Word文件";
			this.OpenFileDialog1.Filter = "Word文档(*.docx, *.doc, *.docm)|*.docx;*.doc;*.docm";
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
			short count = (short) this.ListBoxDocuments.SelectedIndices.Count;
			for (short i = (short)(count - 1); i >= 0; i--)
			{
				short index = (short) (this.ListBoxDocuments.SelectedIndices[i]);
				this.ListBoxDocuments.Items.RemoveAt(index);
			}
		}
		
		/// <summary>
		/// 在DataGridView中，添加新行时，将其搜索方向设置为“按行”。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void MyDataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
		{
			if (e.RowIndex >= 1)
			{
				var a = this.MyDataGridView1[2, e.RowIndex - 1];
				if (a.Value == null)
				{
					a.Value = "按行";
				}
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
			
			if (!this.BkgWk_Extract.IsBusy)
			{
				//
				this.blnIsBeingExtracting = true;
				//打开进行数据操作的Excel程序
				if (this.F_ExcelApp == null)
				{
					this.F_ExcelApp = new Microsoft.Office.Interop.Excel.Application();
					this.F_ExcelApp.WorkbookBeforeClose += this.F_ExcelApp_WorkbookBeforeClose;
					this.F_ExcelApp.DisplayAlerts = false;
					//一般情况下，默认是隐藏的，如果原来是打开的，则手动隐藏
					this.F_ExcelApp.Visible = false;
				}
				
				//打开进行数据提取的Word程序
				if (this.F_WordApp == null)
				{
					this.F_WordApp = new Microsoft.Office.Interop.Word.Application();
					this.F_WordApp.DocumentBeforeClose += this.F_WordApp_DocumentBeforeClose;
					this.F_WordApp.Visible = false;
				}
				
				
				//初始化错误列表
				this.F_ErrorList = new List<string>();
				//
				string strWorkBook_ExportedTo = this.txtbxSavePath.Text;
				
				
				//---------- 打开保存数据的工作簿，并提取其中的所有工作表 ----------------
				List<string> listPointsTag = new List<string>();
				List<string> listSheetNameInWkbk = new List<string>();
				try
				{
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
					this.F_BufferSheet = this.F_WorkBook_ExportedTo.Worksheets.Add() as Worksheet;
					//
					object AllSheets = F_WorkBook_ExportedTo.Worksheets;
					foreach (Worksheet shtInWorkbook in (IEnumerable) AllSheets)
					{
						listSheetNameInWkbk.Add(shtInWorkbook.Name);
					}
				}
				catch (Exception)
				{
					MessageBox.Show("保存数据的Word文档打开出错，请检查或者关闭此文档。", "Error", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				
				// ------------- 提取每一个工作表与Range范围的格式 -------------并返回DataGridView中的所有数据
				this.F_DicPointsInfo = SearchPointsInfo(this.F_WorkBook_ExportedTo);
				if (F_DicPointsInfo == null)
				{
					return;
				}
				
				
				// -----------进行数据提取的Document对象数组------------------------
				System.Windows.Forms.ListBox.ObjectCollection DocItems = this.ListBoxDocuments.Items;
				int DocsCount = DocItems.Count;
				//记录DataGridView控件中所有数据的数组
				string[] arrDocsPath = new string[DocsCount - 1 + 1];
				for (int i = 0; i <= DocsCount - 1; i++)
				{
					arrDocsPath[i] = DocItems[i].ToString();
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
				// ---------------------- 开始提取数据 ---------------------
				ExportToWorksheet Export = new ExportToWorksheet(F_WorkBook_ExportedTo, arrDocsPath, blnParseDateFromFilePath, this.F_BufferSheet);
				this.BkgWk_Extract.RunWorkerAsync(Export);
			}
		}
		private Dictionary<Worksheet, PointsInfoForExport> SearchPointsInfo(Microsoft.Office.Interop.Excel.Workbook wkbk)
		{
			Dictionary<Worksheet, PointsInfoForExport> listRangeInfo = new Dictionary<Worksheet, PointsInfoForExport>();
			string strTestRange = "";
			//
			//记录DataGridView控件中所有数据的数组
			try
			{
				int RowsCount = this.MyDataGridView1.Rows.Count;
				for (int RowIndex = 0; RowIndex <= RowsCount - 2; RowIndex++)
				{
					DataGridViewRow RowObject = MyDataGridView1.Rows[RowIndex];
					
					//要进行提取的测点标签
					string strPointName = RowObject.Cells[0].Value.ToString();
					//设置与测点标签对应的excel工作表对象，并为其命名
					Worksheet sht = null;
					try
					{
						sht = wkbk.Worksheets.Item[strPointName] as Worksheet;
					}
					catch (Exception)
					{
						//表示工作簿中没有这一工作表
						sht = wkbk.Worksheets.Add() as Worksheet;
						//为新创建的工作表命名
						bool blnNameOk = false;
						var shtName = strPointName;
						do
						{
							try
							{
								sht.Name = shtName;
								blnNameOk = true;
							}
							catch (Exception)
							{
								//表示此名称已经在工作簿中被使用了
								shtName = shtName + "1";
							}
						} while (!blnNameOk);
					}
					
					//测点数据距离测点标签的偏移位置
					byte Offset = byte.Parse(RowObject.Cells[1].Value.ToString());
					//搜索的方向：按行或者是按列
					Microsoft.Office.Interop.Excel.XlSearchOrder SearchDirection = default(Microsoft.Office.Interop.Excel.XlSearchOrder);
					DataGridViewComboBoxCell comboBox = (DataGridViewComboBoxCell) (RowObject.Cells[2]);
					if ((string) comboBox.Value == "按行")
					{
						SearchDirection = XlSearchOrder.xlByRows;
					}
					else if ((string) comboBox.Value == "按列")
					{
						SearchDirection = XlSearchOrder.xlByColumns;
					}
					else
					{
						MessageBox.Show("请先输入搜索方向", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return default(Dictionary<Worksheet, PointsInfoForExport>);
					}
					
					PointsInfoForExport RangeInfo = new PointsInfoForExport(strPointName, Offset, SearchDirection);
					listRangeInfo.Add(sht, RangeInfo);
				}
			}
			catch (Exception)
			{
				MessageBox.Show("数据的格式出错 : " + "\r\n" + strTestRange + "，请重新输入", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return default(Dictionary<Worksheet, PointsInfoForExport>);
			}
			return listRangeInfo;
		}
		
		//在后台线程中执行操作
		/// <summary>
		/// 在后台线程中执行操作
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void StartToDoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			//定义初始变量
			ExportToWorksheet ExportToWorksheet = (ExportToWorksheet) e.Argument;
			string[] arrDocsPath = ExportToWorksheet.arrDocsPath;
			Microsoft.Office.Interop.Excel.Workbook WorkBook_ExportedTo = ExportToWorksheet.WorkBook_ExportedTo;
			bool blnParseDateFromFilePath = ExportToWorksheet.ParseDateFromFilePath;
			Microsoft.Office.Interop.Excel.Worksheet bufferSheet = ExportToWorksheet.BufferSheet;
			
			//一共要处理的工作表数(工作簿个数*每个工作簿中提取的工作表数)，用来显示进度条的长度
			int Count_Documents = this.ListBoxDocuments.Items.Count;
			//
			int percent = 0;
			//每一份数据所对应的进度条长度
			float unit = 0;
			unit = (float) ((double) (this.ProgressBar1.Maximum - this.ProgressBar1.Minimum) / Count_Documents);
			//报告进度
			this.BkgWk_Extract.ReportProgress(percent, "");
			//开始提取数据
			for (short iDoc = 0; iDoc <= Count_Documents - 1; iDoc++)
			{
				string strDocPath = arrDocsPath[iDoc];
				Microsoft.Office.Interop.Word.Document Doc = null;
				try
				{
					//下面有可能会出现文档打开出错
					Doc = this.F_WordApp.Documents.Open(FileName:  strDocPath,ReadOnly: true, Visible: false);
					//
					Microsoft.Office.Interop.Word.Table myTable = default(Microsoft.Office.Interop.Word.Table);
					short CountTables = (short) Doc.Tables.Count;
					if (CountTables > 0)
					{
						for (short iTable = 1; iTable <= CountTables; iTable++)
						{
							myTable = Doc.Tables[iTable];
							// ------------- 正式开始提取数据 -------------
							
							ExportData(DataTableInWord:  myTable);
							
							// ------------- 正式开始提取数据 -------------
							
							this.BkgWk_Extract.ReportProgress(System.Convert.ToInt32((iDoc + (double) iTable / CountTables) * unit), "正在提取文档：" + strDocPath);
						} //文档中的下一个表格Table对象
					}
				}
				catch (Exception ex)
				{
					//文档打开出错
					string strError = "Document文档：" + Doc.FullName + " 打开时出错。  " + "\r\n" + ex.Message;
					this.F_ErrorList.Add(strError);
				}
				finally
				{
					if (Doc != null) //说明工作簿顺利打开
					{
						Doc.Close(SaveChanges: false);
					}
					this.BkgWk_Extract.ReportProgress(System.Convert.ToInt32((iDoc + 1) * unit), "正在提取文档：" + strDocPath);
				}
				
				//更新下一个文档的数据在对应的Excel工作表中所保存的列号
				//以及表头信息
				for (short iSheet = 0; iSheet <= F_DicPointsInfo.Count - 1; iSheet++)
				{
					Microsoft.Office.Interop.Excel.Worksheet sht = F_DicPointsInfo.Keys[iSheet] as Worksheet;
					PointsInfoForExport pointinfo = this.F_DicPointsInfo.Values(iSheet);
					//此工作簿所对应的表头的数据：工作簿的名称或者是工作簿中包含的日期信息
					string ColumnTitle = GetColumnTitle(strDocPath, blnParseDateFromFilePath);
					sht.Cells[cstRowNum_ColumnTitle, pointinfo.ColNumToBeWritten].Value = ColumnTitle;
					//
					pointinfo.ColNumToBeWritten += pointinfo.ColumnsCountToBeAdd;
					pointinfo.RowNumToBeWritten = cstRowNum_FirstData;
				}
			} //Next Document下一个文档
		}
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
		///  !!! 正式开始提取数据
		/// </summary>
		/// <param name="DataTableInWord">进行数据提取的word中的表格Table对象</param>
		/// <remarks>提取的基本思路：已有一个doc对象，并对其中的一个测点进行提取。</remarks>
		private void ExportData(Microsoft.Office.Interop.Word.Table DataTableInWord)
		{
			try
			{
				Microsoft.Office.Interop.Word.Range rgTable = DataTableInWord.Range;
				rgTable.Copy();
				this.F_BufferSheet.UsedRange.Clear();
				this.F_BufferSheet.Activate();
				this.F_BufferSheet.UsedRange.Clear();
				this.F_BufferSheet.Cells[1, 1].select();
				this.F_BufferSheet.Paste();
				
				
				//此文档中的每一个要提取的测点。
				foreach (Microsoft.Office.Interop.Excel.Worksheet sheetExportTo in this.F_DicPointsInfo.Keys)
				{
					PointsInfoForExport PointInfo = this.F_DicPointsInfo.Item(sheetExportTo);
					
					// ------------ 从暂存工作表中将测点标签与对应的数据提取到目标工作表中 ----------
					//搜索得到的第一个结果的range对象，如果没有搜索到，则返回nothing。
					Microsoft.Office.Interop.Excel.Range rgNextPoint = default(Microsoft.Office.Interop.Excel.Range);
					rgNextPoint = this.F_BufferSheet.UsedRange.Find(What: ref PointInfo.PointTag, SearchOrder: ref PointInfo.SearchOrder, LookAt: ref XlLookAt.xlPart, LookIn: ref XlFindLookIn.xlValues, SearchDirection: ref XlSearchDirection.xlNext, MatchCase: false);
					if (rgNextPoint != null)
					{
						//当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。
						//发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。
						string firstAddress = rgNextPoint.Address;
						//提取数据并写入最终的工作表
						do
						{
							sheetExportTo.Cells[PointInfo.RowNumToBeWritten, PointInfo.ColNumToBeWritten].Value = rgNextPoint.Value;
							sheetExportTo.Cells[PointInfo.RowNumToBeWritten, PointInfo.ColNumToBeWritten + 1].Value = rgNextPoint PointInfo.Offset[0, PointInfo.Offset].Value;
							PointInfo.RowNumToBeWritten++;
							rgNextPoint = this.F_BufferSheet.UsedRange.FindNext(rgNextPoint);
						} while (rgNextPoint != null && string.Compare(rgNextPoint.Address, firstAddress) != 0);
						
					}
				}
			}
			catch (Exception)
			{
				//数据提取出错
				string strError = "";
				this.F_ErrorList.Add(strError);
			}
			finally
			{
				
			}
		}
		
		//报告进度
		/// <summary>
		/// 报告进度
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BackgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
		{
			string strHandlePath = System.Convert.ToString(e.UserState.ToString());
			this.lbSheetName.Text = strHandlePath;
			this.ProgressBar1.Value = e.ProgressPercentage;
		}
		
		//操作完成
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
				//Call ReportError(ErrorFilePath, Me.F_ErrorList)
				
				//删除用来缓存数据的中间工作表
				this.F_BufferSheet.Delete();
				this.F_BufferSheet = null;
				
				//保存最终结果数据
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
			//
			this.blnIsBeingExtracting = false;
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
			this.lbSheetName.Text = "";
			this.ProgressTimer.Stop();
			this.ProgressTimer.Dispose();
			this.ProgressTimer = null;
			this.ProgressTimer.Tick += this.ProgressTimer_Tick;
		}
		
#endregion
		
	}
}
