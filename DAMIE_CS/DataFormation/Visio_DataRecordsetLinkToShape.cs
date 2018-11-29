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

using System.Data.OleDb;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
using Microsoft.Office.Interop;
//using DAMIE.Miscellaneous.AdoForExcel;

namespace CableStayedBridge
{
	/// <summary>
	/// Excel数据到Visio形状
	/// </summary>
	/// <remarks></remarks>
	public partial class Visio_DataRecordsetLinkToShape
	{
		
		
#region   ---  Properties
		/// <summary>
		/// 进行形状链接的文档
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Visio.Document F_vsoDoc;
		/// <summary>
		/// 进行形状链接的文档，设置此属性时会触发vsoDocumentChanged事件
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
private Microsoft.Office.Interop.Visio.Document vsoDoc
		{
			get
			{
				return F_vsoDoc;
			}
			set
			{
				this.F_vsoDoc = value;
				this.F_vsoDoc.BeforeDocumentClose += this.F_vsoDoc_BeforeDocumentClose;
				if (vsoDocumentChangedEvent != null)
					vsoDocumentChangedEvent(value);
			}
		}
		
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 当进行数据链接的Visio文档发生改变时触发
		/// </summary>
		/// <remarks></remarks>
		private Action<Microsoft.Office.Interop.Visio.Document> vsoDocumentChangedEvent;
		private event Action<Microsoft.Office.Interop.Visio.Document> vsoDocumentChanged
		{
			add
			{
				vsoDocumentChangedEvent = (Action<Microsoft.Office.Interop.Visio.Document>) System.Delegate.Combine(vsoDocumentChangedEvent, value);
			}
			remove
			{
				vsoDocumentChangedEvent = (Action<Microsoft.Office.Interop.Visio.Document>) System.Delegate.Remove(vsoDocumentChangedEvent, value);
			}
		}
		
		
		/// <summary>
		/// 在Visio文档通过验证，表示可以进行数据链接之时触发
		/// </summary>
		/// <remarks></remarks>
		private Action ShapeIDValidatedEvent;
		private event Action ShapeIDValidated
		{
			add
			{
				ShapeIDValidatedEvent = (Action) System.Delegate.Combine(ShapeIDValidatedEvent, value);
			}
			remove
			{
				ShapeIDValidatedEvent = (Action) System.Delegate.Remove(ShapeIDValidatedEvent, value);
			}
		}
		
		
		/// <summary>
		/// Visio的Application对象，此对象不包含在“群坑分析”的主程序中的那个Visio的Application对象
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Visio.Application F_vsoApplication;
		
		/// <summary>
		/// 进行形状链接的绘图页面
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Visio.Page F_vsoPage;
		
		/// <summary>
		/// 进行链接的数据记录集
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Visio.DataRecordset F_vsoDataRs;
		
		/// <summary>
		/// 在数据记录集中标识“形状ID”的字段列的下标值。在数据记录集中，每一行中的第一列（个）数据的下标值为0。
		/// </summary>
		/// <remarks></remarks>
		private int F_IndexOfShapeID;
		
		private ComboBox[] F_arrCombobox = new ComboBox[3];
		
#endregion
		
#region   ---  构造函数与窗体的加载
		
		public Visio_DataRecordsetLinkToShape()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			
			//设置组合列表框中要进行显示的属性
			string DisplayMember = LstbxDisplayAndItem.DisplayMember;
			Visio_DataRecordsetLinkToShape with_1 = this;
			F_arrCombobox[0] = ComboBox_Page;
			F_arrCombobox[1] = ComboBox_DataRs;
			F_arrCombobox[2] = ComboBox_Column_ShapeID;
			//
			foreach (ComboBox cbb in this.F_arrCombobox)
			{
				cbb.DisplayMember = DisplayMember;
			}
			//
			this.btnLink.Enabled = false;
			//
			
		}
		
		public void Visio_DataRecordsetLinkToShape_Load(object sender, EventArgs e)
		{
			//如果程序中已经有打开的Visio文档，则将此文档作为默认的进行形状链接的文档
			GlobalApplication GlobalApp = GlobalApplication.Application;
			if (GlobalApp != null)
			{
				if (GlobalApp.PlanView_VisioWindow != null)
				{
					Microsoft.Office.Interop.Visio.Document doc = GlobalApp.PlanView_VisioWindow.Page.Document;
					this.vsoDoc = doc;
					this.txtbxVsoDoc.Text = doc.FullName;
				}
			}
			
		}
		
#endregion
		
#region   ---  获取集合中的成员
		/// <summary>
		/// 从Visio文档中返回其中的所有Page对象的数组
		/// </summary>
		/// <param name="Doc"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private LstbxDisplayAndItem[] GetPagesFromDoc(Microsoft.Office.Interop.Visio.Document Doc)
		{
			short pagesCount = Doc.Pages.Count;
			LstbxDisplayAndItem[] arrItems = new LstbxDisplayAndItem[pagesCount - 1 + 1];
			short i = (short) 0;
			foreach (Microsoft.Office.Interop.Visio.Page page in Doc.Pages)
			{
				arrItems[i] = new LstbxDisplayAndItem(page.Name, page);
				i++;
			}
			return arrItems;
		}
		
		/// <summary>
		/// 从Visio文档中返回其中的所有DataRecordset对象的数组
		/// </summary>
		/// <param name="Doc"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private LstbxDisplayAndItem[] GetDataRsFromDoc(Microsoft.Office.Interop.Visio.Document Doc)
		{
			short DRSsCount = (short) Doc.DataRecordsets.Count;
			LstbxDisplayAndItem[] arrItems = new LstbxDisplayAndItem[DRSsCount - 1 + 1];
			short i = (short) 0;
			foreach (Microsoft.Office.Interop.Visio.DataRecordset DRS in Doc.DataRecordsets)
			{
				arrItems[i] = new LstbxDisplayAndItem(DRS.Name, DRS);
				i++;
			}
			return arrItems;
		}
		
		/// <summary>
		/// 从Visio文档的数据记录集中返回其中的字段列对象的数组
		/// </summary>
		/// <param name="DRS"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private LstbxDisplayAndItem[] GetColumnsFromDataRS(Microsoft.Office.Interop.Visio.DataRecordset DRS)
		{
			int ColumnsCount = DRS.DataColumns.Count;
			LstbxDisplayAndItem[] arrItems = new LstbxDisplayAndItem[ColumnsCount - 1 + 1];
			int i = 0;
			foreach (Microsoft.Office.Interop.Visio.DataColumn Column in DRS.DataColumns)
			{
				//在数据记录集中，第一列数据的Index为0。
				arrItems[i] = new LstbxDisplayAndItem(Column.DisplayName, i);
				i++;
			}
			return arrItems;
		}
		
#endregion
		
#region   ---  组合框的选择项发生改变
		
		public void ComboBox_DataRs_SelectedIndexChanged(object sender, EventArgs e)
		{
			LstbxDisplayAndItem lstItem = this.ComboBox_DataRs.SelectedItem;
			try
			{
				Microsoft.Office.Interop.Visio.DataRecordset DataRs = (Microsoft.Office.Interop.Visio.DataRecordset) lstItem.Value;
				this.F_vsoDataRs = DataRs;
				//更新数据记录集中的字段列。
				this.ComboBox_Column_ShapeID.DataSource = GetColumnsFromDataRS(DataRs);
			}
			catch (Exception)
			{
				//MessageBox.Show(ex.Message, "选择数据记录集出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			}
			this.btnLink.Enabled = false;
		}
		
		public void ComboBox_Page_SelectedIndexChanged(object sender, EventArgs e)
		{
			LstbxDisplayAndItem lstItem = this.ComboBox_Page.SelectedItem;
			try
			{
				this.F_vsoPage = (Microsoft.Office.Interop.Visio.Page) lstItem.Value;
			}
			catch (Exception)
			{
				//MessageBox.Show(ex.Message, "选择Visio页面出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			}
			this.btnLink.Enabled = false;
		}
		
		public void ComboBox_Column_ShapeID_SelectedIndexChanged(object sender, EventArgs e)
		{
			LstbxDisplayAndItem lstItem = this.ComboBox_Column_ShapeID.SelectedItem;
			try
			{
				this.F_IndexOfShapeID = (int) lstItem.Value;
			}
			catch (Exception)
			{
				//MessageBox.Show(ex.Message, "选择形状ID的字段出错！", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			}
			this.btnLink.Enabled = false;
		}
		
#endregion
		
#region   ---  按钮操作
		//验证
		public void btnValidate_Click(object sender, EventArgs e)
		{
			if (ValidateShapes(this.F_vsoPage, this.F_vsoDataRs, this.F_IndexOfShapeID))
			{
				MessageBox.Show("形状ID验证成功！", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				if (ShapeIDValidatedEvent != null)
					ShapeIDValidatedEvent();
			}
			else
			{
				this.btnLink.Enabled = false;
			}
		}
		//链接
		public void btnLink_Click(object sender, EventArgs e)
		{
			if (PassDataRecordsetToShape(this.F_vsoDataRs, this.F_vsoPage))
			{
				MessageBox.Show("形状ID数据链接到形状成功！", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			}
			else
			{
				MessageBox.Show("形状ID数据链接到形状失败！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		
#endregion
		
		//选择新的Visio文档
		/// <summary>
		/// 选择新的Visio文档
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnChooseVsoDoc_Click(object sender, EventArgs e)
		{
			
			string FilePath = "";
			this.OpenFileDialog1.Title = "选择进行数据链接的Visio文档";
			this.OpenFileDialog1.Filter = "Visio文件(*.vsd)|*.vsd";
			this.OpenFileDialog1.FilterIndex = 2;
			this.OpenFileDialog1.Multiselect = false;
			if (this.OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				FilePath = this.OpenFileDialog1.FileName;
			}
			else
			{
				return;
			}
			if (FilePath.Length > 0)
			{
				//
				this.txtbxVsoDoc.Text = FilePath;
				//
				if (this.F_vsoApplication == null)
				{
					this.F_vsoApplication = new Microsoft.Office.Interop.Visio.Application();
					this.F_vsoApplication.BeforeQuit += this.F_vsoApplication_BeforeQuit;
				}
				//
				try
				{
					this.vsoDoc = this.F_vsoApplication.Documents.Open(FilePath);
				}
				catch (Exception)
				{
					MessageBox.Show("Visio文档打开出错，请检查后重新打开。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				this.F_vsoApplication.Visible = true;
			}
			
		}
		
		//验证
		/// <summary>
		/// 验证页面中是否包含所有位于数据记录集中所记录的形状ID。
		/// </summary>
		/// <param name="page"></param>
		/// <param name="DRS"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private bool ValidateShapes(Microsoft.Office.Interop.Visio.Page page, Microsoft.Office.Interop.Visio.DataRecordset DRS, int intIndexOfShapeID)
		{
			bool blnValidated = true;
			if (DRS != null)
			{
				int[] lngRowIDs = DRS.GetDataRowIDs("");
				Microsoft.Office.Interop.Visio.Shape shp;
				foreach (int id in lngRowIDs)
				{
					object[] RowData = DRS.GetRowData(id);
					object shapeID = RowData[intIndexOfShapeID];
					try
					{
						shp = page.Shapes.ItemFromID(System.Convert.ToInt32(shapeID));
					}
					catch (Exception)
					{
						blnValidated = false;
						var Result = MessageBox.Show("在页面中没有找到与形状ID \"" + System.Convert.ToString(shapeID) + "\" 相匹配的形状。请仔细检查记录的形状ID值。", 
							"Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
						if (Result == Windows.Forms.DialogResult.OK)
						{
						}
						else if (Result == Windows.Forms.DialogResult.Cancel)
						{
							//不再提示这条错误
							goto endOfForLoop;
						}
					}
				}
endOfForLoop:
				1.GetHashCode() ; //VBConversions note: C# requires an executable line here, so a dummy line was added.
			}
			else
			{
				blnValidated = false;
			}
			return blnValidated;
		}
		
		//链接
		/// <summary>
		/// 将Visio中的外部数据链接到Page中的指定形状。
		/// 此操作的作用：通过Visio的数据图形功能，在对应的形状上显示出它所链接的数据，比如此图形对应的开挖深度。
		/// </summary>
		/// <param name="DataRS">数据链接的源数据记录集</param>
		/// <param name="Page">要进行数据链接的形状所在的Page</param>
		/// <param name="ColumnIndex_PrimaryKey">在数据记录集中，用来记录形状的名称的数据所在的列号。如果是第一列，则为0.</param>
		/// <param name="DeleteDataRecordset">是否要在数据记录集的数据链接到形状后，将此数据记录集删除。</param>
		/// <remarks></remarks>
		private dynamic PassDataRecordsetToShape(Microsoft.Office.Interop.Visio.DataRecordset DataRS, Microsoft.Office.Interop.Visio.Page Page, short ColumnIndex_PrimaryKey = 0, bool DeleteDataRecordset = false)
		{
			bool blnSucceeded = true;
			int[] IDs = null;
			//  ------------------ GetDataRowIDs ---------------------
			//获取数据记录集内所有行的 ID 组成的数组，其中每一行均代表一个数据记录。
			//若要不应用筛选器（即获取所有行），则传递一个空字符串 ("") 即可。
			IDs = DataRS.GetDataRowIDs("");
			//
			Microsoft.Office.Interop.Visio.Shape shp = default(Microsoft.Office.Interop.Visio.Shape);
			try
			{
				foreach (int RowID in IDs)
				{
					int shapeID = System.Convert.ToInt32(DataRS.GetRowData(RowID)[ColumnIndex_PrimaryKey]);
					//ItemFromID可以进行页面或者形状集合中的全局索引，即可以索引子形状中的嵌套形状，而Item一般只能索引其下的子形状。
					shp = Page.Shapes.ItemFromID(shapeID);
					shp.LinkToData(DataRS.ID, RowID, false);
				}
			}
			catch (Exception)
			{
				blnSucceeded = false;
			}
			
			//是否要在数据记录集的数据链接到形状后，将此数据记录集删除。
			if (DeleteDataRecordset)
			{
				DataRS.Delete();
				DataRS = null;
			}
			return blnSucceeded;
		}
		
#region   ---  用户定义的事件
		
		/// <summary>
		/// Visio文档改变
		/// </summary>
		/// <param name="vsoDoc"></param>
		/// <remarks></remarks>
		public void DocumentChanged(Microsoft.Office.Interop.Visio.Document vsoDoc)
		{
			//
			this.ComboBox_Page.DataSource = GetPagesFromDoc(vsoDoc);
			this.ComboBox_DataRs.DataSource = GetDataRsFromDoc(vsoDoc);
			//
			this.btnValidate.Enabled = true;
			this.btnLink.Enabled = false;
			//
		}
		
		/// <summary>
		/// 形状ID验证成功
		/// </summary>
		/// <remarks></remarks>
		public void Visio_DataRecordsetLinkToShape_ShapeIDValidated()
		{
			this.btnLink.Enabled = true;
		}
		
		/// <summary>
		/// Visio程序关闭
		/// </summary>
		/// <param name="app"></param>
		/// <remarks></remarks>
		private void F_vsoApplication_BeforeQuit(Microsoft.Office.Interop.Visio.Application app)
		{
			this.F_vsoApplication = null;
			this.F_vsoApplication.BeforeQuit += this.F_vsoApplication_BeforeQuit;
		}
		
		/// <summary>
		/// Visio文档关闭
		/// </summary>
		/// <param name="Doc"></param>
		/// <remarks></remarks>
		private void F_vsoDoc_BeforeDocumentClose(Microsoft.Office.Interop.Visio.Document Doc)
		{
			this.F_vsoDoc = null;
			this.F_vsoDoc.BeforeDocumentClose += this.F_vsoDoc_BeforeDocumentClose;
			this.F_vsoDataRs = null;
			//
			ClearUI();
		}
		
		/// <summary>
		/// 委托：在主程序界面上清空列表框的显示
		/// </summary>
		/// <remarks></remarks>
		private delegate void BeforeDocumentCloseHander();
		/// <summary>
		/// 在主程序界面上清空列表框的显示
		/// </summary>
		/// <remarks></remarks>
		public void ClearUI()
		{
			Visio_DataRecordsetLinkToShape with_1 = this;
			if (with_1.InvokeRequired)
			{
				//非UI线程，再次封送该方法到UI线程
				this.BeginInvoke(new BeforeDocumentCloseHander(this.ClearUI));
			}
			else
			{
				this.txtbxVsoDoc.Text = "";
				this.btnValidate.Enabled = false;
				this.btnLink.Enabled = false;
				foreach (ComboBox cbb in this.F_arrCombobox)
				{
					try
					{
						cbb.DataSource = null;
						cbb.DisplayMember = LstbxDisplayAndItem.DisplayMember;
					}
					catch (Exception)
					{
						Debug.Print("重新设置列表框的数据源出错！");
					}
				}
			}
		}
		
#endregion
		
		
	}
}
