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
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop;

namespace CableStayedBridge
{
	
	public partial class DiaFrm_PointsTreeView
	{
		
#region   ---  Constants
		
		/// <summary>
		/// 工作表“监测点编号与对应坐标”中，第一个测点所在的行号
		/// </summary>
		/// <remarks></remarks>
		const byte cstRowNum_FirstPoint = 2;
		/// <summary>
		/// 在TreeView中，表示监测点的数据在控件中的深度（即对应TreeNode的level属性）（第一级深度为0）
		/// </summary>
		/// <remarks></remarks>
		const byte cstDepth_PointInTreeView = 2;
#endregion
		
#region   ---  Fields
		
		/// <summary>
		/// 存放监测点位信息的Excel表格
		/// </summary>
		/// <remarks></remarks>
		private Microsoft.Office.Interop.Excel.Worksheet F_wkshtPoints;
		
		
		/// <summary>
		/// 指示绘制监测点位的窗口Form是否已经加载
		/// </summary>
		/// <remarks></remarks>
		private bool blnHasLoaded = false;
		
		/// <summary>
		/// Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
		/// </summary>
		/// <remarks></remarks>
		private ClsDrawing_PlanView.MonitorPointsInformation F_MonitorPointsInfo;
		
		/// <summary>
		/// 监测点位所要绘制的Visio图形对象
		/// </summary>
		/// <remarks></remarks>
		private ClsDrawing_PlanView F_PaintingPage;
		
		/// <summary>
		/// CAD坐标系与Visio坐标系进行线性转换的斜率与截距.kx、cx、ky、cy
		/// </summary>
		/// <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
		private GeneralMethods.Cdnt_Cvsion ConversionParameter;
		
		/// <summary>
		/// 在TreeView中选择的监测点的完整路径，索引对应的TreeNode对象
		/// </summary>
		/// <remarks>监测点只包括最后一级要进行绘图的测点，而不包括其父节点</remarks>
		private Dictionary<string, TreeNode> F_dicChosenPoints_TreeView = new Dictionary<string, TreeNode>();
		/// <summary>
		/// 选择的监测点，索引的item为传递到列表框中的文本
		/// </summary>
		/// <remarks></remarks>
		private Dictionary<string, TreeNode> F_dicListedPoints = new Dictionary<string, TreeNode>();
		
		/// <summary>
		/// 此visio图形中的所有监测点的形状的字典集合,
		/// Dictionary(Of 监测点在列表框中显示的文本,监测点的形状ID)
		/// </summary>
		/// <remarks></remarks>
		private Dictionary<string, short> F_dicVisioPoints = new Dictionary<string, short>();
		
#endregion
		
#region   ---  structures
		
		/// <summary>
		/// TreeView中的父项目的文本与其在大数组中对应的行号区间
		/// </summary>
		/// <remarks></remarks>
		public struct RowsToDictionary
		{
			/// <summary>
			/// 监测项目或者基坑ID的数据在大数组中对应的行号的区域。
			/// 数组中有两个元素，第一个表示此项目中的第一个元素的行号，第二个表示此项目中的最后一个元素的行号。
			/// </summary>
			/// <remarks></remarks>
			public int[] RowsSpan;
			/// <summary>
			/// 一个字典对象
			/// </summary>
			/// <remarks></remarks>
			public object dic;
		}
		
		
		/// <summary>
		/// Visio中的监测点的形状的基本信息
		/// </summary>
		/// <remarks></remarks>
		public struct PointInVisio
		{
			/// <summary>
			/// 监测项目的名称，如“测斜”
			/// </summary>
			/// <remarks></remarks>
			public string strItem;
			/// <summary>
			/// 监测点的编号，如“CX1”
			/// </summary>
			/// <remarks></remarks>
			public string strPoint;
			/// <summary>
			/// 监测点在Visio中的坐标值(x,y)
			/// </summary>
			/// <remarks></remarks>
			public double[] Coordinates;
		}
#endregion
		
		/// <summary>
		/// 构造函数
		/// </summary>
		/// <param name="PaintingPage">监测点位所要绘制的Visio图形对象</param>
		/// <param name="wkshtPoints">存放监测点位信息的Excel表格</param>
		/// <remarks></remarks>
		public DiaFrm_PointsTreeView(Microsoft.Office.Interop.Excel.Worksheet wkshtPoints, 
			ClsDrawing_PlanView PaintingPage)
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			this.F_PaintingPage = PaintingPage;
			this.F_MonitorPointsInfo = PaintingPage.MonitorPointsInfo;
			this.F_wkshtPoints = wkshtPoints;
		}
		
		/// <summary>
		/// 窗口加载
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks>先提取坐标变换的斜率及截距；再从监测点编号与其对应的坐标创建树形列表</remarks>
		public void FrmTreeView_Load(object sender, EventArgs e)
		{
			//blnHasLoaded是为了解决showdialog在每次加载窗口时都触发一次load事件的问题。
			if (!blnHasLoaded)
			{
				try
				{
					// ------------------ 提取坐标变换的斜率及截距
					Microsoft.Office.Interop.Visio.Shape CP1 = default(Microsoft.Office.Interop.Visio.Shape);
					Microsoft.Office.Interop.Visio.Shape CP2 = default(Microsoft.Office.Interop.Visio.Shape);
					CP1 = this.F_PaintingPage.Page.Shapes.ItemFromID(F_MonitorPointsInfo.pt_Visio_BottomLeft_ShapeID);
					CP2 = this.F_PaintingPage.Page.Shapes.ItemFromID(F_MonitorPointsInfo.pt_Visio_UpRight_ShapeID);
					//visio中的坐标以inch为单位
					double xp1 = 0;
					double yp1 = 0;
					double xp2 = 0;
					double yp2 = 0;
					//x,y,xp,yp都是以inch为单位的
					double x = 0;
					double y = 0;
					x = System.Convert.ToDouble(CP1.Cells("LocPinX").Result(Microsoft.Office.Interop.Visio.VisUnitCodes.visInches));
					y = System.Convert.ToDouble(CP1.Cells("LocPinY").Result(Microsoft.Office.Interop.Visio.VisUnitCodes.visInches));
					CP1.XYToPage(x, y, ref xp1, ref yp1);
					x = System.Convert.ToDouble(CP2.Cells("LocPinX").Result(Microsoft.Office.Interop.Visio.VisUnitCodes.visInches));
					y = System.Convert.ToDouble(CP2.Cells("LocPinY").Result(Microsoft.Office.Interop.Visio.VisUnitCodes.visInches));
					CP2.XYToPage(x, y, ref xp2, ref yp2);
					ConversionParameter = GeneralMethods.Coordinate_Conversion(this.F_MonitorPointsInfo.pt_CAD_BottomLeft.X, this.F_MonitorPointsInfo.pt_CAD_BottomLeft.Y, 
						this.F_MonitorPointsInfo.pt_CAD_UpRight.X, this.F_MonitorPointsInfo.pt_CAD_UpRight.Y, 
						xp1, yp1, xp2, yp2);
					
					//从监测点编号与其对应的坐标创建树形列表
					constructTreeView(F_wkshtPoints, TreeViewPoints);
					//标记：此窗口已经在主程序中加载过，后面就不用再加载了。因为不会对其进行close，只会将其隐藏
					blnHasLoaded = true;
				}
				catch (Exception ex)
				{
					
					MessageBox.Show("不能正确地提取监测点位信息。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					//Exit Sub
					//Me.Close()
					this.Dispose();
				}
			}
		}
		
		/// <summary>
		/// 从监测点编号与其对应的坐标创建树形列表
		/// </summary>
		/// <param name="wksht">保存数据的工作表</param>
		/// <param name="myTreeView">用于列举监测点编号的TreeView对象</param>
		/// <remarks></remarks>
		private void constructTreeView(Microsoft.Office.Interop.Excel.Worksheet wksht, TreeView myTreeView)
		{
			
			//提取数据工作表中所有测点的数据,确保其中至少有一行数据
			object arrData = null;
			// ************ 看情况进行排序 ************
			//
			// ***************************************
			if (wksht.UsedRange.Rows.Count - cstRowNum_FirstPoint < 0)
			{
				MessageBox.Show("No data rows found in the specified worksheet" + "\r\n" + "please check the DataBase file or search another worksheet", 
					"Warning", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
				return;
			}
			else
			{
				arrData = wksht.UsedRange.Rows[cstRowNum_FirstPoint + ":" + System.Convert.ToString(wksht.UsedRange.Rows.Count)].value;
			}
			//大数组的数据行的上、下标值
			byte lb = 0;
			int ub = 0;
			lb = (byte) 0;
			ub = arrData.Length - 1;
			//
			//-*****************************************************-
			// ------ DicItem ------------ 第一列：监测项目
			Dictionary<string, RowsToDictionary> DicItems = new Dictionary<string, RowsToDictionary>();
			
			//              cx1:(x,y)
			//          ID1
			//              cx2:(x,y)
			//   Item1
			//              cx1:(x,y)
			//          ID2
			//              cx2:(x,y)
			//DicItems
			//              cx1:(x,y)
			//          ID1
			//              cx2:(x,y)
			//   Item2
			//              cx1:(x,y)
			//          ID2
			//              cx2:(x,y)
			
			RowsToDictionary rtd = new RowsToDictionary();
			if (ub == lb)
			{
				rtd = new RowsToDictionary();
				rtd.RowsSpan = new[] {lb, lb};
				DicItems.Add(arrData[lb, lb], rtd);
			}
			else
			{
				//
				int rowOldItem = lb;
				for (int rowNewItem = lb; rowNewItem <= ub - 1; rowNewItem++)
				{
					if (string.Compare(System.Convert.ToString(arrData[rowNewItem, lb]), System.Convert.ToString(arrData[rowNewItem + 1, lb]), true) != 0)
					{
						rtd = new RowsToDictionary();
						rtd.RowsSpan = new[] {rowOldItem, rowNewItem};
						DicItems.Add(arrData[rowNewItem, lb], rtd);
						//QueItem.Enqueue({arrData(rowNewItem, lb), rowOldItem, rowNewItem, Nothing})
						rowOldItem = rowNewItem + 1;
					}
				} //下一个监测项目
				//处理最后一个监测项目，它可以与倒数第二个不一样，这时要单独拿出来处理，因为它并不包含在上面的循环中。
				if (string.Compare(System.Convert.ToString(arrData[ub, lb]), System.Convert.ToString(arrData[ub - 1, lb]), true) == 0)
				{
					rtd = new RowsToDictionary();
					rtd.RowsSpan = new[] {rowOldItem, ub};
					DicItems.Add(arrData[ub, lb], rtd);
				}
				else
				{
					rtd = new RowsToDictionary();
					rtd.RowsSpan = new[] {rowOldItem, ub - 1};
					DicItems.Add(arrData[rowOldItem, lb], rtd);
					//
					rtd = new RowsToDictionary();
					rtd.RowsSpan = new[] {ub, ub};
					DicItems.Add(arrData[rowOldItem, lb], rtd);
				}
			}
			
			//-*****************************************************-
			// ------ DicIDs ---------- 第二列：基坑ID
			
			int startrow = 0;
			int endrow = 0;
			for (int i_item = 0; i_item <= DicItems.Count - 1; i_item++)
			{
				RowsToDictionary struct_value = DicItems.Values(i_item);
				startrow = struct_value.RowsSpan[0];
				endrow = struct_value.RowsSpan[1];
				//
				Dictionary<string, RowsToDictionary> DicIDs = new Dictionary<string, RowsToDictionary>();
				//
				if (startrow == endrow) //如果此监测项目下只有一行数据
				{
					rtd = new RowsToDictionary();
					rtd.RowsSpan = new[] {startrow, startrow};
					DicIDs.Add(arrData[startrow, lb + 1], rtd);
				}
				else
				{
					//
					//
					int rowOldID = startrow;
					for (int rowNewID = startrow; rowNewID <= endrow - 1; rowNewID++)
					{
						if (string.Compare(System.Convert.ToString(arrData[rowNewID, lb + 1]), System.Convert.ToString(arrData[rowNewID + 1, lb + 1]), true) != 0)
						{
							rtd = new RowsToDictionary();
							rtd.RowsSpan = new[] {rowOldID, rowNewID};
							DicIDs.Add(arrData[rowNewID, lb + 1], rtd);
							
							rowOldID = rowNewID + 1;
						}
					} //下一个基坑ID
					
					//处理最后一个基坑，它可以与倒数第二个不一样，这时要单独拿出来处理，因为它并不包含在上面的循环中。
					if (string.Compare(System.Convert.ToString(arrData[endrow, lb + 1]), System.Convert.ToString(arrData[endrow - 1, lb + 1]), true) == 0)
					{
						rtd = new RowsToDictionary();
						rtd.RowsSpan = new[] {rowOldID, endrow};
						DicIDs.Add(arrData[rowOldID, lb + 1], rtd);
						
					}
					else
					{
						rtd = new RowsToDictionary();
						rtd.RowsSpan = new[] {rowOldID, endrow - 1};
						DicIDs.Add(arrData[rowOldID, lb + 1], rtd);
						//
						rtd = new RowsToDictionary();
						rtd.RowsSpan = new[] {endrow, endrow};
						DicIDs.Add(arrData[endrow, lb + 1], rtd);
						//
						
					}
				}
				struct_value.dic = DicIDs;
				DicItems.Item(DicItems.Keys(i_item)) = struct_value;
			} //下一个监测项目
			
			//-*****************************************************-
			// ------ TreeViewPoints ---------- 第三列：测点编号
			
			TreeViewPoints.Nodes.Clear();
			
			for (int i = 0; i <= DicItems.Keys.Count - 1; i++)
			{
				
				string strItem = System.Convert.ToString(DicItems.Keys(i));
				TreeNode itemNode = new TreeNode(strItem);
				TreeViewPoints.Nodes.Add(itemNode);
				
				//dicid：每一个监测项目下的基坑ID的集合
				Dictionary<string, RowsToDictionary> dicID = DicItems.Values(i).dic;
				
				//---------
				for (int j = 0; j <= dicID.Keys.Count - 1; j++)
				{
					
					//For Each struct_ID As RowsToDictionary In dicID.Values
					string strID = System.Convert.ToString(dicID.Keys(j));
					TreeNode IDNode = new TreeNode(strID);
					itemNode.Nodes.Add(IDNode);
					
					
					//dicPoints：每一个基坑中的测点的集合
					Dictionary<string, double[]> dicPoints = new Dictionary<string, double[]>();
					
					//坐标点赋值
					int[] rowspan = new int[2];
					rowspan = dicID.Values(j).RowsSpan;
					for (int row = rowspan[0]; row <= rowspan[1]; row++)
					{
						
						//获取测点编号
						string strPoint = System.Convert.ToString(arrData[row, lb + 2]);
						//获取测点对应的坐标，并进行坐标转换，将工作表中的CAD坐标系中的坐标转换成Visio中的坐标系（以inch为单位）
						double[] Coodinates = new double[2];
						Coodinates = new[] {arrData[row, lb + 3] * ConversionParameter.kx + ConversionParameter.cx, System.Convert.ToDouble(arrData[row, lb + 4]) * ConversionParameter.ky + ConversionParameter.cy};
						//
						dicPoints.Add(arrData[row, lb + 2], Coodinates);
						TreeNode NdPoint = IDNode.Nodes.Add(strPoint);
						//将测点编号链接对应的坐标值
						NdPoint.Tag = Coodinates;
						
					} //下一个测点
					RowsToDictionary @struct = dicID.Values(j);
					@struct.dic = dicPoints;
					dicID.Item(dicID.Keys(j)) = @struct;
					// --- Note ----
					//dicID.Item(dicID.Keys(j)).dic = dicPoints
					//上面的语句失败，因为它仅为 dicID.Item(dicID.Keys(j)) 属性返回的 RowsToDictionary 结构提供了临时的分配。
					//这是一个值类型的结构，该语句运行后不保留临时结构。
					//解决该问题的方法是：为 RowsToDictionary 结构的属性声明一个变量，并使用这个变量，从而为 RowsToDictionary 结构创建更为永久的分配。
					// \--- Note ----
				} //下一个基坑
			} //下一个监测项目
			//
		}
		
#region   ---  控件操作
		
		/// <summary>
		/// 在TreeView的Checkbox框中进行选择或者取消选择时的操作
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void TreeViewPoints_AfterCheck(object sender, TreeViewEventArgs e)
		{
			TreeNode nd = e.Node;
			bool blnChecked = nd.Checked;
			if (nd.Level == cstDepth_PointInTreeView)
			{
				try
				{
					if (blnChecked)
					{
						F_dicChosenPoints_TreeView.Add(nd.FullPath, nd);
					}
					else
					{
						F_dicChosenPoints_TreeView.Remove(nd.FullPath);
					}
				}
				catch (Exception)
				{
				}
			}
			//将节点的子节点也全部选中或全部取消选择
			foreach (TreeNode childNode in nd.Nodes)
			{
				childNode.Checked = blnChecked;
			}
		}
		
		/// <summary>
		/// 将TreeView中选择的测点添加进列表中
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnAdd_Click(object sender, EventArgs e)
		{
			
			List<LstbxDisplayAndItem> lstItem = new List<LstbxDisplayAndItem>();
			string[] keys = F_dicChosenPoints_TreeView.Keys.ToArray;
			TreeNode[] values = F_dicChosenPoints_TreeView.Values.ToArray;
			//
			ListBoxChosenItems.DataSource = F_dicChosenPoints_TreeView.Keys.ToArray;
			//重新构造字典dicListedPoints
			
			F_dicListedPoints.Clear();
			for (int i = 0; i <= F_dicChosenPoints_TreeView.Count - 1; i++)
			{
				F_dicListedPoints.Add(F_dicChosenPoints_TreeView.Keys(i), F_dicChosenPoints_TreeView.Values(i));
			}
		}
		
		/// <summary>
		/// 移除选定的测点
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnRemove_Click(object sender, EventArgs e)
		{
			var indices = ListBoxChosenItems.SelectedIndices;
			//
			string[] strData = ListBoxChosenItems.DataSource;
			List<string> lst = new List<string>();
			foreach (string s in strData)
			{
				lst.Add(s);
			}
			//
			for (var i = indices.Count - 1; i >= 0; i--)
			{
				int index = indices[i];
				lst.RemoveAt(index);
				F_dicListedPoints.Remove(F_dicListedPoints.Keys(index));
			}
			ListBoxChosenItems.DataSource = lst.ToArray;
		}
		
		/// <summary>
		/// 清空最后的测点
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnClear_Click(object sender, EventArgs e)
		{
			ListBoxChosenItems.DataSource = null;
			F_dicListedPoints.Clear();
		}
		
		/// <summary>
		/// CollapseAll
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnColpsAll_Click(object sender, EventArgs e)
		{
			TreeViewPoints.CollapseAll();
		}
		
		/// <summary>
		/// 清空treeview中所有节点的选择
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks>这里将代码限死为3层节点，这种方法有待改进</remarks>
		public void ClearAllCheckedNode(object sender, EventArgs e)
		{
			TreeViewPoints.AfterCheck -= TreeViewPoints_AfterCheck;
			F_dicChosenPoints_TreeView.Clear();
			//取消treeview中所有节点的选择
			foreach (TreeNode nd0 in TreeViewPoints.Nodes)
			{
				nd0.Checked = false;
				foreach (TreeNode nd1 in nd0.Nodes)
				{
					nd1.Checked = false;
					foreach (TreeNode nd2 in nd1.Nodes)
					{
						nd2.Checked = false;
					}
				}
			}
			TreeViewPoints.AfterCheck += TreeViewPoints_AfterCheck;
		}
		
#endregion
		
		/// <summary>
		/// 在点击“确定”时，是否还需要再执行一次绘制测点的操作，
		/// 因为有可能从前面进行预览之后，所要绘制的点位就没有发生改变，所以此时点击确认就只需要直接将窗口关闭就可以了。
		/// </summary>
		/// <remarks></remarks>
		private bool blnRefreshed = false;
		public void btnPreview_Click(object sender, EventArgs e)
		{
			DrawPoints();
		}
		
		/// <summary>
		/// 点击确定时将窗口隐藏
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks>如果已经进行过预览，则直接隐藏窗口，如果没有进行过预览，则先进行一次预览，再隐藏窗口</remarks>
		public void BtnOk_Click(object sender, EventArgs e)
		{
			if (!blnRefreshed)
			{
				DrawPoints();
			}
			this.Hide();
		}
		/// <summary>
		/// 点击关闭窗体时将窗体隐藏
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks>对于模式窗体（以showdialog显示），在点击关闭时，其实就是执行hide操作</remarks>
		public void FrmPointsTreeView_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.Hide();
			e.Cancel = true;
		}
		
		/// <summary>
		/// ！进行监测点位的绘制
		/// </summary>
		/// <remarks></remarks>
		private void DrawPoints()
		{
			ClsDrawing_PlanView vso = GlobalApplication.Application.PlanView_VisioWindow;
			if (vso != null)
			{
				Microsoft.Office.Interop.Visio.Page pg = vso.Page;
				try
				{
					Microsoft.Office.Interop.Visio.Window vsoWindow = pg.Application.ActiveWindow;
					vsoWindow.Page = pg;
					vsoWindow.Application.ShowChanges = false;
					//
					string[] arrListedPoints = F_dicListedPoints.Keys.ToArray;
					string[] arrExistedPoints = F_dicVisioPoints.Keys.ToArray;
					//
					object[] arrAddOrRemove = new object[2];
					arrAddOrRemove = GetPointsToBeProcessed(arrExistedPoints, arrListedPoints);
					string[] arrToBeAdded = arrAddOrRemove[0];
					string[] arrToBeRemoved = arrAddOrRemove[1];
					
					
					// ----- arrPointsToBeAdded ------------- 处理要进行添加的图形
					PointInVisio[] arrPointsToBeAdded = new PointInVisio[arrToBeAdded.Length - 1 + 1];
					int i = 0;
					string strSeparator = TreeViewPoints.PathSeparator;
					foreach (string tag_Add in arrToBeAdded)
					{
						TreeNode nd = F_dicListedPoints.Item(tag_Add);
						PointInVisio struct_Point = new PointInVisio();
						var str = tag_Add;
						struct_Point.strItem = str.Substring(0, str.IndexOf(strSeparator));
						struct_Point.strPoint = nd.Text;
						struct_Point.Coordinates = nd.Tag;
						arrPointsToBeAdded[i] = struct_Point;
						i++;
					}
					
					AddMonitorPoints(vsoWindow, arrToBeAdded, arrPointsToBeAdded);
					
					// ----- arrToBeRemoved ------------- 处理要进行删除的图形
					foreach (string strPointTag in arrToBeRemoved)
					{
						Microsoft.Office.Interop.Visio.Shape shp = default(Microsoft.Office.Interop.Visio.Shape);
						shp = pg.Shapes.ItemFromID(System.Convert.ToInt32(F_dicVisioPoints.Item(strPointTag)));
						shp.Delete();
						//
						F_dicVisioPoints.Remove(strPointTag);
					}
					
					// -----------------------
					this.blnRefreshed = true;
					vsoWindow.Application.ShowChanges = true;
				}
				catch (Exception ex)
				{
					Debug.Print(ex.Message);
					MessageBox.Show("出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			else
			{
				MessageBox.Show("Visio绘图已经关闭，请重新打开。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}
		
		/// <summary>
		/// 如果列表框的数据源发生改变，则在点击确定时必须要重新绘制测点
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ListBoxChosenItems_DataSourceChanged(object sender, EventArgs e)
		{
			this.blnRefreshed = false;
		}
		
#region   ---  本地子方法
		
		/// <summary>
		/// 根据Visio中与列表框中的测点字符的差异，得到分别要进行添加和删除的测点
		/// </summary>
		/// <param name="arrExistedPoints">在Visio中存在的测点的字符数据</param>
		/// <param name="arrListedPoints">在ListBox列表中存在的测点的字符数据</param>
		/// <returns>返回一个数组，数组中有两个元素，元素各自又都为一个数组，第一个为要在Visio进行添加的测点的字符的数组；
		/// 第二个为要从Visio进行删除的测点图形的字符的数组</returns>
		/// <remarks>这里的“字符数据”都是指测点在TreeView中的完整路径的字符</remarks>
		private object[] GetPointsToBeProcessed(string[] arrExistedPoints, string[] arrListedPoints)
		{
			object[] arrAddOrRemove = new object[2];
			
			//将两个数组中的元素进行汇总组合并排序，且去掉其中中重复项
			SortedSet<string> ssMixedPoints = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
			foreach (var ep in arrExistedPoints)
			{
				ssMixedPoints.Add(ep);
			}
			foreach (var lp in arrListedPoints)
			{
				ssMixedPoints.Add(lp);
			}
			//
			List<string> listToBeAdded = new List<string>();
			List<string> listToBeRemoved = new List<string>();
			//
			short PointsCount = (short) ssMixedPoints.Count;
			SByte sbIn_E = default(SByte);
			SByte sbIn_L = default(SByte);
			SByte cmp = default(SByte);
			for (int i = 0; i <= PointsCount - 1; i++)
			{
				if (arrExistedPoints.Contains(ssMixedPoints[i]))
				{
					sbIn_E = 1;
				}
				else
				{
					sbIn_E = 0;
				}
				if (arrListedPoints.Contains(ssMixedPoints[i]))
				{
					sbIn_L = 1;
				}
				else
				{
					sbIn_L = 0;
				}
				//对上面两个数组中分别的对应元素进行比较——相减
				cmp = sbIn_E - sbIn_L;
				if (cmp == 1) //说明Visio中存在而在列表中没有列出，那么在处理时要从visio中将此图形删除
				{
					listToBeRemoved.Add(ssMixedPoints[i]);
				} //说明Visio中没有，但是列表中有，那么要在Visio中添加进此图形
				else if (cmp == -1)
				{
					listToBeAdded.Add(ssMixedPoints[i]);
				}
			}
			return new[] {listToBeAdded.ToArray, listToBeRemoved.ToArray};
		}
		
		/// <summary>
		/// 根据指定的监测点在visio图形中绘出对应的测点图形
		/// </summary>
		/// <param name="vsoWindow">进行绘图的visio图形</param>
		/// <param name="strTags">要绘制的所有监测点的在列表中对应的文本</param>
		/// <param name="PointsToBeAdded">要绘制的所有监测点的集合</param>
		/// <remarks></remarks>
		public void AddMonitorPoints(Microsoft.Office.Interop.Visio.Window vsoWindow, string[] strTags, PointInVisio[] PointsToBeAdded)
		{
			int n = PointsToBeAdded.Count();
			if (n > 0)
			{
				
				//
				object[] arrMaster = new object[n - 1 + 1];
				double[] arrCoordinate = new double[2 * n - 1 + 1];
				short[] arrIDOut = new short[n - 1 + 1];
				string[] arrPointTag = new string[n - 1 + 1];
				//
				
				Microsoft.Office.Interop.Visio.Masters DocumentMasters = vsoWindow.Document.Masters;
				
				int i = 0;
				foreach (var pt in PointsToBeAdded)
				{
					try
					{
						arrMaster[i] = DocumentMasters[pt.strItem];
					}
					catch (Exception)
					{
						Debug.Print("测点类型： {0} 在Visio模板文件中并不存在，此时用\"OtherTypes\"来表示。", pt.strItem);
						arrMaster[i] = DocumentMasters["OtherTypes"];
					}
					arrPointTag[i] = System.Convert.ToString(pt.strPoint);
					arrCoordinate[2 * i] = System.Convert.ToDouble(pt.Coordinates(0));
					arrCoordinate[2 * i + 1] = System.Convert.ToDouble(pt.Coordinates(1));
					i++;
				}
				
				Microsoft.Office.Interop.Visio.Page pg = vsoWindow.Page;
				int ProcessedCount = pg.DropMany(ref arrMaster, ref arrCoordinate, ref arrIDOut);
				
				//设置每一个测点形状上显示的文字
				pg.Application.ShowChanges = true;
				string tag = F_MonitorPointsInfo.ShapeName_MonitorPointTag;
				try
				{
					
					for (int i_p = 0; i_p <= n - 1; i_p++)
					{
						int id = arrIDOut[i_p];
						//设置主控形状的实例对象中，“某形状”的文本，“某形状”的索引等规范化。
						pg.Shapes.ItemFromID(id).Shapes.Item(tag).Text = arrPointTag[i_p];
						F_dicVisioPoints.Add(strTags[i_p], arrIDOut[i_p]);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("主控形状中没有找到子形状：" + "tag" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
		
#endregion
		
		
		
	}
}
