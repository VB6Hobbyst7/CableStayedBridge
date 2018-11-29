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

using System.ComponentModel;
using CableStayedBridge.DataBase;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
using Microsoft.Office.Interop.Visio;


namespace CableStayedBridge
{
	namespace All_Drawings_In_Application
	{
		
		/// <summary>
		/// Visio的开挖平面图
		/// </summary>
		/// <remarks></remarks>
		public class ClsDrawing_PlanView : IAllDrawings, IRolling
		{
			
#region   ---  Declarations & Definitions
			
#region   ---  Types
			
			/// <summary>
			/// Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
			/// </summary>
			/// <remarks></remarks>
			public struct MonitorPointsInformation
			{
				
				// ------ 不删！！！ ------ 不删！！！ ------ 不删！！！ ------
				// ''' <summary>
				// ''' visio中包含两个定位点的组合形状的形状名
				// ''' </summary>
				// ''' <remarks></remarks>
				//Const cstShapeName_OverView As String = "OVERVIEW"
				// ''' <summary>
				// ''' visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性。
				// ''' </summary>
				// ''' <remarks></remarks>
				//Const cstShapeName_MonitorPointTag As String = "Tag"
				// ''' <summary>
				// ''' Visio平面图中用于坐标变换的两个定位点的形状ID，
				// ''' 这两个点分别代表ABCD基坑群的左下角与右上角。
				// ''' </summary>
				// ''' <remarks></remarks>
				//Const cstShapeName_ConversionPoint1 As String = "Location Reference 1"
				//Const cstShapeName_ConversionPoint2 As String = "Location Reference 2"
				//'CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
				//'Const cstCADLocation_Point1 As New PointF(309598.527, -119668.436)
				//'Const cstCADLocation_Point2 As New PointF(536642.644, 201852.14)
				//Const cstCADLocation_x_Point1 As Single = 309598.527
				//Const cstCADLocation_y_Point1 As Single = -119668.436
				//Const cstCADLocation_x_Point2 As Single = 536642.644
				//Const cstCADLocation_y_Point2 As Single = 201852.14
				//
				
				/// <summary>
				/// visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性。
				/// </summary>
				/// <remarks>在表示监测点位的主控形状的模板中，每一个监测点位的主控形状中，
				/// 都有一个子形状，其Name属性为Tag，以此来索引此文本形状。</remarks>
				public string ShapeName_MonitorPointTag;
				/// <summary>
				/// Visio平面图中用于坐标变换的两个定位点的形状ID，
				/// 这两个点分别代表ABCD基坑群的左下角与右上角。
				/// </summary>
				/// <remarks></remarks>
				public int pt_Visio_BottomLeft_ShapeID;
				/// <summary>
				/// Visio平面图中用于坐标变换的两个定位点的形状ID，
				/// 这两个点分别代表ABCD基坑群的左下角与右上角。
				/// </summary>
				/// <remarks></remarks>
				public int pt_Visio_UpRight_ShapeID;
				/// <summary>
				/// CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
				/// </summary>
				/// <remarks></remarks>
				public PointF pt_CAD_BottomLeft;
				/// <summary>
				/// CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
				/// </summary>
				/// <remarks></remarks>
				public PointF pt_CAD_UpRight;
				//"OVERVIEW","Tag",197,217,New PointF(309598.527, -119668.436),New PointF(536642.644, 201852.14)"
				
			}
			
#endregion
			
#region   ---  Constants
			
			const string cstDrawingTag = "平面开挖分区";
#endregion
			
#region   ---  Events
			
			/// <summary>
			/// 鼠标在Visio绘图的窗口中双击
			/// </summary>
			/// <remarks></remarks>
			private Action WindowDoubleClickEvent;
			private event Action WindowDoubleClick
			{
				add
				{
					WindowDoubleClickEvent = (Action) System.Delegate.Combine(WindowDoubleClickEvent, value);
				}
				remove
				{
					WindowDoubleClickEvent = (Action) System.Delegate.Remove(WindowDoubleClickEvent, value);
				}
			}
			
			
#endregion
			
#region   --- Properties
			
			/// <summary>
			/// Visio程序界面
			/// </summary>
			/// <remarks></remarks>
			private Application P_Application;
			/// <summary>
			/// Visio程序界面
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Application Application
			{
				get
				{
					return P_Application;
				}
				private set
				{
					//在关闭Visio程序时不弹出保存文档的对话框
					value.AlertResponse = (short) 7; //0表示弹出任何警告或模型对话框，7表示不弹出对话框而默认选择IDNo。
					P_Application = value;
					P_Application.BeforeQuit += Application_Quit;
				}
			}
			
			/// <summary>
			/// 进行绘画的那一个页面
			/// </summary>
			/// <remarks></remarks>
			private Page P_Page;
			/// <summary>
			/// 进行绘画的那一个页面
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Page Page
			{
				get
				{
					return P_Page;
				}
				set
				{
					P_Page = value;
				}
			}
			
			/// <summary>
			/// Visio绘图中的窗口，此窗口并不限制于页面所在的范围，而是可以扩展到页面范围以外。
			/// </summary>
			/// <remarks></remarks>
			private Window P_Window;
			/// <summary>
			/// Visio绘图中的窗口，此窗口并不限制于页面所在的范围，而是可以扩展到页面范围以外。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Window Window
			{
				get
				{
					return this.P_Window;
				}
				set
				{
					this.P_Window = value;
					this.P_Window.MouseUp += this.DoubleClick_Up;
				}
			}
			
			/// <summary>
			/// 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据，
			/// 其中分别以向量的形式记录了这两列数据。向量的第一个元素的下标值为0
			/// </summary>
			/// <remarks></remarks>
			private clsData_ShapeID_FinishedDate P_ShapeID_FinishedDate;
			/// <summary>
			/// 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据，
			/// 其中分别以向量的形式记录了这两列数据。向量的第一个元素的下标值为0
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public clsData_ShapeID_FinishedDate ShapeID_FinishedDate
			{
				get
				{
					return P_ShapeID_FinishedDate;
				}
				set
				{
					P_ShapeID_FinishedDate = value;
				}
			}
			
			private DateSpan P_DateSpan;
			/// <summary>
			/// 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public DateSpan DateSpan
			{
				get
				{
					return this.P_DateSpan;
				}
				private set
				{
					// 扩展MainForm.TimeSpan的区间
					GlobalApplication.Application.refreshGlobalDateSpan(value);
					this.P_DateSpan = value;
				}
			}
			
#region   ---  开挖平面图的标签信息
			
			private DrawingType P_DrawingType;
			/// <summary>
			/// 此图表所属的类型，由枚举DrawingType提供
			/// </summary>
			/// <value></value>
			/// <returns>DrawingType枚举类型，指示此图形所属的类型</returns>
			/// <remarks></remarks>
public DrawingType Type
			{
				get
				{
					return P_DrawingType;
				}
			}
			
			/// <summary>
			/// 此绘图画面的标签，用来进行索引
			/// </summary>
			/// <remarks></remarks>
			private string P_Name;
			/// <summary>
			/// 此绘图画面的标签，用来进行索引
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string Name
			{
				get
				{
					return this.P_Name;
				}
			}
			
			private long P_UniqueID;
			/// <summary>
			/// 图表的全局独立ID
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks>以当时时间的Tick属性来定义</remarks>
public long UniqueID
			{
				get
				{
					return P_UniqueID;
				}
			}
			
#endregion
			
#region   ---  在开挖平面图中绘制测点
			/// <summary>
			/// 此Visio平面图中是否可以用来绘制监测点位的信息
			/// </summary>
			/// <remarks></remarks>
			private bool P_HasMonitorPointsInfo = false;
			/// <summary>
			/// 此Visio平面图中是否可以用来绘制监测点位的信息
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public bool HasMonitorPointsInfo
			{
				get
				{
					return this.P_HasMonitorPointsInfo;
				}
				set
				{
					this.P_HasMonitorPointsInfo = value;
				}
			}
			
			/// <summary>
			/// Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
			/// </summary>
			/// <remarks></remarks>
			private MonitorPointsInformation P_MonitorPointsInfo;
			/// <summary>
			/// Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public MonitorPointsInformation MonitorPointsInfo
			{
				get
				{
					return this.P_MonitorPointsInfo;
				}
			}
			
#endregion
			
#endregion
			
#region   ---  Fields
			
			/// <summary>
			/// 后台工作线程
			/// </summary>
			/// <remarks></remarks>
			private BackgroundWorker F_BackgroundWorker = new BackgroundWorker();
			
			/// <summary>
			/// 开挖平面图在Visio中的页面名称
			/// </summary>
			/// <remarks></remarks>
			private string F_PageName_PlanView; //= "开挖平面"
			/// <summary>
			/// 所有分区的组合形状的ID值
			/// </summary>
			/// <remarks></remarks>
			private int F_ShapeID_AllRegions; //= 5078
			/// <summary>
			/// 记录开挖信息的文本框
			/// </summary>
			/// <remarks></remarks>
			private int F_InfoBoxID; //= 5079
			//
			/// <summary>
			/// 绘图页面中的“所有形状”的集合
			/// </summary>
			/// <remarks></remarks>
			private Shapes F_shps;
			/// <summary>
			/// 图层：“显示”
			/// </summary>
			/// <remarks></remarks>
			Layer F_layer_Show;
			/// <summary>
			/// 图层：“显示”
			/// </summary>
			/// <remarks></remarks>
			Layer F_layer_Hide;
#endregion
			
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="strFilePath">要打开的Visio文档的绝对路径</param>
			/// <param name="type">Visio平面图所属的绘图类型</param>
			/// <param name="PageName_PlanView">开挖平面图在Visio中的页面名称</param>
			/// <param name="ShapeID_AllRegions">所有分区的组合形状的ID值</param>
			/// <param name="InfoBoxID">记录开挖信息的文本框</param>
			/// <remarks></remarks>
			public ClsDrawing_PlanView(string strFilePath, DrawingType type, string PageName_PlanView, int ShapeID_AllRegions, int InfoBoxID, bool HasMonitorPointsInfo, MonitorPointsInformation MonitorPointsInfo)
				{
				//开挖平面图信息
				this.F_PageName_PlanView = PageName_PlanView;
				this.F_ShapeID_AllRegions = ShapeID_AllRegions;
				this.F_InfoBoxID = InfoBoxID;
				//监测点位信息
				this.P_HasMonitorPointsInfo = HasMonitorPointsInfo;
				this.P_MonitorPointsInfo = MonitorPointsInfo;
				//
				this.P_ShapeID_FinishedDate = GlobalApplication.Application.DataBase.ShapeIDAndFinishedDate;
				
				this.F_BackgroundWorker.WorkerReportsProgress = true;
				this.F_BackgroundWorker.WorkerSupportsCancellation = true;
				//开始绘图
				if (!this.F_BackgroundWorker.IsBusy)
				{
					//在工作线程中执行绘图操作
					this.F_BackgroundWorker.RunWorkerAsync(new[] {strFilePath, type});
				}
				
			}
			
#region   ---  打开Visio平面图
			
			private void F_BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
			{
				//主程序界面的进度条的UI显示
				APPLICATION_MAINFORM.MainForm.ShowProgressBar_Marquee();
				//
				string strFilePath = System.Convert.ToString(e.Argument(0));
				DrawingType type = e.Argument(1);
				//执行具体的绘图操作
				try
				{
					this.Application = NewVsoApp();
				}
				catch (Exception ex)
				{
					Debug.Print(ex.Message);
					e.Cancel = true;
					return;
				}
				
				
				//
				try
				{
					Document vsoDoc = OpenDocument(this.P_Application, strFilePath);
					if (vsoDoc == null)
					{
						e.Cancel = true;
						return;
					}
					//将指定page（开挖平面）的窗口指定给对应变量WndDrawing
					//在NewApplication方法中，已经将指定page对象的窗口进行了激活。
					this.P_Page = vsoDoc.Pages[F_PageName_PlanView];
					//------------设置visioWindow的TimeSpan区间
					//！这里必须将数组进行Clone！以复制一个新副本，不然后面的Sort方法就会改变MainForm.DataBase中的数组的排序。
					DateTime[] TimeRange = P_ShapeID_FinishedDate.FinishedDate.Clone();
					this.DateSpan = this.GetAndRefreshDateSpan(TimeRange);
					//设置此图形所属的类型
					this.P_DrawingType = type;
					//此图形的独立ID
					this.P_UniqueID = GeneralMethods.GetUniqueID();
					this.P_Name = cstDrawingTag;
					//
					// ----------------- 创建索引
					GenerateReference(this.P_Page, ref F_layer_Show, ref F_layer_Hide);
					//---------------- 设置形状的显示与隐藏
					F_shps = this.P_Page.Shapes;
					Shape shpAllRegion = default(Shape);
					shpAllRegion = F_shps.ItemFromID(F_ShapeID_AllRegions);
					SetVisible(this.P_Page, shpAllRegion, F_layer_Show, F_layer_Hide);
				}
				catch (Exception)
				{
					Debug.Print("创建新的Visio对象出错！");
				}
				finally
				{
					//显示窗口与刷新屏幕
					if (this.P_Application != null)
					{
						this.P_Application.Visible = true;
						this.P_Application.ShowChanges = true;
					}
				}
			}
			
			/// <summary>
			/// 创建一个新的Visio.Application对象
			/// </summary>
			/// <returns></returns>
			/// <remarks></remarks>
			private Application NewVsoApp()
			{
				if (this.Application != null)
				{
					return this.P_Application;
				}
				else
				{
					
					//创建新Visio窗口并为其命名
					var vsoApp = new Application();
					
					// Dim vsoApp  = CreateObject("visio.application")
					//但是，这两种方法都有一个问题，就是在刚执行完这一句后，Visio程序的界面就出现了，
					//即使后面紧跟着vsoApp.Visible = False，在UI显示上还是会出现一个界面由出现到隐藏的闪动。
					vsoApp.Visible = false;
					//
					return vsoApp;
				}
			}
			
			/// <summary>
			/// 打开一个Visio文档
			/// </summary>
			/// <param name="vsoApp"></param>
			/// <param name="FilePath">文档的绝对路径</param>
			/// <returns></returns>
			/// <remarks></remarks>
			private Document OpenDocument(Application vsoApp, string FilePath)
			{
				Document vsoDoc = default(Document);
				try
				{
					//有可能出现此visio文档已经在系统中打开的情况
					vsoDoc = vsoApp.Documents.Open(FilePath);
					//Visio界面美化
					VsoApplicationBeauty(vsoApp, vsoDoc);
				}
				catch (Exception)
				{
					vsoApp.Quit();
					this.P_Application = null;
					this.P_Application.BeforeQuit += this.Application_Quit;
					MessageBox.Show("选择的文件已经打开，请将其手动关闭并重新打开。", "Warning", MessageBoxButtons.OK,
						MessageBoxIcon.Warning);
					return null;
				}
				return vsoDoc;
			}
			
			/// <summary>
			/// 获取此文档所对应的施工日期跨度
			/// </summary>
			/// <remarks></remarks>
			private DateSpan GetAndRefreshDateSpan(DateTime[] TimeRange)
			{
				Array.Sort(TimeRange);
				DateSpan DS = new DateSpan();
				DS.StartedDate = TimeRange[0];
				DS.FinishedDate = TimeRange[TimeRange.Length - 1];
				return DS;
			}
			
			/// <summary>
			/// ！创建索引，将变量索引到Visio中的形状集合与相应图层。
			/// </summary>
			/// <param name="DrawingPage"></param>
			/// <param name="layer_Show">页面中用来放置“显示”的对象的图层</param>
			/// <param name="Layer_Hide">页面中用来放置“隐藏”的对象的图层</param>
			/// <remarks></remarks>
			private void GenerateReference(Page DrawingPage, ref Layer layer_Show, ref Layer Layer_Hide)
			{
				//创建“显示”与“隐藏”图层的索引
				//如果文档中已经有相应的图层，则直接索引，如果文档中还没有这两个图层，则添加这两个图层
				const string cstLayerName_Visible = "Show";
				const string cstLayerName_InVisible = "Hide";
				//当要添加的图层已经存在时，会直接返回那个图层，并且保留那个图层所有的设置信息。
				layer_Show = DrawingPage.Layers.Add(cstLayerName_Visible);
				Layer_Hide = DrawingPage.Layers.Add(cstLayerName_InVisible);
				
				//设置“隐藏”与“显示”图层的可见性
				layer_Show.CellsC((System.Int16) VisCellIndices.visLayerVisible).ResultIU = 1;
				Layer_Hide.CellsC((System.Int16) VisCellIndices.visLayerVisible).ResultIU = 0;
			}
			
			/// <summary>
			/// 设置文档中开挖分块形状的显示与隐藏
			/// </summary>
			/// <param name="DrawingPage"></param>
			/// <param name="shpAllRegion">所有开挖分块形状所在的组合形状</param>
			/// <param name="layer_Show"></param>
			/// <param name="Layer_Hide"></param>
			/// <remarks>
			/// 对于文档中的形状的显示与隐藏的处理方法:
			/// 基本原理：如果一个形状同时位于多个图层中，只要这些图层中有一个图层可见，则此形状可见。
			/// 1、先将所有的分块形状都从所有的图层中删除，然后将它们添加到“Show”与“Hide”这两个图层中。
			/// 2、在后续的操作中，如果将形状从“Show”中移除，则此形状会被隐藏，如果将形状添加进“Show”中，则此形状会显示出来；</remarks>
			private void SetVisible(Page DrawingPage, Shape shpAllRegion, Layer layer_Show, Layer Layer_Hide)
			{
				
				Page with_1 = DrawingPage;
				foreach (Layer l in with_1.Layers)
				{
					//如果要移除的形状对象本来就不在该图层中，则此代码不生效，也不会报错。
					l.Remove(shpAllRegion, (short) 0);
				}
				Layer_Hide.Add(shpAllRegion, (short) 0);
				//Layer.Add说明：如果形状为组并且 fPresMems 不为零，则该组的组件形状保留它们当前的图层分配，并且将组件形状添加到该图层。
				//如果 fPresMems 为零，则组件形状将重新分配到该图层，并且丢失它们当前的图层分配。
				layer_Show.Add(shpAllRegion, (short) 1);
				// --------------------------------------------------------
			}
			
			/// <summary>
			/// 进行Visio界面的美化
			/// </summary>
			/// <param name="VsoApp"></param>
			/// <param name="vsoDoc">进行美化时,要先打开Visio文档,如果没有打开任何文档,
			/// 则Application的ActiveDocument与ActiveWindow属性会返回Nothing.</param>
			/// <remarks></remarks>
			private void VsoApplicationBeauty(Application VsoApp, Document vsoDoc)
			{
				
				Application with_1 = VsoApp;
				with_1.Visible = false;
				with_1.ShowChanges = false;
				
				//激活绘图页面
				Page ActivePage = default(Page);
				if (this.F_PageName_PlanView != null)
				{
					//如果visio中还没有打开任何文档，则ActiveDocument返回nothing
					ActivePage = vsoDoc.Pages[F_PageName_PlanView];
				}
				else
				{
					ActivePage = vsoDoc.Pages[1];
				}
				
				//如果visio中还没有打开任何文档，则ActiveWindow返回nothing
				with_1.ActiveWindow.Page = ActivePage;
				this.P_Window = with_1.ActiveWindow;
				this.P_Window.MouseUp += this.DoubleClick_Up;
				//窗口最大化
				with_1.ActiveWindow.WindowState = (System.Int32) VisWindowStates.visWSMaximized;
				
				with_1.ActiveWindow.ShowScrollBars = false;
				with_1.ActiveWindow.ShowRulers = false;
				with_1.ActiveWindow.ShowPageTabs = false;
				//禁止鼠标滚动，但是可以用鼠标滚动+Ctrl进行缩放。
				//.ScrollLock = True
				//让窗口的显示适应页面
				with_1.ActiveWindow.ViewFit = (System.Int32) VisWindowFit.visFitPage;
				//打开“形状”窗格，以供后面用DoCmd进行切换时将其隐藏。
				with_1.ActiveWindow.Windows.ItemFromID((System.Int32) VisWinTypes.visWinIDShapeSearch).Visible = true;
				
				//隐藏状态栏与工具栏
				with_1.ShowStatusBar = false;
				with_1.ShowToolbar = false;
				
				//对于“形状”窗格进行特殊处理——切换显示
				with_1.DoCmd((System.Int16) VisUICmds.visCmdShapesWindow);
				
				//关闭所有的子窗口
				foreach (Window wnd in P_Application.ActiveWindow.Windows)
				{
					wnd.Visible = false;
				}
			}
			
			private void F_BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
			{
				
				if (e.Cancelled || this.Application == null) //说明Visio文档打开异常
				{
					//在绘图完成后，隐藏进度
					APPLICATION_MAINFORM.MainForm.HideProgress("Visio文档打开异常。");
					//启动同步滚动按钮
					APPLICATION_MAINFORM.MainForm.MainUI_RollingObjectCreated();
					
					//---将对象传递给mainform的属性中
					GlobalApplication.Application.PlanView_VisioWindow = null;
					//标记：程序中已经打开的Visio平面图
					// .HasVisioPlanView = False
				}
				else //正常地打开了Visio平面图
				{
					//在绘图完成后，隐藏进度条
					APPLICATION_MAINFORM.MainForm.HideProgress("Done");
					
					//---将对象传递给mainform的属性中
					GlobalApplication.Application.PlanView_VisioWindow = this;
					
				}
			}
#endregion
			
#region   ---  图形滚动
			
			/// <summary>
			/// 图形滚动
			/// </summary>
			/// <param name="dateThisDay"></param>
			/// <remarks>在Visio中的数据记录集上，用来控制形状的显示与否的，就只有“形状ID”与“完成日期”这两列的数据。
			/// Layer.Add说明：如果形状为组并且 fPresMems 不为零，则该组合形状的子形状保留它们当前的图层分配，并且将子形状也添加到该图层。
			///               如果 fPresMems 为零，则子形状将重新分配到该图层，并且丢失它们状原来的图层分配信息，即此时子形状与组合形状的在的图层完全相同。
			/// Layer.Remove说明：如果形状为一个组合，而 fPresMems 为非零值，该组合的成员形状将不会受到影响。
			///                  如果 fPresMems 为零 (0)，该组合的成员形状也将从图层中移除。</remarks>
			public void Rolling(DateTime dateThisDay)
			{
				
				//工作表“开挖分块”中的“形状ID”与“完成日期”两列的数据，其中分别以向量的形式记录了这两列数据。
				//向量的第一个元素的下标值为0()
				int[] arrShapeID = P_ShapeID_FinishedDate.ShapeID;
				DateTime[] arrFinishedDate = P_ShapeID_FinishedDate.FinishedDate;
				//
				Shape shp = default(Shape);
				DateTime FinishedDate = default(DateTime);
				int CompareDate = 0;
				lock(this)
				{
					P_Application.ShowChanges = false;
					P_Application.ScreenUpdating = false;
					for (int IRow = 0; IRow <= (arrShapeID.Length - 1); IRow++)
					{
						//形状的索引是从
						shp = F_shps.ItemFromID(arrShapeID[IRow]);
						FinishedDate = arrFinishedDate[IRow];
						CompareDate = DateTime.Compare(dateThisDay, FinishedDate);
						if (CompareDate >= 0) //说明今天已经过了这个基坑区域完成的日期，所以应该将这个区域显示出来
						{
							this.F_layer_Show.Add(shp, (short) 0);
						}
						else //说明今天还没挖到相应的区域，所以应该将这个区域隐藏起来
						{
							//F_layer_Show.Add(shp, 1)   '将组合形状本身添加进图层，而不添加其子形状。
							F_layer_Show.Remove(shp, (short) 0);
						}
					}
					
					//在记录信息的文本框中显示相关信息
					F_shps.ItemFromID(F_InfoBoxID).Characters.Text = "施工日期：" + dateThisDay.ToString("yyyy/MM/dd");
					
					//刷新屏幕
					P_Application.ShowChanges = true;
				}
			}
			
#endregion
			
			/// <summary>
			/// 绘图界面被删除时引发的事件
			/// </summary>
			/// <param name="app"></param>
			/// <remarks>此事件会在Visio弹出是否要保存文档的对话框之后才被触发，所以不能在此事件中去设置app.AlertResponse。</remarks>
			private void Application_Quit(Application app)
			{
				foreach (Document doc in app.Documents)
				{
					object null_object = null;
					object null_object2 = null;
					object null_object3 = null;
					doc.Close(ref null_object, ref null_object2, ref null_object3);
				}
				//
				this.P_Application = null;
				this.P_Application.BeforeQuit += this.Application_Quit;
				//
				GlobalApplication.Application.PlanView_VisioWindow = null;
				//.HasVisioPlanView = False
			}
			
#region   ---  窗口双击
			
			private DateTime DoubleClick_StartTime;
			private bool DoubleClick_FirstClickInitialized;
			
			private void DoubleClick_Up(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
			{
				if (Button == 1) //Left mouse button released
				{
					if (DoubleClick_FirstClickInitialized)
					{
						int DoubleClickInterval = 300; //定义双击的时间间隔
						int Interval = DateTime.Now.Subtract(DoubleClick_StartTime).Milliseconds;
						if (Interval < DoubleClickInterval)
						{
							if (WindowDoubleClickEvent != null)
								WindowDoubleClickEvent();
							Debug.Print("双击成功！");
							DoubleClick_FirstClickInitialized = false;
						}
						else
						{
							DoubleClick_StartTime = DateTime.Now;
							DoubleClick_FirstClickInitialized = true;
						}
					}
					else
					{
						DoubleClick_StartTime = DateTime.Now;
						DoubleClick_FirstClickInitialized = true;
					}
				}
			}
			
			/// <summary>
			/// 在窗口中双击时让窗口适应页面，即缩放页面使之在窗口中完全显示。
			/// </summary>
			/// <remarks></remarks>
			private void ViewFit()
			{
				this.P_Window.ViewFit = (System.Int32) VisWindowFit.visFitPage;
			}
			
#endregion
			
			/// <summary>
			/// 关闭绘图的Visio文档以及其所在的Application程序
			/// </summary>
			/// <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
			/// <remarks></remarks>
			public void Close(bool SaveChanges = false)
			{
				try
				{
					ClsDrawing_PlanView with_1 = this;
					Document vsoDoc = with_1.Page.Document;
					object null_object = null;
					object null_object2 = null;
					object null_object3 = null;
					vsoDoc.Close(ref null_object, ref null_object2, ref null_object3);
					//此时开挖平面图已经关闭
					with_1.Application.Quit();
				}
				catch (Exception ex)
				{
					MessageBox.Show("关闭Visio开挖平面图出错！" + "\r\n" + ex.Message,
						"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			
		}
	}
}
