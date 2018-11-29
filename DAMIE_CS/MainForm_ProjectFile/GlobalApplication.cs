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
using CableStayedBridge.DataBase;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop;


namespace CableStayedBridge
{
	namespace GlobalApp_Form
	{
		
		public class GlobalApplication // GlobalApplication
		{
			
#region   ---  属性值的定义
			
			/// <summary>
			/// 用来索引用来保存全局数据的类实例
			/// </summary>
			private static GlobalApplication F_Application; //GlobalApplication.Application
			/// <summary>
			/// 共享属性，用来索引用来保存全局数据的类实例
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public static GlobalApplication Application
			{
				get
				{
					return F_Application;
				}
			}
			
			
			/// <summary>
			/// 整个程序中用来放置各种隐藏的Excel数据文档的Application对象
			/// </summary>
			/// <remarks></remarks>F:\基坑数据\程序编写\群坑分析\AME\MainForm_ProjectFile\GlobalApplication.vb
			private Microsoft.Office.Interop.Excel.Application F_ExcelApplication_DB;
			/// <summary>
			/// 读取或设置整个程序中用来放置各种隐藏的Excel数据文档的Application对象
			/// </summary>
			/// <value></value>
			/// <returns>一个Excel.Application对象，用来装载整个程序中的所有隐藏的后台数据的Excel文档</returns>
			/// <remarks></remarks>
public Microsoft.Office.Interop.Excel.Application ExcelApplication_DB
			{
				get
				{
					//如果此时还没有打开装载Excel数据库工作簿的Excel程序，则先创建一个Excel程序
					if (this.F_ExcelApplication_DB == null)
					{
						this.F_ExcelApplication_DB = new Microsoft.Office.Interop.Excel.Application();
						this.F_ExcelApplication_DB.Visible = false;
					}
					//
					return this.F_ExcelApplication_DB;
				}
				set
				{
					this.F_ExcelApplication_DB = value;
				}
			}
			
			/// <summary>
			/// 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
			/// </summary>
			/// <remarks></remarks>
			private DateSpan F_DateSpan;
			/// <summary>
			/// 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
			/// 这个全局的DateSpan值是只有扩充，不会缩减的。即在程序中每新添加一个图表，
			/// 就会对其进行扩充，而将图表删除时，这个日期跨度值并不会缩减。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public DateSpan DateSpan
			{
				get
				{
					return F_DateSpan;
				}
			}
			
			
#region   -！-！-  绘图程序界面
			
			/// <summary>
			/// 用来绘制剖面标高图的那一个Excel程序
			/// </summary>
			/// Excel程序界面  ---  标高剖面图
			private ClsDrawing_ExcavationElevation F_ElevationDrawing;
			/// <summary>
			/// 用来绘制剖面标高图的那一个Excel程序
			/// </summary>
public ClsDrawing_ExcavationElevation ElevationDrawing
			{
				get
				{
					return F_ElevationDrawing;
				}
				set
				{
					F_ElevationDrawing = value;
					//刷新滚动窗口的列表框的界面显示
					APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed();
				}
			}
			
			//Excel程序界面  ---  集合  ---  监测曲线绘图
			private Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> F_MntDrawing_ExcelApps = new Dictionary_AutoKey<Cls_ExcelForMonitorDrawing>();
			/// <summary>
			/// Excel程序界面  ---  集合  ---  监测曲线绘图
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> MntDrawing_ExcelApps
			{
				get
				{
					return F_MntDrawing_ExcelApps;
				}
				set
				{
					this.F_MntDrawing_ExcelApps = value;
				}
			}
			
			//Visio程序界面  ---  开挖平面图
			private ClsDrawing_PlanView F_PlanView_VisioWindow;
public ClsDrawing_PlanView PlanView_VisioWindow
			{
				get
				{
					return F_PlanView_VisioWindow;
				}
				set
				{
					F_PlanView_VisioWindow = value;
					//刷新滚动窗口的列表框的界面显示
					APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed();
				}
			}
			
#endregion
			
#region   ---  项目文件与数据库
			
			/// <summary>
			/// 程序中正在运行的那个项目文件
			/// </summary>
			/// <remarks></remarks>
			private clsProjectFile F_ProjectFileContents;
			/// <summary>
			/// 程序中正在运行的那个项目文件
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public clsProjectFile ProjectFile
			{
				get
				{
					if (this.F_ProjectFileContents == null)
					{
						this.F_ProjectFileContents = new clsProjectFile();
					}
					return this.F_ProjectFileContents;
				}
				set
				{
					this.F_ProjectFileContents = value;
				}
			}
			
			/// <summary>
			/// 主程序的数据库文件
			/// </summary>
			private ClsData_DataBase F_DataBase;
			/// <summary>
			/// 主程序的数据库文件
			/// </summary>
			/// <remarks>用来加载所有Excel的数据文档，在整个过程中不可见，即在隐藏的状态下进行数据的交换与存储。</remarks>
public ClsData_DataBase DataBase
			{
				get
				{
					return F_DataBase;
				}
				set
				{
					F_DataBase = value;
				}
			}
			
#endregion
			
			
#endregion
			
#region   ---  主程序的加载与关闭
			
			public GlobalApplication()
			{
				//设置一个初始值blnTimeSpanInitialized，说明MainForm.TimeSpan还没有被初始化过，后期应该先对其赋值以进行初始化，然后再进行比较
				this.blnTimeSpanInitialized = false;
				GlobalApplication.F_Application = this;
			}
			
#endregion
			
#region   ---  遍历程序中的绘图
			
			/// <summary>
			/// 主程序中所有的绘图，包括不能进行滚动的图形。比如开挖平面图，监测曲线图，开挖剖面图。
			/// </summary>
			/// <returns></returns>
			/// <remarks></remarks>
			public AmeDrawings ExposeAllDrawings()
			{
				Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> Excel_Monitor = default(Dictionary_AutoKey<Cls_ExcelForMonitorDrawing>);
				//
				ClsDrawing_PlanView Visio_PlanView = null;
				ClsDrawing_ExcavationElevation Elevation = default(ClsDrawing_ExcavationElevation);
				List<ClsDrawing_Mnt_Base> MonitorData = new List<ClsDrawing_Mnt_Base>();
				//
				Elevation = GlobalApplication.Application.ElevationDrawing;
				Excel_Monitor = GlobalApplication.Application.MntDrawing_ExcelApps;
				//
				
				//Me.F_AllDrawingsCount = 0
				//
				//对象名称与对象
				try
				{
					//剖面图
					Elevation = GlobalApplication.Application.ElevationDrawing;
					if (Elevation != null)
					{
						//Me.F_AllDrawingsCount += 1
					}
				}
				catch (Exception)
				{
					
				}
				
				try
				{
					//平面图
					Visio_PlanView = GlobalApplication.Application.PlanView_VisioWindow;
					if (Visio_PlanView != null)
					{
						//Me.F_AllDrawingsCount += 1
					}
				}
				catch (Exception)
				{
					
				}
				
				try
				{
					//监测曲线图
					foreach (Cls_ExcelForMonitorDrawing MonitorSheets in Excel_Monitor.Values)
					{
						foreach (ClsDrawing_Mnt_Base sht in MonitorSheets.Mnt_Drawings.Values)
						{
							
							MonitorData.Add(sht);
							//Me.F_AllDrawingsCount += 1
						}
					}
				}
				catch (Exception)
				{
				}
				AmeDrawings AllDrawings = new AmeDrawings(Elevation, 
					Visio_PlanView, 
					MonitorData);
				return AllDrawings;
			}
			
			/// <summary>
			/// 提取出主程序中含有Rolling方法的对象
			/// </summary>
			/// <returns>其中的每一个元素都是一个字典，以对象名称来索引对象
			/// 其中有三个元素，依次代表：剖面图、平面图和监测曲线图，
			/// 它不是指从mainform中提取出来的三个App的属性值，而是从这三个App属性中挑选出来的，正确地带有Rolling方法的相关类的实例对象。
			/// </returns>
			/// <remarks></remarks>
			public RollingEnabledDrawings ExposeRollingDrawings()
			{
				//
				ClsDrawing_ExcavationElevation Elevation = default(ClsDrawing_ExcavationElevation);
				ClsDrawing_PlanView Visio_PlanView = default(ClsDrawing_PlanView);
				Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> Excel_Monitor = default(Dictionary_AutoKey<Cls_ExcelForMonitorDrawing>);
				Elevation = GlobalApplication.Application.ElevationDrawing;
				Visio_PlanView = GlobalApplication.Application.PlanView_VisioWindow;
				Excel_Monitor = GlobalApplication.Application.MntDrawing_ExcelApps;
				//
				//对象名称与对象
				try
				{
					//剖面图
					if (Elevation != null)
					{
						//RollingDrawings.SectionalView.Add(Elevation)
					}
				}
				catch (Exception)
				{
					
				}
				
				try
				{
					//平面图
					if (Visio_PlanView != null)
					{
						//RollingDrawings.PlanView.Add(Visio_PlanView)
					}
				}
				catch (Exception)
				{
					
				}
				List<clsDrawing_Mnt_RollingBase> RollingMntDrawings = new List<clsDrawing_Mnt_RollingBase>();
				try
				{
					//监测曲线图
					foreach (Cls_ExcelForMonitorDrawing MonitorSheets in Excel_Monitor.Values)
					{
						foreach (ClsDrawing_Mnt_Base sht in MonitorSheets.Mnt_Drawings.Values)
						{
							if (sht.CanRoll)
							{
								RollingMntDrawings.Add(sht);
							}
						}
					}
					//'数组中的每一个元素都是一个字典，以对象名称来索引对象
				}
				catch (Exception)
				{
				}
				//
				RollingEnabledDrawings RollingDrawings = new RollingEnabledDrawings(Elevation, Visio_PlanView, RollingMntDrawings);
				return RollingDrawings;
			}
			
#endregion
			
			/// <summary>
			/// 布尔值，用来指示主程序的时间跨度是否已被初始化。
			/// </summary>
			/// <remarks>在操作上，如果还没有被初始化，则以添加的形式进行初始化；
			/// 如果已初始化，则是比较并扩展的形式进行初始化</remarks>
			private bool blnTimeSpanInitialized = false;
			/// <summary>
			/// 将主程序的TimeSpan属性与指定的TimeSpan值进行比较，并将主程序的TimeSpan扩展为二者的并集的大区间
			/// </summary>
			/// <param name="ComparedDateSpan">要与主程序的TimeSpan进行比较的TimeSpan值</param>
			/// <remarks></remarks>
			public void refreshGlobalDateSpan(DateSpan ComparedDateSpan)
			{
				GlobalApplication with_1 = this;
				if (!blnTimeSpanInitialized)
				{
					with_1.F_DateSpan = ComparedDateSpan;
					blnTimeSpanInitialized = true;
				}
				else
				{
					with_1.F_DateSpan = GeneralMethods.ExpandDateSpan(with_1.DateSpan, ComparedDateSpan);
				}
			}
			
			/// <summary>
			/// 绘制或删除测点在Visio中的对应形状。
			/// </summary>
			/// <remarks>一个Form对象，用来绘制或删除测点在Visio中的对应形状。</remarks>
			private DiaFrm_PointsTreeView Form_TreeView_MonitorPoints;
			/// <summary>
			/// 打开用于在Visio中绘制监测点位形状的对话框
			/// </summary>
			/// <remarks></remarks>
			public void DrawingPointsInVisio()
			{
				//是否有可以正常打开的，用来绘制监测点位图的窗口对象。
				if (this.F_PlanView_VisioWindow != null) //如果Visio绘图已经打开
				{
					if (this.F_PlanView_VisioWindow.HasMonitorPointsInfo)
					{
						// --------- 判断绘制测点图的窗口是否存在并有效
						bool blnPointFormValidated = true;
						if (this.Form_TreeView_MonitorPoints == null)
						{
							blnPointFormValidated = false;
						}
						else
						{
							if (this.Form_TreeView_MonitorPoints.IsDisposed)
							{
								blnPointFormValidated = false;
							}
						}
						// --------- 判断绘制测点图的窗口是否存在并有效
						
						if (blnPointFormValidated)
						{
							this.Form_TreeView_MonitorPoints.ShowDialog();
						}
						else
						{
							if (this.DataBase != null)
							{
								if (this.DataBase.sheet_Points_Coordinates != null)
								{
									this.Form_TreeView_MonitorPoints = new DiaFrm_PointsTreeView(
										this.DataBase.sheet_Points_Coordinates, 
										this.F_PlanView_VisioWindow);
									this.Form_TreeView_MonitorPoints.ShowDialog();
								}
							}
							else
							{
								MessageBox.Show("没有找到监测点位的数据信息所在的工作表！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							}
						}
					}
					else
					{
						MessageBox.Show("在Visio绘图中没有指定测点信息！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
				}
				else //说明Visio平面图还没有打开
				{
					if (MessageBox.Show("在程序中未检测到Visio平面图，" + "\r\n" + "是否要打开平面图", "Warning", 
						MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) 
						== System.Windows.Forms.DialogResult.OK)
					{
						this.DrawVisioPlanView();
					}
				}
				
			}
			
			/// <summary>
			/// 生成Visio的平面开挖图
			/// </summary>
			/// <remarks></remarks>
			public void DrawVisioPlanView()
			{
				//判断程序中是否有Visio绘图中必须的与每一个分场形状相对应的“完成日期”的数据。
				bool blnHasFinishedDate = false;
				if (this.DataBase != null)
				{
					if (this.DataBase.ShapeIDAndFinishedDate != null)
					{
						APPLICATION_MAINFORM.MainForm.Form_VisioPlanView.ShowDialog();
						blnHasFinishedDate = true;
					}
				}
				if (!blnHasFinishedDate)
				{
					MessageBox.Show("程序中没有与Visio中的形状相对应的\"完成日期\"的数据！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			
			/// <summary>
			/// 程序中所有图表的窗口的句柄值，用来对窗口进行禁用或者启用窗口
			/// </summary>
			/// <param name="AllDrawings"></param>
			/// <returns></returns>
			/// <remarks></remarks>
			public static IntPtr[] GetWindwosHandles(AmeDrawings AllDrawings)
			{
				IntPtr[] arrHandles = new IntPtr[AllDrawings.Count() - 1 + 1];
				try
				{
					int index = 0;
					AmeDrawings with_1 = AllDrawings;
					if (with_1.PlanView != null)
					{
						arrHandles[index] = with_1.PlanView.Application.Window.WindowHandle32;
						index++;
					}
					if (with_1.SectionalView != null)
					{
						arrHandles[index] = with_1.SectionalView.Application.Hwnd;
						index++;
					}
					foreach (ClsDrawing_Mnt_Base mnt in with_1.MonitorData)
					{
						arrHandles[index] = mnt.Application.Hwnd;
						index++;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("获取窗口句柄出错" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return arrHandles;
			}
			
		}
		
		/// <summary>
		/// 主程序中含有Rolling方法,而且可以正确地执行滚动方法的对象
		/// </summary>
		/// <remarks>其中的每一个元素都是一个字典，以对象名称来索引对象
		/// 其中有三个元素，依次代表：剖面图、平面图和监测曲线图，
		/// 它不是指从mainform中提取出来的三个App的属性值，而是从这三个App属性中挑选出来的，正确地带有Rolling方法的相关类的实例对象。</remarks>
		public struct RollingEnabledDrawings
		{
			/// <summary>
			/// 剖面图
			/// </summary>
			/// <remarks></remarks>
			public ClsDrawing_ExcavationElevation SectionalView;
			/// <summary>
			/// 平面图
			/// </summary>
			/// <remarks></remarks>
			public ClsDrawing_PlanView PlanView;
			/// <summary>
			/// 监测曲线图
			/// </summary>
			/// <remarks></remarks>
			public List<clsDrawing_Mnt_RollingBase> MonitorData;
			
			/// <summary>
			/// 主程序中所有可以滚动的图表的数量
			/// </summary>
			/// <returns></returns>
			/// <remarks></remarks>
			public UInt16 Count()
			{
				byte btCount = (byte) 0;
				//
				if (this.SectionalView != null)
				{
					btCount++;
				}
				//
				if (this.PlanView != null)
				{
					btCount++;
				}
				//
				if (this.MonitorData != null)
				{
					btCount += System.Convert.ToByte(this.MonitorData.Count);
				}
				//
				return btCount;
			}
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="SectionalView">剖面图</param>
			/// <param name="PlanView">平面图</param>
			/// <param name="MonitorData">监测曲线图</param>
			/// <remarks></remarks>
			public RollingEnabledDrawings(ClsDrawing_ExcavationElevation SectionalView, 
				ClsDrawing_PlanView PlanView, 
				List<clsDrawing_Mnt_RollingBase> MonitorData)
			{
				this.SectionalView = SectionalView;
				this.PlanView = PlanView;
				this.MonitorData = MonitorData;
			}
			
		}
		
		/// <summary>
		/// 主程序中所有的绘图，包括不能进行滚动的图形。比如开挖平面图，监测曲线图，开挖剖面图。
		/// </summary>
		/// <remarks></remarks>
		public struct AmeDrawings
		{
			/// <summary>
			/// 剖面图
			/// </summary>
			/// <remarks></remarks>
			public ClsDrawing_ExcavationElevation SectionalView;
			/// <summary>
			/// 平面图
			/// </summary>
			/// <remarks></remarks>
			public ClsDrawing_PlanView PlanView;
			/// <summary>
			/// 监测曲线图，其中的每一种监测曲线图的具体类型可以通过其Type属性来进行判断。
			/// </summary>
			/// <remarks></remarks>
			public List<ClsDrawing_Mnt_Base> MonitorData;
			
			/// <summary>
			/// 主程序中所有图表的数量
			/// </summary>
			/// <returns>主程序中所有图表的数量</returns>
			/// <remarks></remarks>
			public UInt16 Count()
			{
				
				byte btCount = (byte) 0;
				//
				if (this.SectionalView != null)
				{
					btCount++;
				}
				//
				if (this.PlanView != null)
				{
					btCount++;
				}
				//
				if (this.MonitorData != null)
				{
					btCount += System.Convert.ToByte(this.MonitorData.Count);
				}
				//
				return btCount;
			}
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="SectionalView">剖面图</param>
			/// <param name="PlanView">平面图</param>
			/// <param name="MonitorData">监测曲线图</param>
			/// <remarks></remarks>
			public AmeDrawings(ClsDrawing_ExcavationElevation SectionalView, 
				ClsDrawing_PlanView PlanView, 
				List<ClsDrawing_Mnt_Base> MonitorData)
			{
				this.SectionalView = SectionalView;
				this.PlanView = PlanView;
				this.MonitorData = MonitorData;
			}
			
		}
		
	}
	
}
