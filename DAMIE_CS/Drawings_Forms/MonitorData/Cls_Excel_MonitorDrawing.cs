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
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop.Excel;
using CableStayedBridge.Constants;
using eZstd.eZAPI;

namespace CableStayedBridge
{
	namespace All_Drawings_In_Application
	{
		public class Cls_ExcelForMonitorDrawing : Dictionary_AutoKey<Cls_ExcelForMonitorDrawing>.I_Dictionary_AutoKey
		{
			
#region   ---  Properties
			
			private Application P_Excelapp;
public Application Application
			{
				get
				{
					return this.P_Excelapp;
				}
				private set
				{
					if (value != null)
					{
						//不弹出警告对话框
						value.DisplayAlerts = false;
						//获取Excel的进程
						int processId = 0;
						APIWindows.GetWindowThreadProcessId(value.Hwnd, ref processId);
						F_ExcelProcess = Process.GetProcessById(processId);
					}
					else
					{
						F_ExcelProcess = null;
					}
					this.P_Excelapp = value;
					this.P_Excelapp.WorkbookBeforeClose += this.AppQuit;
				}
			}
			
			/// <summary>
			/// 此Excel监测曲线绘图窗口在主程序的集合中的关键字，用来在集合键值对中对此窗口进行索引
			/// </summary>
			/// <remarks></remarks>
			private int P_Key;
			/// <summary>
			/// 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
			/// 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
			/// 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
			/// 在集合中索引此元素：集合.item(me.key)
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public int Key
			{
				get
				{
					return P_Key;
				}
			}
			
			/// <summary>
			/// 在一个监测曲线绘图的Application的一个工作簿中，当前活动的那一个绘图工作表。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public ClsDrawing_Mnt_Base ActiveMntDrawingSheet
			{
				get
				{
					return P_Mnt_Drawings.Last.Value;
				}
			}
			
			//主界面中的所有监测曲线图的集合
			private Dictionary_AutoKey<ClsDrawing_Mnt_Base> P_Mnt_Drawings = new Dictionary_AutoKey<ClsDrawing_Mnt_Base>();
public Dictionary_AutoKey<ClsDrawing_Mnt_Base> Mnt_Drawings
			{
				get
				{
					return P_Mnt_Drawings;
				}
				set
				{
					P_Mnt_Drawings = value;
				}
			}
			
#endregion
			
#region   ---  Fields
			
			/// <summary>
			/// Excel程序的进程对象，用来对此进程进行相差的操作，比如，关闭进程
			/// </summary>
			/// <remarks></remarks>
			private Process F_ExcelProcess;
			
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="Application">进行绘图的Excel绘图程序</param>
			/// <remarks></remarks>
			public Cls_ExcelForMonitorDrawing(Application Application)
			{
				//在集合中以此对象的ID值来索引此对象
				this.Application = Application;
				this.P_Key = System.Convert.ToInt32(GlobalApplication.Application.MntDrawing_ExcelApps.Add(this));
			}
			
			
			/// <summary>
			/// Excel程序关闭时触发的事件
			/// </summary>
			/// <param name="wkbk"></param>
			/// <param name="cancel"></param>
			/// <remarks>在处理过程中不能再执行wkbk.Close方法，不然会多次执行文档的关闭，从而出错。</remarks>
			private void AppQuit(Workbook wkbk, ref bool cancel)
			{
				try
				{
					//方法一：利用进程：退出进程，其操作与与用户使用系统菜单关闭应用程序主窗口的行为一样
					//F_ExcelProcess.CloseMainWindow()
					//方法二:Application.Quit()
					unFixWindow(wkbk.Application.Hwnd);
					wkbk.Application.Quit();
					//但是还有一个问题，上面两种方法，都会出现是否要保存文档更改的弹窗，
					//为了不出现此弹窗而直接关闭程序，可以设置Application.DisplayAlerts = False。
					this.Application = null;
				}
				catch (Exception ex)
				{
					MessageBox.Show("关闭Excel程序出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					this.RemoveFormCollection();
					//刷新滚动窗口的列表框的界面显示
					APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed();
				}
			}
			
			/// <summary>
			/// 将自己从所在的集合中删除
			/// </summary>
			/// <returns></returns>
			/// <remarks></remarks>
			private bool RemoveFormCollection()
			{
				try
				{
					//！！这这里可能会出现执行Excel事件时主程序被初始化——主程序的调用问题
					//所以采用了如下全局共享属性的解决方法
					GlobalApplication.Application.MntDrawing_ExcelApps.Remove(this.Key);
					return true;
				}
				catch (Exception)
				{
					return false;
				}
			}
			
		}
	}
}
