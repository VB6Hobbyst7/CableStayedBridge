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

using System.Configuration;
using CableStayedBridge.All_Drawings_In_Application;

namespace CableStayedBridge
{
	namespace Miscellaneous
	{
		/// <summary>
		/// 与程序的UI显示相关的属性
		/// ApplicationSettingsBase类中可以保存的数据类型：
		/// 1、基本数据类型，如integer、string、single等；
		/// 2、 基本数据类型组成的一维数组（不能是二维或多维数组，但是可以在一维数组中嵌套一维数组，比如以行向量作为列向量的元素来构造二维数组。）；
		/// 3、
		/// ApplicationSettingsBase类中不能保存的数据类型：
		/// 1、泛型Dictionary
		/// 2、</summary>
		/// <remarks></remarks>
		public sealed class mySettings_UI : System.Configuration.ApplicationSettingsBase
		{
			
			/// <summary>
			/// 主程序界面的窗口状态
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting()]public FormWindowState WindowState
			{
				get
				{
					return this["WindowState"];
				}
				set
				{
					this["WindowState"] = value;
				}
			}
			
			/// <summary>
			/// 主程序界面的窗口位置
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting(), DefaultSettingValue("0,0")]public Point WindowLocation
			{
				get
				{
					return this["WindowLocation"];
				}
				set
				{
					this["WindowLocation"] = value;
				}
			}
			
			/// <summary>
			/// 主程序界面的窗口尺寸
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting(), DefaultSettingValue("1000, 650")]public Size WindowSize
			{
				get
				{
					return this["WindowSize"];
				}
				set
				{
					this["WindowSize"] = value;
				}
			}
			
		}
		
		/// <summary>
		/// 与程序的数据相关的属性
		/// ApplicationSettingsBase类中可以保存的数据类型：
		/// 1、基本数据类型，如integer、string、single等；
		/// 2、 基本数据类型组成的一维数组（不能是二维或多维数组，但是可以在一维数组中嵌套一维数组，比如以行向量作为列向量的元素来构造二维数组。）；
		/// 3、
		/// ApplicationSettingsBase类中不能保存的数据类型：
		/// 1、泛型Dictionary
		/// 2、    ''' </summary>
		/// <remarks></remarks>
		public class mySettings_Application : System.Configuration.ApplicationSettingsBase
		{
			
			/// <summary>
			/// 在Visio中的绘制监测点位信息时所需要的参数值
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting()]public ClsDrawing_PlanView.MonitorPointsInformation MonitorPointsInfo
			{
				get
				{
					return this["MonitorPointsInfo"];
				}
				set
				{
					this["MonitorPointsInfo"] = value;
				}
			}
			
			/// <summary>
			/// 是否要强制将Excel单元格中的空字符串转换为Nothing。因为对于一个单元格而言，如果其中只有几个空字符，
			/// 那么将它转换为Object时，它可以会是String类型，而不是Nothing。这样在画图时，它会以数据0.0显示。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting(), DefaultSettingValue("True")]public bool CheckForEmpty
			{
				get
				{
					return this["CheckForEmpty"];
				}
				set
				{
					this["CheckForEmpty"] = value;
				}
			}
			
			/// <summary>
			/// 对于进行滚动的曲线进行批量处理
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
[UserScopedSetting()]public object[] Curve_BatchProcessing
			{
				get
				{
					return this["Curve_BatchProcessing"];
				}
				set
				{
					this["Curve_BatchProcessing"] = value;
				}
			}
			
		}
		
	}
	
	
}
