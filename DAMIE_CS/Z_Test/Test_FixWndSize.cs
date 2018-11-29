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

//using eZstd.eZAPI.APIWindows;
using eZstd.eZAPI;
using Microsoft.Office.Interop.Excel;

namespace CableStayedBridge
{
	public partial class Test_FixWndSize
	{
		public Test_FixWndSize()
		{
			InitializeComponent();
		}
		private Application ExcelApp;
		public void TestForm_Load(object sender, EventArgs e)
		{
			ExcelApp = new Application();
			ExcelApp.Visible = true;
		}
		public void Button2_Click(object sender, EventArgs e)
		{
			NoResizeExcel(ExcelApp);
		}
		public void Button1_Click(object sender, EventArgs e)
		{
			ResizeExcel(ExcelApp);
		}
		/// <summary>
		/// 禁止Excel程序窗口的缩放
		/// </summary>
		/// <param name="App"></param>
		/// <remarks></remarks>
		public void NoResizeExcel(Application App)
		{
			App.ScreenUpdating = false;
			App.WindowState = XlWindowState.xlNormal;
			Application with_1 = App;
			with_1.Width = 400;
			with_1.Height = 453;
			App.ScreenUpdating = true;
			//Dim hWnd As IntPtr = FindWindow("XLMAIN", App.Caption)
			IntPtr hWnd = App.Hwnd;
			FixWindow(hWnd);
		}
		
		/// <summary>
		/// 允许Excel程序窗口的缩放
		/// </summary>
		/// <param name="App"></param>
		/// <remarks></remarks>
		public void ResizeExcel(Application App)
		{
			App.WindowState = XlWindowState.xlNormal;
			//Dim hWnd As IntPtr = FindWindow("XLMAIN", App.Caption)
			IntPtr hWnd = App.Hwnd;
			unFixWindow(hWnd);
		}
		
		/// <summary>
		/// 通过释放窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
		/// </summary>
		/// <param name="hWnd">要释放大小的窗口的句柄</param>
		/// <remarks></remarks>
		private void FixWindow(IntPtr hWnd)
		{
			int hStyle = eZstd.eZAPI.APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE);
			//禁用最大化的标头及拖拽视窗
			eZstd.eZAPI.APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle & ~WindowStyle.WS_MAXIMIZEBOX & ~WindowStyle.WS_EX_APPWINDOW);
		}
		
		/// <summary>
		/// 通过禁用窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
		/// </summary>
		/// <param name="hWnd">要固定大小的窗口的句柄</param>
		/// <remarks></remarks>
		private void unFixWindow(IntPtr hWnd)
		{
			int hStyle = eZstd.eZAPI.APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE);
			eZstd.eZAPI.APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle | WindowStyle.WS_MAXIMIZEBOX | WindowStyle.WS_EX_APPWINDOW);
		}
		
	}
	
}
