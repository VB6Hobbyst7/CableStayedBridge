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

namespace CableStayedBridge
{
	
	/// <summary>
	/// 对于监测曲线进行批量操作：锁定或者删除
	/// </summary>
	/// <remarks></remarks>
	public partial class frmLockCurves
	{
		
#region    ---   窗口的加载与关闭
		public frmLockCurves()
		{
			// This call is required by the designer.
			InitializeComponent();
			// Add any initialization after the InitializeComponent() call.
			
			
			//从存储文件中提取数据
			mySettings_Application mySetting = new mySettings_Application();
			object[] Date_Handle = null;
			Date_Handle = mySetting.Curve_BatchProcessing;
			if (Date_Handle != null)
			{
				// 将数据刷新到界面的列表中
				int rowscount = Date_Handle.Length;
				this.MyDataGridView1.Rows.Add(rowscount); //必须要先添加行，然后才能在后面进行赋值，否则会出现索引不在集合内的报错。
				//注意：第一行数据的行标为0，列表头所在行的下标为-1。
				for (int rowNum = 0; rowNum <= rowscount - 1; rowNum++)
				{
					this.MyDataGridView1[0, rowNum].Value = Date_Handle[rowNum][0];
					this.MyDataGridView1[1, rowNum].Value = Date_Handle[rowNum][1];
				}
			}
		}
		
		/// <summary>
		/// 清空列表中的所有数据
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnClear_Click(object sender, EventArgs e)
		{
			this.MyDataGridView1.Rows.Clear();
		}
		
		/// <summary>
		/// 关闭窗口
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btn_Cancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		
		
#endregion
		
#region    ---   数据验证
		
		/// <summary>
		/// 在DataGridView中，添加新行时，将其搜索方向设置为“锁定”。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void MyDataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
		{
			if (e.RowIndex >= 1)
			{
				var handle = with_1.Item(1, e.RowIndex - 1);
				if (Handle.Value == null)
				{
					Handle.Value = "锁定";
				}
			}
		}
		
		/// <summary>
		/// 校验单元格中的日期数据
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void MyDataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex >= 0 & e.ColumnIndex == 0)
			{
				DataGridViewCell cell = this.MyDataGridView1[e.ColumnIndex, e.RowIndex];
				string v = System.Convert.ToString(cell.Value);
				DateTime dt = default(DateTime);
				try //将单元格中的内容转换为日期
				{
					dt = System.Convert.ToDateTime(v);
					cell.Value = dt.ToShortDateString();
				}
				catch (Exception)
				{
					short l = (short) v.Length;
					try
					{
						switch (l)
						{
							case (short) 4:
								dt = new DateTime(DateAndTime.Today.Year, int.Parse(v.Substring(2, 2)), int.Parse(v.Substring(4, 2)));
								cell.Value = dt.ToShortDateString();
								break;
							case (short) 6:
								dt = new DateTime((int.Parse(v.Substring(0, 2))) + 2000, int.Parse(v.Substring(2, 2)), int.Parse(v.Substring(4, 2)));
								cell.Value = dt.ToShortDateString();
								break;
							case (short) 8:
								dt = new DateTime(int.Parse(v.Substring(0, 4)), int.Parse(v.Substring(4, 2)), int.Parse(v.Substring(6, 2)));
								cell.Value = dt.ToShortDateString();
								break;
							default:
								cell.Value = "";
								break;
						}
					}
					catch (Exception)
					{
						cell.Value = "";
					}
				}
			}
		}
		
		
#endregion
		
#region    ---   执行操作
		/// <summary>
		/// 执行
		/// </summary>
		/// <remarks></remarks>
		public void btn_OK_Click(object sender, EventArgs e)
		{
			var a = this.MyDataGridView1.DataSource;
			
			// 提取frmRolling中已经选择的要进行滚动的监测曲线图对象
			frmRolling RollingForm = APPLICATION_MAINFORM.MainForm.Form_Rolling;
			List<clsDrawing_Mnt_RollingBase> RollingMnts = RollingForm.F_SelectedDrawings.RollingMnt;
			// 提到表格中的数据
			int rowscount = this.MyDataGridView1.Rows.Count;
			object[] Date_Handle = new object[rowscount - 2 + 1];
			for (int rowNum = 0; rowNum <= rowscount - 2; rowNum++)
			{
				Date_Handle[rowNum] = new[] {this.MyDataGridView1[0, rowNum].Value, this.MyDataGridView1[1, rowNum].Value};
			}
			//将结果保存到文件中
			mySettings_Application mySetting = new mySettings_Application();
			mySetting.Curve_BatchProcessing = Date_Handle; //将结果保存到文件中
			mySetting.Save();
			//
			//为每一个Excel图形创建一个线程
			foreach (clsDrawing_Mnt_RollingBase RollingMnt in RollingMnts)
			{
				Thread thd = new Thread(new System.Threading.ThreadStart(LockOrDelete));
				thd.Name = RollingMnt.Chart_App_Title;
				thd.Start(new[] {RollingMnt, Date_Handle});
			}
		}
		
		/// <summary>
		/// 在工作线程中进行曲线的锁定或者删除
		/// </summary>
		/// <param name="Arg"></param>
		/// <remarks></remarks>
		private void LockOrDelete(object Arg)
		{
			clsDrawing_Mnt_RollingBase RollingMnt = Arg(0);
			object[] Date_Handle = Arg(1);
			for (int row = 0; row <= (Date_Handle.Length - 1); row++)
			{
				DateTime RollingDate = System.Convert.ToDateTime(Date_Handle[row][0]);
				string handle = System.Convert.ToString(Date_Handle[row][1]);
				if (handle == "锁定")
				{
					RollingAndLock(RollingMnt, RollingDate);
				}
				else
				{
					DeleteCurve(RollingMnt, RollingDate);
				}
			}
		}
		
		/// <summary>
		/// 图形滚动，并锁定曲线
		/// </summary>
		/// <param name="RollingMnt">要进行滚动的曲线所在的监测图</param>
		/// <param name="RollingDate">要进行滚动的的曲线所代表的日期。</param>
		/// <remarks></remarks>
		private void RollingAndLock(clsDrawing_Mnt_RollingBase RollingMnt, DateTime RollingDate)
		{
			//首先，要绘制的日期必须在此测点的监测数据的日期跨度范围之内。
			if (RollingMnt.DateSpan.Contains(RollingDate))
			{
				// 检查指定的日期是否早就已经绘制在了图表中
				var a = RollingMnt.Dic_SeriesIndex_Tag;
				bool blnMatched = false; //如果图形中已经有了这条曲线，就不添加或者锁定了。
				int SeriesIndex = 0;
				var Indices = a.Keys;
				foreach (int tempLoopVar_SeriesIndex in Indices)
				{
					SeriesIndex = tempLoopVar_SeriesIndex;
					if (DateTime.Compare(System.Convert.ToDateTime(a.Item(SeriesIndex).ConstructionDate), RollingDate) == 0)
					{
						blnMatched = true;
						break;
					}
				}
				// 如果指定的日期与图表中的第一条曲线的日期相匹配，则默认为不进行绘制
				if (!blnMatched)
				{
					RollingMnt.Rolling(RollingDate);
					//进行滚动的曲线总是seriesCollection中的第一条曲线，其下标值为1。
					RollingMnt.CopySeries(1);
				}
			}
		}
		
		/// <summary>
		/// 删除图形中的曲线，如果图形中的没有对应日期的曲线，就不删除
		/// </summary>
		/// <param name="RollingMnt">要进行滚动的曲线所在的监测图</param>
		/// <param name="RollingDate">要进行滚动的的曲线所代表的日期。</param>
		/// <remarks></remarks>
		private void DeleteCurve(clsDrawing_Mnt_RollingBase RollingMnt, DateTime RollingDate)
		{
			var a = RollingMnt.Dic_SeriesIndex_Tag;
			short CurveCount = System.Convert.ToInt16(a.Count);
			bool blnMatched = false; //只删除图形中已经绘制的曲线
			int SeriesIndex = 0;
			// 检查图形中的曲线所对应的日期是否有与要进行删除的日期相匹配的。
			var Indices = a.Keys;
			foreach (int tempLoopVar_SeriesIndex in Indices)
			{
				SeriesIndex = tempLoopVar_SeriesIndex;
				if (DateTime.Compare(System.Convert.ToDateTime(a.Item(SeriesIndex).ConstructionDate), RollingDate) == 0)
				{
					if (SeriesIndex != 1)
					{
						//删除曲线，其中第一条曲线是用来进行滚动的，不能删除，其下标值为1。
						RollingMnt.DeleteSeries(SeriesIndex);
						break;
					}
				}
			}
		}
#endregion
		
	}
}
