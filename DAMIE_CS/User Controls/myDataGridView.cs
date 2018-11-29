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


namespace CableStayedBridge
{
	namespace AME_UserControl
	{
		/// <summary>
		/// 自定义控件：DataGridView，向其中增加了：插入行、删除行、显示行号等功能
		/// </summary>
		/// <remarks></remarks>
		public class myDataGridView : System.Windows.Forms.DataGridView
		{
			
			private System.ComponentModel.Container components = null;
			private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemInsert;
			private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemRemove;
			private System.Windows.Forms.ContextMenuStrip CMS_DeleteRows;
			private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemRemoveRows;
			private System.Windows.Forms.ContextMenuStrip CMS_RowHeader;
			
#region   ---  控件的初始属性
			
			public myDataGridView()
			{
				InitializeComponent();
			}
			
			[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
			{
				this.components = new System.ComponentModel.Container();
				this.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(myDataGridView_RowsNumberChanged);
				this.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(myDataGridView_RowsNumberChanged);
				this.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(RowsResizable);
				this.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(myDataGridView_RowStateChanged);
				this.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(myDataGridView_RowHeaderMouseClick);
				this.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(myDataGridView_CellMouseClick);
				this.KeyDown += new System.Windows.Forms.KeyEventHandler(myDataGridView_KeyDown);
				this.CMS_RowHeader = new System.Windows.Forms.ContextMenuStrip(this.components);
				this.ToolStripMenuItemInsert = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemInsert.Click += new System.EventHandler(this.InsertRow);
				this.ToolStripMenuItemRemove = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemRemove.Click += new System.EventHandler(this.RemoveRow);
				this.CMS_DeleteRows = new System.Windows.Forms.ContextMenuStrip(this.components);
				this.ToolStripMenuItemRemoveRows = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemRemoveRows.Click += new System.EventHandler(this.ToolStripMenuItemRemoveRows_Click);
				this.CMS_RowHeader.SuspendLayout();
				this.CMS_DeleteRows.SuspendLayout();
				((System.ComponentModel.ISupportInitialize) this).BeginInit();
				this.SuspendLayout();
				//
				//CMS_RowHeader
				//
				this.CMS_RowHeader.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.ToolStripMenuItemInsert, this.ToolStripMenuItemRemove});
				this.CMS_RowHeader.Name = "CMS_RowHeader";
				this.CMS_RowHeader.ShowImageMargin = false;
				this.CMS_RowHeader.Size = new System.Drawing.Size(76, 48);
				//
				//ToolStripMenuItemInsert
				//
				this.ToolStripMenuItemInsert.Name = "ToolStripMenuItemInsert";
				this.ToolStripMenuItemInsert.Size = new System.Drawing.Size(75, 22);
				this.ToolStripMenuItemInsert.Text = "插入";
				//
				//ToolStripMenuItemRemove
				//
				this.ToolStripMenuItemRemove.Name = "ToolStripMenuItemRemove";
				this.ToolStripMenuItemRemove.Size = new System.Drawing.Size(75, 22);
				this.ToolStripMenuItemRemove.Text = "移除";
				//
				//CMS_DeleteRows
				//
				this.CMS_DeleteRows.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.ToolStripMenuItemRemoveRows});
				this.CMS_DeleteRows.Name = "CMS_RowHeader";
				this.CMS_DeleteRows.ShowImageMargin = false;
				this.CMS_DeleteRows.Size = new System.Drawing.Size(112, 26);
				//
				//ToolStripMenuItemRemoveRows
				//
				this.ToolStripMenuItemRemoveRows.Name = "ToolStripMenuItemRemoveRows";
				this.ToolStripMenuItemRemoveRows.Size = new System.Drawing.Size(111, 22);
				this.ToolStripMenuItemRemoveRows.Text = "删除所选行";
				//
				//myDataGridView
				//
				this.ColumnHeadersHeight = 25;
				this.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
				this.RowTemplate.Height = 23;
				this.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
				this.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
				this.Size = new System.Drawing.Size(346, 110);
				this.CMS_RowHeader.ResumeLayout(false);
				this.CMS_DeleteRows.ResumeLayout(false);
				((System.ComponentModel.ISupportInitialize) this).EndInit();
				this.ResumeLayout(false);
				
			}
			
#endregion
			
#region   ---  显示行号
			
			/// <summary>
			/// 行数改变时的事件：显示行号
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void myDataGridView_RowsNumberChanged(object sender, dynamic e)
			{
				long longRow = 0;
				for (longRow = System.Convert.ToInt32(e.RowIndex + e.RowCount) - 1; longRow <= this.Rows.GetLastRow(DataGridViewElementStates.Displayed); longRow++)
				{
					this.Rows[longRow].HeaderCell.Value = (longRow + 1).ToString();
				}
			}
			
			/// <summary>
			/// 设置新添加的一行的Resizable属性为False
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void RowsResizable(object sender, DataGridViewRowsAddedEventArgs e)
			{
				this.Rows[e.RowIndex].Resizable = DataGridViewTriState.False;
			}
			
			private void myDataGridView_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
			{
				e.Row.HeaderCell.Value = (e.Row.Index + 1).ToString();
			}
			
#endregion
			
#region   ---  右键菜单的关联与显示
			
			private void myDataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
			{
				//如果是右击
				if (e.Button == Windows.Forms.MouseButtons.Right)
				{
					if (e.RowIndex != this.Rows.Count)
					{
						//如果行数只有一行
						if (this.Rows.Count <= 1)
						{
							this.ToolStripMenuItemRemove.Enabled = false;
						}
						else
						{
							this.ToolStripMenuItemRemove.Enabled = true;
						}
						//选择右击项的那一行
						this.ClearSelection();
						this.Rows[e.RowIndex].Selected = true;
						//显示菜单栏
						CMS_RowHeader.Show();
						CMS_RowHeader.Left = MousePosition.X;
						CMS_RowHeader.Top = MousePosition.Y;
					}
				}
			}
			
			private void myDataGridView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
			{
				if (e.Button == Windows.Forms.MouseButtons.Right)
				{
					
					int R = e.RowIndex;
					int C = e.ColumnIndex;
					if (R >= 0 & C >= 0)
					{
						//显示菜单栏
						if (this.SelectedRows.Count == 0 || this.Rows.Count < 2)
						{
							this.ToolStripMenuItemRemoveRows.Enabled = false;
						}
						else
						{
							this.ToolStripMenuItemRemoveRows.Enabled = true;
						}
						CMS_DeleteRows.Show();
						CMS_DeleteRows.Left = MousePosition.X;
						CMS_DeleteRows.Top = MousePosition.Y;
					}
				}
				
			}
			
#endregion
			
#region   ---  行的插入与删除
			
			/// <summary>
			/// 插入一行
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void InsertRow(object sender, EventArgs e)
			{
				myDataGridView with_1 = this;
				int SelectedIndex = with_1.SelectedRows[0].Index;
				with_1.Rows.Insert(SelectedIndex, 1);
			}
			
			/// <summary>
			/// 移除一行
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void RemoveRow(object sender, EventArgs e)
			{
				myDataGridView with_1 = this;
				var Row = with_1.SelectedRows[0];
				if (Row.Index < with_1.Rows.Count - 1)
				{
					//当删除最后一行（不带数据，自动添加的行）时会报错：无法删除未提交的新行。
					with_1.Rows.Remove(Row);
				}
			}
			
			/// <summary>
			/// 移除多行
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void ToolStripMenuItemRemoveRows_Click(object sender, EventArgs e)
			{
				myDataGridView with_1 = this;
				//下面的 For Each 是从下往上索引的，即前面的Row对象的index的值大于后面的Index的值
				foreach (DataGridViewRow Row in with_1.SelectedRows)
				{
					if (Row.Index < with_1.Rows.Count - 1)
					{
						//当删除最后一行（不带数据，自动添加的行）时会报错：无法删除未提交的新行。
						with_1.Rows.Remove(Row);
					}
				}
			}
			
#endregion
			
#region   ---  数据的复制与粘贴
			
			/// <summary>
			/// 如下按下Ctrl+V，则将表格中的数据粘贴到DataGridView控件中
			/// </summary>
			/// <remarks>DataGridView表格的索引：行号：表头为-1，第一行为0，列号：表示行编号的列为-1，第一个数据列的列号为0.
			/// DataGridView.Rows.Count与DataGridView.Columns.Count均只计算数据区域，而不包含表头与列头。</remarks>
			private void myDataGridView_KeyDown(object sender, KeyEventArgs e)
			{
				if (e.KeyCode == Keys.Delete)
				{
					// 删除选择的单元格中的数据
					foreach (DataGridViewCell c in this.SelectedCells)
					{
						c.Value = null;
					}
					
				}
				else if (e.Control & e.KeyCode == Keys.V)
				{
					var a = this.SelectedCells;
					var count = a.Count;
					//Dim s As String = "行号" & vbTab & "列号" & vbCrLf
					//For i = 0 To count - 1
					//    s = s & a.Item(i).RowIndex & vbTab & a.Item(i).ColumnIndex & vbCrLf
					//Next
					//MessageBox.Show(s)
					if (count != 1)
					{
						MessageBox.Show("请选择某一个单元格，来作为粘贴的起始位置。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					DataGridViewCell c = this.SelectedCells[0];
					int rownum = c.RowIndex;
					PasteFromTable(c.RowIndex, c.ColumnIndex);
				}
			}
			
			/// <summary> 将表格中的数据粘贴到DataGridView控件中 </summary>
			/// <param name="startRow">粘贴的起始单元格的行位置</param>
			/// <param name="startCol">粘贴的起始单元格的列位置</param>
			/// <remarks>DataGridView表格的索引：行号：表头为-1，第一行为0，列号：表示行编号的列为-1，第一个数据列的列号为0.
			/// DataGridView.Rows.Count与DataGridView.Columns.Count均只计算数据区域，而不包含表头与列头。总行数包括最后一行空数据行。</remarks>
			private void PasteFromTable(int startRow, int startCol)
			{
				string pastTest = Clipboard.GetText();
				if (string.IsNullOrEmpty(pastTest))
				{
					return;
				}
				//excel中是以"空格"和"换行"来当做字段和行，所以用"\r\n"来分隔，即"回车+换行"
				string[] lines = pastTest.Split(new char[] {'\r', '\n'}, StringSplitOptions.RemoveEmptyEntries);
				try
				{
					//For Each line As String In lines
					//    '在每一行的单元格间，作为单元格的分隔的字符为"\t",即水平换行符
					//    Dim strs As String() = line.Split(Chr(9))
					//    Me.Rows.Add(strs)
					//Next
					// MessageBox.Show(startRow & startCol & Me.Rows.Count & Me.Columns.Count)
					//
					int WriteRowsCount = lines.Length; //要写入多少行数据
					int WriteColsCount = lines[0].Split('\t').Length; //要写入的每一行数据中有多少列
					//
					int endRow = startRow + WriteRowsCount - 1; // 要修改的最后一行的行号
					if (endRow > this.Rows.Count - 2) // 说明要额外添加这么多行才能放置要粘贴进来的数据
					{
						this.Rows.Add(endRow + 2 - this.Rows.Count);
					}
					int endCol = 0; // 要修改的最后面的那一列的列号
					endCol = startCol + WriteColsCount <= this.Columns.Count ? startCol + WriteColsCount - 1 : this.Columns.Count - 1;
					//
					string strline = "";
					string[] strs = null;
					for (int r = startRow; r <= endRow; r++)
					{
						strline = lines[r - startRow];
						strs = strline.Split('\t'); //在每一行的单元格间，作为单元格的分隔的字符为"\t",即水平换行符
						for (int c = startCol; c <= endCol; c++)
						{
							this.Rows[r].Cells[c].Value = strs[c - startCol];
						}
					}
				}
				catch (Exception)
				{
					
				}
			}
			
#endregion
			
		}
	}
}
