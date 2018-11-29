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
	namespace GlobalApp_Form
	{
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class APPLICATION_MAINFORM : System.Windows.Forms.Form
		{
			
			//Form 重写 Dispose，以清理组件列表。
			[System.Diagnostics.DebuggerNonUserCode()]protected override void Dispose(bool disposing)
			{
				try
				{
					if (disposing && components != null)
					{
						components.Dispose();
					}
				}
				finally
				{
					base.Dispose(disposing);
				}
			}
			
			//Windows 窗体设计器所必需的
			private System.ComponentModel.Container components = null;
			
			//注意: 以下过程是 Windows 窗体设计器所必需的
			//可以使用 Windows 窗体设计器修改它。
			//不要使用代码编辑器修改它。
			[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
			{
				this.components = new System.ComponentModel.Container();
				myTimer.Tick += new System.EventHandler(TimerEventProcessor);
				this.Load += new System.EventHandler(mainForm_Load);
				this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(mainForm_FormClosing);
				this.DragDrop += new System.Windows.Forms.DragEventHandler(APPLICATION_MAINFORM_DragDrop);
				this.DragEnter += new System.Windows.Forms.DragEventHandler(APPLICATION_MAINFORM_DragEnter);
				System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(APPLICATION_MAINFORM));
				this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
				this.BackgroundWorker = new System.ComponentModel.BackgroundWorker();
				this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
				this.StatusStrip1 = new System.Windows.Forms.StatusStrip();
				this.ProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
				this.StatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
				this.NotifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
				this.ToolStrip1 = new System.Windows.Forms.ToolStrip();
				this.TlStrpBtn_Roll = new System.Windows.Forms.ToolStripButton();
				this.TlStrpBtn_Roll.Click += new System.EventHandler(this.StartRolling);
				this.ToolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
				this.MenuStrip1 = new System.Windows.Forms.MenuStrip();
				this.MenuItemFile = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_NewProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_NewProject.Click += new System.EventHandler(this.MenuItem_NewProject_Click);
				this.MenuItem_OpenProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_OpenProject.Click += new System.EventHandler(this.MenuItem_OpenProject_Click);
				this.MenuItem_EditProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_EditProject.Click += new System.EventHandler(this.MenuItem_EditProject_Click);
				this.MenuItem_CloseProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_CloseProject.Click += new System.EventHandler(this.MenuItem_CloseProject_Click);
				this.ToolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
				this.MenuItem_SaveProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_SaveProject.Click += new System.EventHandler(this.MenuItem_SaveProject_Click);
				this.MenuItem_SaveAsProject = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_SaveAsProject.Click += new System.EventHandler(this.MenuItem_SaveAsProject_Click);
				this.ToolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
				this.MenuItemExport = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemExport.Click += new System.EventHandler(this.ExportToWord);
				this.ToolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
				this.MenuItemPreference = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
				this.MenuItemExit = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemExit.Click += new System.EventHandler(this.MenuItemExit_Click);
				this.MenuItemEdit = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemDrawingPoints = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemDrawingPoints.Click += new System.EventHandler(this.MenuItemDrawingPoints_Click);
				this.MenuItemExtractData = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemExtractDataFromExcel = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemExtractDataFromExcel.Click += new System.EventHandler(this.ExtractDataFromExcel);
				this.ToolStripMenuItemExtractDataFromWord = new System.Windows.Forms.ToolStripMenuItem();
				this.ToolStripMenuItemExtractDataFromWord.Click += new System.EventHandler(this.ExtractDataFromWord);
				this.VisioToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
				this.VisioToolStripMenuItem.Click += new System.EventHandler(this.VisioToolStripMenuItem_Click);
				this.MenuItemNew = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemSectionalView = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemSectionalView.Click += new System.EventHandler(this.MenuItemSectionalView_Click);
				this.MenuItemPlanView = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemPlanView.Click += new System.EventHandler(this.ShowForm_DrawingVisioPlanView);
				this.MenuItemMntData_Incline = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemMntData_Incline.Click += new System.EventHandler(this.MenuItemDataMonitored_Click);
				this.MenuItemMntData_Others = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemMntData_Others.Click += new System.EventHandler(this.MenuItemOtherCurves_Click);
				this.MenuItem_Window = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_Arrange_Vertical = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_Arrange_Vertical.Click += new System.EventHandler(this.ChildrenFormAlligment_Vertical);
				this.MenuItem_Arrange_Horizontal = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_Arrange_Horizontal.Click += new System.EventHandler(this.ChildrenFormAlligment_Horizontal);
				this.MenuItem_Arrange_Cascade = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItem_Arrange_Cascade.Click += new System.EventHandler(this.ChildrenFormAlligment_Cascade);
				this.TSMenuItem_Union = new System.Windows.Forms.ToolStripMenuItem();
				this.MenuItemHelp = new System.Windows.Forms.ToolStripMenuItem();
				this.AboutAMEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
				this.StatusStrip1.SuspendLayout();
				this.ToolStrip1.SuspendLayout();
				this.MenuStrip1.SuspendLayout();
				this.SuspendLayout();
				//
				//OpenFileDialog1
				//
				this.OpenFileDialog1.FileName = "OpenFileDialog1";
				//
				//BackgroundWorker
				//
				this.BackgroundWorker.WorkerReportsProgress = true;
				this.BackgroundWorker.WorkerSupportsCancellation = true;
				//
				//StatusStrip1
				//
				this.StatusStrip1.BackColor = System.Drawing.Color.FromArgb(System.Convert.ToInt32(System.Convert.ToByte(88)), System.Convert.ToInt32(System.Convert.ToByte(53)), System.Convert.ToInt32(System.Convert.ToByte(98)));
				this.StatusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.ProgressBar1, this.StatusLabel1});
				this.StatusStrip1.Location = new System.Drawing.Point(0, 540);
				this.StatusStrip1.Name = "StatusStrip1";
				this.StatusStrip1.Size = new System.Drawing.Size(784, 22);
				this.StatusStrip1.TabIndex = 3;
				this.StatusStrip1.Text = "StatusStrip1";
				//
				//ProgressBar1
				//
				this.ProgressBar1.Name = "ProgressBar1";
				this.ProgressBar1.Size = new System.Drawing.Size(250, 16);
				this.ProgressBar1.Visible = false;
				//
				//StatusLabel1
				//
				this.StatusLabel1.ForeColor = System.Drawing.SystemColors.ButtonFace;
				this.StatusLabel1.Name = "StatusLabel1";
				this.StatusLabel1.Size = new System.Drawing.Size(17, 17);
				this.StatusLabel1.Text = "...";
				this.StatusLabel1.Visible = false;
				//
				//NotifyIcon1
				//
				this.NotifyIcon1.Icon = (System.Drawing.Icon) (resources.GetObject("NotifyIcon1.Icon"));
				this.NotifyIcon1.Text = "基坑群实测数据动态分析";
				//
				//ToolStrip1
				//
				this.ToolStrip1.BackColor = System.Drawing.Color.White;
				this.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
				this.ToolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.TlStrpBtn_Roll, this.ToolStripSeparator3});
				this.ToolStrip1.Location = new System.Drawing.Point(0, 25);
				this.ToolStrip1.Name = "ToolStrip1";
				this.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
				this.ToolStrip1.Size = new System.Drawing.Size(784, 25);
				this.ToolStrip1.TabIndex = 5;
				this.ToolStrip1.Text = "ToolStrip1";
				//
				//TlStrpBtn_Roll
				//
				this.TlStrpBtn_Roll.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
				this.TlStrpBtn_Roll.Image = global::My.Resources.Resources.btn_Roll;
				this.TlStrpBtn_Roll.ImageTransparentColor = System.Drawing.Color.Magenta;
				this.TlStrpBtn_Roll.Name = "TlStrpBtn_Roll";
				this.TlStrpBtn_Roll.Size = new System.Drawing.Size(23, 22);
				this.TlStrpBtn_Roll.Text = "ToolStripButton1";
				this.TlStrpBtn_Roll.ToolTipText = "同步滚动";
				//
				//ToolStripSeparator3
				//
				this.ToolStripSeparator3.Name = "ToolStripSeparator3";
				this.ToolStripSeparator3.Size = new System.Drawing.Size(6, 25);
				//
				//MenuStrip1
				//
				this.MenuStrip1.BackColor = System.Drawing.SystemColors.Control;
				this.MenuStrip1.BackgroundImage = global::My.Resources.Resources.菜单栏;
				this.MenuStrip1.GripMargin = new System.Windows.Forms.Padding(0);
				this.MenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItemFile, this.MenuItemEdit, this.MenuItemNew, this.MenuItem_Window, this.TSMenuItem_Union});
				this.MenuStrip1.Location = new System.Drawing.Point(0, 0);
				this.MenuStrip1.Name = "MenuStrip1";
				this.MenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
				this.MenuStrip1.Size = new System.Drawing.Size(784, 25);
				this.MenuStrip1.TabIndex = 1;
				this.MenuStrip1.Text = "主菜单栏";
				//
				//MenuItemFile
				//
				this.MenuItemFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItem_NewProject, this.MenuItem_OpenProject, this.MenuItem_EditProject, this.MenuItem_CloseProject, this.ToolStripSeparator4, this.MenuItem_SaveProject, this.MenuItem_SaveAsProject, this.ToolStripSeparator5, this.MenuItemExport, this.ToolStripSeparator1, this.MenuItemPreference, this.ToolStripSeparator2, this.MenuItemExit});
				this.MenuItemFile.Name = "MenuItemFile";
				this.MenuItemFile.Size = new System.Drawing.Size(58, 21);
				this.MenuItemFile.Text = "文件(&F)";
				//
				//MenuItem_NewProject
				//
				this.MenuItem_NewProject.Name = "MenuItem_NewProject";
				this.MenuItem_NewProject.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N);
				this.MenuItem_NewProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_NewProject.Text = "新建项目(&N)";
				//
				//MenuItem_OpenProject
				//
				this.MenuItem_OpenProject.Name = "MenuItem_OpenProject";
				this.MenuItem_OpenProject.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O);
				this.MenuItem_OpenProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_OpenProject.Text = "打开项目(&O)";
				//
				//MenuItem_EditProject
				//
				this.MenuItem_EditProject.Name = "MenuItem_EditProject";
				this.MenuItem_EditProject.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E);
				this.MenuItem_EditProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_EditProject.Text = "编辑项目";
				//
				//MenuItem_CloseProject
				//
				this.MenuItem_CloseProject.Name = "MenuItem_CloseProject";
				this.MenuItem_CloseProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_CloseProject.Text = "关闭项目";
				//
				//ToolStripSeparator4
				//
				this.ToolStripSeparator4.Name = "ToolStripSeparator4";
				this.ToolStripSeparator4.Size = new System.Drawing.Size(196, 6);
				//
				//MenuItem_SaveProject
				//
				this.MenuItem_SaveProject.Name = "MenuItem_SaveProject";
				this.MenuItem_SaveProject.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S);
				this.MenuItem_SaveProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_SaveProject.Text = "保存(&S)";
				//
				//MenuItem_SaveAsProject
				//
				this.MenuItem_SaveAsProject.Name = "MenuItem_SaveAsProject";
				this.MenuItem_SaveAsProject.ShortcutKeys = (System.Windows.Forms.Keys) ((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
					| System.Windows.Forms.Keys.S);
				this.MenuItem_SaveAsProject.Size = new System.Drawing.Size(199, 22);
				this.MenuItem_SaveAsProject.Text = "另存为...";
				//
				//ToolStripSeparator5
				//
				this.ToolStripSeparator5.Name = "ToolStripSeparator5";
				this.ToolStripSeparator5.Size = new System.Drawing.Size(196, 6);
				//
				//MenuItemExport
				//
				this.MenuItemExport.Name = "MenuItemExport";
				this.MenuItemExport.Size = new System.Drawing.Size(199, 22);
				this.MenuItemExport.Text = "结果输出到Word(&E)...";
				//
				//ToolStripSeparator1
				//
				this.ToolStripSeparator1.Name = "ToolStripSeparator1";
				this.ToolStripSeparator1.Size = new System.Drawing.Size(196, 6);
				//
				//MenuItemPreference
				//
				this.MenuItemPreference.Name = "MenuItemPreference";
				this.MenuItemPreference.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P);
				this.MenuItemPreference.Size = new System.Drawing.Size(199, 22);
				this.MenuItemPreference.Text = "选项(&P)";
				//
				//ToolStripSeparator2
				//
				this.ToolStripSeparator2.Name = "ToolStripSeparator2";
				this.ToolStripSeparator2.Size = new System.Drawing.Size(196, 6);
				//
				//MenuItemExit
				//
				this.MenuItemExit.Name = "MenuItemExit";
				this.MenuItemExit.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4);
				this.MenuItemExit.Size = new System.Drawing.Size(199, 22);
				this.MenuItemExit.Text = "退出(&Q)";
				//
				//MenuItemEdit
				//
				this.MenuItemEdit.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItemDrawingPoints, this.MenuItemExtractData});
				this.MenuItemEdit.Name = "MenuItemEdit";
				this.MenuItemEdit.Size = new System.Drawing.Size(59, 21);
				this.MenuItemEdit.Text = "编辑(&E)";
				//
				//MenuItemDrawingPoints
				//
				this.MenuItemDrawingPoints.Name = "MenuItemDrawingPoints";
				this.MenuItemDrawingPoints.Size = new System.Drawing.Size(139, 22);
				this.MenuItemDrawingPoints.Text = "绘制测点(&P)";
				//
				//MenuItemExtractData
				//
				this.MenuItemExtractData.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.ToolStripMenuItemExtractDataFromExcel, this.ToolStripMenuItemExtractDataFromWord, this.VisioToolStripMenuItem});
				this.MenuItemExtractData.Name = "MenuItemExtractData";
				this.MenuItemExtractData.Size = new System.Drawing.Size(139, 22);
				this.MenuItemExtractData.Text = "数据提取...";
				//
				//ToolStripMenuItemExtractDataFromExcel
				//
				this.ToolStripMenuItemExtractDataFromExcel.Image = global::My.Resources.Resources.DatafromExcel;
				this.ToolStripMenuItemExtractDataFromExcel.Name = "ToolStripMenuItemExtractDataFromExcel";
				this.ToolStripMenuItemExtractDataFromExcel.Size = new System.Drawing.Size(142, 22);
				this.ToolStripMenuItemExtractDataFromExcel.Text = "Excel (&E) ...";
				//
				//ToolStripMenuItemExtractDataFromWord
				//
				this.ToolStripMenuItemExtractDataFromWord.Image = global::My.Resources.Resources.DataFromWord;
				this.ToolStripMenuItemExtractDataFromWord.Name = "ToolStripMenuItemExtractDataFromWord";
				this.ToolStripMenuItemExtractDataFromWord.Size = new System.Drawing.Size(142, 22);
				this.ToolStripMenuItemExtractDataFromWord.Text = "Word(&W) ...";
				//
				//VisioToolStripMenuItem
				//
				this.VisioToolStripMenuItem.Image = global::My.Resources.Resources.IdToShape;
				this.VisioToolStripMenuItem.Name = "VisioToolStripMenuItem";
				this.VisioToolStripMenuItem.Size = new System.Drawing.Size(142, 22);
				this.VisioToolStripMenuItem.Text = "Visio (&V) ...";
				//
				//MenuItemNew
				//
				this.MenuItemNew.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItemSectionalView, this.MenuItemPlanView, this.MenuItemMntData_Incline, this.MenuItemMntData_Others});
				this.MenuItemNew.Name = "MenuItemNew";
				this.MenuItemNew.Size = new System.Drawing.Size(61, 21);
				this.MenuItemNew.Text = "绘图(&D)";
				//
				//MenuItemSectionalView
				//
				this.MenuItemSectionalView.Name = "MenuItemSectionalView";
				this.MenuItemSectionalView.Size = new System.Drawing.Size(178, 22);
				this.MenuItemSectionalView.Text = "开挖剖面图(&S)";
				//
				//MenuItemPlanView
				//
				this.MenuItemPlanView.Name = "MenuItemPlanView";
				this.MenuItemPlanView.Size = new System.Drawing.Size(178, 22);
				this.MenuItemPlanView.Text = "开挖平面图(&P)";
				//
				//MenuItemMntData_Incline
				//
				this.MenuItemMntData_Incline.Name = "MenuItemMntData_Incline";
				this.MenuItemMntData_Incline.Size = new System.Drawing.Size(178, 22);
				this.MenuItemMntData_Incline.Text = "测斜曲线图(&M)";
				//
				//MenuItemMntData_Others
				//
				this.MenuItemMntData_Others.Name = "MenuItemMntData_Others";
				this.MenuItemMntData_Others.Size = new System.Drawing.Size(178, 22);
				this.MenuItemMntData_Others.Text = "其他监测曲线图(&O)";
				//
				//MenuItem_Window
				//
				this.MenuItem_Window.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItem_Arrange_Vertical, this.MenuItem_Arrange_Horizontal, this.MenuItem_Arrange_Cascade});
				this.MenuItem_Window.Name = "MenuItem_Window";
				this.MenuItem_Window.Size = new System.Drawing.Size(64, 21);
				this.MenuItem_Window.Text = "窗口(&W)";
				//
				//MenuItem_Arrange_Vertical
				//
				this.MenuItem_Arrange_Vertical.Name = "MenuItem_Arrange_Vertical";
				this.MenuItem_Arrange_Vertical.Size = new System.Drawing.Size(140, 22);
				this.MenuItem_Arrange_Vertical.Text = "垂直并排(&V)";
				//
				//MenuItem_Arrange_Horizontal
				//
				this.MenuItem_Arrange_Horizontal.Name = "MenuItem_Arrange_Horizontal";
				this.MenuItem_Arrange_Horizontal.Size = new System.Drawing.Size(140, 22);
				this.MenuItem_Arrange_Horizontal.Text = "水平并排(&V)";
				//
				//MenuItem_Arrange_Cascade
				//
				this.MenuItem_Arrange_Cascade.Name = "MenuItem_Arrange_Cascade";
				this.MenuItem_Arrange_Cascade.Size = new System.Drawing.Size(140, 22);
				this.MenuItem_Arrange_Cascade.Text = "层叠(&C)";
				//
				//TSMenuItem_Union
				//
				this.TSMenuItem_Union.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {this.MenuItemHelp, this.AboutAMEToolStripMenuItem});
				this.TSMenuItem_Union.Name = "TSMenuItem_Union";
				this.TSMenuItem_Union.Size = new System.Drawing.Size(61, 21);
				this.TSMenuItem_Union.Text = "帮助(&H)";
				//
				//MenuItemHelp
				//
				this.MenuItemHelp.Name = "MenuItemHelp";
				this.MenuItemHelp.ShortcutKeys = (System.Windows.Forms.Keys) (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F1);
				this.MenuItemHelp.Size = new System.Drawing.Size(184, 22);
				this.MenuItemHelp.Text = "View Help";
				//
				//AboutAMEToolStripMenuItem
				//
				this.AboutAMEToolStripMenuItem.Name = "AboutAMEToolStripMenuItem";
				this.AboutAMEToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
				this.AboutAMEToolStripMenuItem.Text = "About DAMIE";
				//
				//APPLICATION_MAINFORM
				//
				this.AllowDrop = true;
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.BackColor = System.Drawing.SystemColors.Control;
				this.ClientSize = new System.Drawing.Size(784, 562);
				this.Controls.Add(this.ToolStrip1);
				this.Controls.Add(this.StatusStrip1);
				this.Controls.Add(this.MenuStrip1);
				this.Icon = (System.Drawing.Icon) (resources.GetObject("$this.Icon"));
				this.IsMdiContainer = true;
				this.MainMenuStrip = this.MenuStrip1;
				this.Name = "APPLICATION_MAINFORM";
				this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
				this.Text = "DAMIE";
				this.StatusStrip1.ResumeLayout(false);
				this.StatusStrip1.PerformLayout();
				this.ToolStrip1.ResumeLayout(false);
				this.ToolStrip1.PerformLayout();
				this.MenuStrip1.ResumeLayout(false);
				this.MenuStrip1.PerformLayout();
				this.ResumeLayout(false);
				this.PerformLayout();
				
			}
			internal System.Windows.Forms.MenuStrip MenuStrip1;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemFile;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemExit;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemEdit;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_Window;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_Arrange_Horizontal;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_Arrange_Cascade;
			internal System.Windows.Forms.ToolStripMenuItem TSMenuItem_Union;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemExport;
			internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator2;
			internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemNew;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemSectionalView;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemPlanView;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemMntData_Incline;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemMntData_Others;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_Arrange_Vertical;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemHelp;
			internal System.Windows.Forms.ToolStripMenuItem AboutAMEToolStripMenuItem;
			internal System.ComponentModel.BackgroundWorker BackgroundWorker;
			internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator1;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemPreference;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemDrawingPoints;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_NewProject;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_OpenProject;
			internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator4;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_EditProject;
			internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_SaveProject;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_SaveAsProject;
			internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator5;
			internal System.Windows.Forms.StatusStrip StatusStrip1;
			internal System.Windows.Forms.ToolStripProgressBar ProgressBar1;
			internal System.Windows.Forms.ToolStripStatusLabel StatusLabel1;
			internal System.Windows.Forms.NotifyIcon NotifyIcon1;
			internal System.Windows.Forms.ToolStripMenuItem MenuItemExtractData;
			internal System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemExtractDataFromExcel;
			internal System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemExtractDataFromWord;
			internal System.Windows.Forms.ToolStrip ToolStrip1;
			internal System.Windows.Forms.ToolStripButton TlStrpBtn_Roll;
			internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator3;
			internal System.Windows.Forms.ToolStripMenuItem VisioToolStripMenuItem;
			internal System.Windows.Forms.ToolStripMenuItem MenuItem_CloseProject;
		}
	}
}
