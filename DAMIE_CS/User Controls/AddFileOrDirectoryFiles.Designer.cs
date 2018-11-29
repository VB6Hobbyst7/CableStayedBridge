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
		
		
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class AddFileOrDirectoryFiles : System.Windows.Forms.UserControl
		{
			
			//UserControl overrides dispose to clean up the component list.
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
			
			//Required by the Windows Form Designer
			private System.ComponentModel.Container components = null;
			
			//NOTE: The following procedure is required by the Windows Form Designer
			//It can be modified using the Windows Form Designer.
			//Do not modify it using the code editor.
			[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
			{
				this.PanelAddFileOrDir = new System.Windows.Forms.Panel();
				this.MouseLeave += new System.EventHandler(PanelAdd_MouseLeave);
				this.LostFocus += new System.EventHandler(PanelAdd_MouseLeave);
				this.MouseLeave += new System.EventHandler(PanelAdd_MouseLeave);
				this.lbAddFile = new System.Windows.Forms.Label();
				this.lbAddFile.Click += new System.EventHandler(this._AddFile);
				this.lbAddFile.Click += new System.EventHandler(this.PanelAdd_MouseLeave);
				this.lbAddFile.MouseEnter += new System.EventHandler(this.colorFocused);
				this.lbAddFile.MouseLeave += new System.EventHandler(this.colorLostFocus);
				this.lbAddDir = new System.Windows.Forms.Label();
				this.lbAddDir.Click += new System.EventHandler(this._AddDire);
				this.lbAddDir.Click += new System.EventHandler(this.PanelAdd_MouseLeave);
				this.lbAddDir.MouseEnter += new System.EventHandler(this.colorFocused);
				this.lbAddDir.MouseLeave += new System.EventHandler(this.colorLostFocus);
				this.btnAdd = new System.Windows.Forms.Button();
				this.btnAdd.Click += new System.EventHandler(this._AddFile);
				this.btnAdd.Click += new System.EventHandler(this.PanelAdd_MouseLeave);
				this.btnAdd.MouseEnter += new System.EventHandler(this.btnAdd_MouseEnter);
				this.PanelAddFileOrDir.SuspendLayout();
				this.SuspendLayout();
				//
				//PanelAddFileOrDir
				//
				this.PanelAddFileOrDir.Controls.Add(this.lbAddFile);
				this.PanelAddFileOrDir.Controls.Add(this.lbAddDir);
				this.PanelAddFileOrDir.Location = new System.Drawing.Point(2, 24);
				this.PanelAddFileOrDir.Name = "PanelAddFileOrDir";
				this.PanelAddFileOrDir.Size = new System.Drawing.Size(96, 41);
				this.PanelAddFileOrDir.TabIndex = 7;
				this.PanelAddFileOrDir.Visible = false;
				//
				//lbAddFile
				//
				this.lbAddFile.BackColor = System.Drawing.Color.White;
				this.lbAddFile.Font = new System.Drawing.Font("SimSun", (float) (10.5F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(134));
				this.lbAddFile.ForeColor = System.Drawing.SystemColors.InfoText;
				this.lbAddFile.Location = new System.Drawing.Point(0, 0);
				this.lbAddFile.Name = "lbAddFile";
				this.lbAddFile.Size = new System.Drawing.Size(98, 21);
				this.lbAddFile.TabIndex = 0;
				this.lbAddFile.Text = "添加文件";
				this.lbAddFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
				//
				//lbAddDir
				//
				this.lbAddDir.BackColor = System.Drawing.Color.White;
				this.lbAddDir.Font = new System.Drawing.Font("SimSun", (float) (10.5F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(134));
				this.lbAddDir.ForeColor = System.Drawing.SystemColors.InfoText;
				this.lbAddDir.Location = new System.Drawing.Point(0, 20);
				this.lbAddDir.Name = "lbAddDir";
				this.lbAddDir.Size = new System.Drawing.Size(98, 21);
				this.lbAddDir.TabIndex = 0;
				this.lbAddDir.Text = "添加文件夹";
				this.lbAddDir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
				//
				//btnAdd
				//
				this.btnAdd.Location = new System.Drawing.Point(1, 1);
				this.btnAdd.Name = "btnAdd";
				this.btnAdd.Size = new System.Drawing.Size(98, 24);
				this.btnAdd.TabIndex = 6;
				this.btnAdd.Text = "添加";
				this.btnAdd.UseVisualStyleBackColor = true;
				//
				//AddFileOrDirectoryFiles
				//
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.BackColor = System.Drawing.Color.Transparent;
				this.Controls.Add(this.PanelAddFileOrDir);
				this.Controls.Add(this.btnAdd);
				this.Margin = new System.Windows.Forms.Padding(0);
				this.Name = "AddFileOrDirectoryFiles";
				this.Size = new System.Drawing.Size(100, 68);
				this.PanelAddFileOrDir.ResumeLayout(false);
				this.ResumeLayout(false);
				
			}
			internal System.Windows.Forms.Panel PanelAddFileOrDir;
			internal System.Windows.Forms.Label lbAddFile;
			internal System.Windows.Forms.Label lbAddDir;
			internal System.Windows.Forms.Button btnAdd;
			
		}
	}
}
