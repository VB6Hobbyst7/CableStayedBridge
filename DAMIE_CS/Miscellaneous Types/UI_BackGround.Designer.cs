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
	namespace UI
	{
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class UI_BackGround : System.Windows.Forms.Form
		{
			
			//Form overrides dispose to clean up the component list.
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
				System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UI_BackGround));
				this.PictureBoxAME = new System.Windows.Forms.PictureBox();
				this.PictureBoxBackGround = new System.Windows.Forms.PictureBox();
				((System.ComponentModel.ISupportInitialize) this.PictureBoxAME).BeginInit();
				((System.ComponentModel.ISupportInitialize) this.PictureBoxBackGround).BeginInit();
				this.SuspendLayout();
				//
				//PictureBoxAME
				//
				this.PictureBoxAME.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
				this.PictureBoxAME.BackColor = System.Drawing.Color.Transparent;
				this.PictureBoxAME.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
				this.PictureBoxAME.Image = (System.Drawing.Image) (resources.GetObject("PictureBoxAME.Image"));
				this.PictureBoxAME.Location = new System.Drawing.Point(376, 246);
				this.PictureBoxAME.Name = "PictureBoxAME";
				this.PictureBoxAME.Size = new System.Drawing.Size(108, 45);
				this.PictureBoxAME.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
				this.PictureBoxAME.TabIndex = 9;
				this.PictureBoxAME.TabStop = false;
				//
				//PictureBoxBackGround
				//
				this.PictureBoxBackGround.BackColor = System.Drawing.Color.Transparent;
				this.PictureBoxBackGround.BackgroundImage = global::My.Resources.Resources.线条背景;
				this.PictureBoxBackGround.Dock = System.Windows.Forms.DockStyle.Fill;
				this.PictureBoxBackGround.Location = new System.Drawing.Point(0, 0);
				this.PictureBoxBackGround.Name = "PictureBoxBackGround";
				this.PictureBoxBackGround.Size = new System.Drawing.Size(496, 303);
				this.PictureBoxBackGround.TabIndex = 8;
				this.PictureBoxBackGround.TabStop = false;
				//
				//UI_BackGround
				//
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.ClientSize = new System.Drawing.Size(496, 303);
				this.ControlBox = false;
				this.Controls.Add(this.PictureBoxAME);
				this.Controls.Add(this.PictureBoxBackGround);
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.Name = "UI_BackGround";
				this.ShowInTaskbar = false;
				this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
				((System.ComponentModel.ISupportInitialize) this.PictureBoxAME).EndInit();
				((System.ComponentModel.ISupportInitialize) this.PictureBoxBackGround).EndInit();
				this.ResumeLayout(false);
				this.PerformLayout();
				
			}
			internal System.Windows.Forms.PictureBox PictureBoxAME;
			internal System.Windows.Forms.PictureBox PictureBoxBackGround;
			
		}
	}
}
