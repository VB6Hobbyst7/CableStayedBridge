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
	//Namespace UI
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
	partial class UI_SplashScreen : System.Windows.Forms.Form
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
			this.Version = new System.Windows.Forms.Label();
			this.Load += new System.EventHandler(SplashScreen1_Load);
			this.Copyright = new System.Windows.Forms.Label();
			this.SuspendLayout();
			//
			//Version
			//
			this.Version.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.Version.BackColor = System.Drawing.Color.Transparent;
			this.Version.Font = new System.Drawing.Font("Microsoft Sans Serif", (float) (9.0F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(0));
			this.Version.Location = new System.Drawing.Point(252, 511);
			this.Version.Name = "Version";
			this.Version.Size = new System.Drawing.Size(198, 17);
			this.Version.TabIndex = 1;
			this.Version.Text = "Version {0}.{1:00}";
			this.Version.UseWaitCursor = true;
			//
			//Copyright
			//
			this.Copyright.BackColor = System.Drawing.Color.Transparent;
			this.Copyright.Font = new System.Drawing.Font("Microsoft Sans Serif", (float) (9.0F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(0));
			this.Copyright.Location = new System.Drawing.Point(252, 536);
			this.Copyright.Name = "Copyright";
			this.Copyright.Size = new System.Drawing.Size(198, 17);
			this.Copyright.TabIndex = 2;
			this.Copyright.Text = "Copyright";
			this.Copyright.UseWaitCursor = true;
			//
			//UI_SplashScreen
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Control;
			this.BackgroundImage = global::My.Resources.Resources.SplashScreen_DAMIE;
			this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.ClientSize = new System.Drawing.Size(462, 564);
			this.ControlBox = false;
			this.Controls.Add(this.Copyright);
			this.Controls.Add(this.Version);
			this.DoubleBuffered = true;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "UI_SplashScreen";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.TopMost = true;
			this.UseWaitCursor = true;
			this.ResumeLayout(false);
			
		}
		internal System.Windows.Forms.Label Version;
		internal System.Windows.Forms.Label Copyright;
		
	}
	//End Namespace
}
