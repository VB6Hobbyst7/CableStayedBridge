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

using Microsoft.Office.Interop.Excel;

namespace CableStayedBridge
{
	
	public class Test_Excel : System.Windows.Forms.Form
	{
		public Test_Excel()
		{
			InitializeComponent();
		}
		public Application App;
		internal System.Windows.Forms.Button Button1;
		public Workbook wkbk;
		
		private void Test_Excel_Load(object sender, EventArgs e)
		{
			App = new Application();
			App.Visible = true;
			App.WindowState = XlWindowState.xlMaximized;
			this.wkbk = App.Workbooks.Add();
			test();
			this.Visible = false;
		}
		private void InitializeComponent()
		{
			this.Button1 = new System.Windows.Forms.Button();
			base.Load += new System.EventHandler(Test_Excel_Load);
			this.SuspendLayout();
			//
			//Button1
			//
			this.Button1.Location = new System.Drawing.Point(79, 48);
			this.Button1.Name = "Button1";
			this.Button1.Size = new System.Drawing.Size(75, 23);
			this.Button1.TabIndex = 0;
			this.Button1.Text = "Button1";
			this.Button1.UseVisualStyleBackColor = true;
			//
			//Test_Excel
			//
			this.ClientSize = new System.Drawing.Size(284, 262);
			this.Controls.Add(this.Button1);
			this.Name = "Test_Excel";
			this.ResumeLayout(false);
			
		}
		
		
		public void test()
		{
			Worksheet sht = default(Worksheet);
			sht = wkbk.Worksheets[1];
			Microsoft.Office.Interop.Excel.Chart cht = default(Microsoft.Office.Interop.Excel.Chart);
			cht = sht.Shapes.AddChart(XlChartType.xlXYScatterSmoothNoMarkers).Chart;
			
			double[] X = new double[41];
			double[] Y = new double[41];
			int j = 0;
			float i = 0;
			for (i = 0; i <= 4; i += 0.1F)
			{
				X[j] = i;
				Y[j] = Math.Pow((4 - Math.Pow((i - 2), 2)), 0.5);
				j++;
			}
			Series s;
			s = cht.SeriesCollection().NewSeries();
			s.XValues = X;
			s.Values = Y;
			
			Axis ax1 = default(Axis);
			ax1 = cht.Axes(XlAxisType.xlCategory);
			Axis ax2 = default(Axis);
			ax2 = cht.Axes(XlAxisType.xlValue);
			PlotArea ptArea = default(PlotArea);
			ptArea = cht.PlotArea;
			ax1.MinimumScale = 0;
			ax1.MaximumScale = 5;
			ax1.MajorUnit = 0.5;
			ax1.MinorUnit = 0.1;
			ax1.MajorGridlines.Delete();
			
			ax2.MinimumScale = 0;
			ax2.MaximumScale = 2.5;
			ax2.MajorUnit = 0.5;
			ax2.MinorUnit = 0.1;
			ax2.MajorGridlines.Delete();
			
			
			cht.ChartArea.Top = 0;
			cht.ChartArea.Left = 0;
			cht.ChartArea.Height = 225;
			cht.ChartArea.Width = 425;
			//
			ptArea.InsideTop = 0;
			ptArea.InsideLeft = 0;
			ptArea.InsideHeight = 200;
			ptArea.InsideWidth = 400;
			
		}
	}
	
	
}
