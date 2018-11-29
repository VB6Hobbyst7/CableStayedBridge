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

using Microsoft.Office.Interop;
using CableStayedBridge.Miscellaneous;
//using DAMIE.All_Drawings_In_Application.ClsDrawing_PlanView;
//using DAMIE.Constants.xmlNodeNames.VisioPlanView_MonitorPoints;
using System.Xml;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.GlobalApp_Form;

namespace CableStayedBridge
{
	/// <summary>
	/// 绘制开挖平面图图的窗口界面
	/// </summary>
	/// <remarks></remarks>
	public partial class frmDrawingPlan
	{
		
#region   ---  Fields
		
		private bool F_HasMonitorPointinfos;
		
		private MonitorPointsInformation MonitorPointinfos;
		/// <summary>
		/// 整个项目的文件路径
		/// </summary>
		/// <remarks></remarks>
		private string ProjectFilePath;
		
#endregion
		
#region   ---  窗口的加载与关闭
		/// <summary>
		/// 构造函数
		/// </summary>
		/// <remarks></remarks>
		public frmDrawingPlan()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			frmDrawingPlan with_1 = this;
			with_1.ChkBx_PointInfo.Checked = false;
			with_1.ChkBx_PointInfo_CheckedChanged(null, null);
			//设置监测点位信息的初始数据
			this.ProjectFilePath = GlobalApplication.Application.ProjectFile.FilePath;
			ImportFromXmlFile(this.ProjectFilePath);
			//
			//Dim settings As New mySettings_Application
			//Call PointsInfoToUI(settings.MonitorPointsInfo)
			
		}
		
		/// <summary>
		/// 窗口关闭
		/// </summary>
		/// <remarks></remarks>
		public void btnCancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		
#endregion
		
		/// <summary>
		/// 打开新的Visio的开挖平面图。如果平面图已经打开，则不能再打开新的平面图。
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void btnChooseVisioPlanView_Click(object sender, EventArgs e)
		{
			string VisioFilepath = string.Empty;
			OpenFileDialog OpenFileDialg = new OpenFileDialog();
			OpenFileDialg.Title = "选择基坑开挖平面图";
			OpenFileDialg.Filter = "Visio Documents  (*.vsd)|*.vsd";
			OpenFileDialg.FilterIndex = 1;
			if (OpenFileDialg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				VisioFilepath = OpenFileDialg.FileName;
			}
			this.TextBoxFilePath.Text = VisioFilepath;
		}
		
		/// <summary>
		/// 生成开挖平面图
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ConstructVisioPlanView(object sender, EventArgs e)
		{
			//检查程序中是否已经有了打开的Visio绘图
			if (GlobalApplication.Application.PlanView_VisioWindow == null)
			{
				string VsoFilepath = this.TextBoxFilePath.Text;
				bool blnFilePathValidated = false;
				if (VsoFilepath.Length > 0)
				{
					if (File.Exists(VsoFilepath))
					{
						if (Path.GetExtension(VsoFilepath) ==".vsd")
						{
							blnFilePathValidated = true;
						}
					}
				}
				if (blnFilePathValidated)
				{
					try
					{
						this.Hide();
						ClsDrawing_PlanView.MonitorPointsInformation PointsInfo = null;
						
						//提取监测点位的信息
						if (F_HasMonitorPointinfos)
						{
							PointsInfo = UIToPointsInfo();
						}
						//提取开挖平面图的信息
						ClsDrawing_PlanView visioWindow = new ClsDrawing_PlanView(strFilePath: ref VsoFilepath, type: DrawingType.Vso_PlanView, PageName_PlanView: this.TextBoxPageName.Text, ShapeID_AllRegions: this.TextBoxAllRegions.Text, InfoBoxID: this.TextBoxInfoBoxID.Text, HasMonitorPointsInfo: ref this.F_HasMonitorPointinfos, MonitorPointsInfo: ref PointsInfo);
						this.Close();
					}
					catch (Exception)
					{
						MessageBox.Show("Visio平面图打开出错，请重新打开。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand);
						this.Visible = true;
						GlobalApplication.Application.PlanView_VisioWindow = null;
					}
				}
				else
				{
					MessageBox.Show("Visio文档不符合规范，请重新选择。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand);
				}
			}
			else
			{
				//不能打开多个Visio平面图
				MessageBox.Show("Visio平面图已经打开。", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Hand);
			}
		}
		
		/// <summary>
		/// 启用或禁用监测点位信息的设置区域
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void ChkBx_PointInfo_CheckedChanged(object sender, EventArgs e)
		{
			switch (this.ChkBx_PointInfo.CheckState)
			{
				case CheckState.Checked:
					this.Panel1.Enabled = true;
					this.F_HasMonitorPointinfos = true;
					break;
				case CheckState.Unchecked:
					this.Panel1.Enabled = false;
					this.F_HasMonitorPointinfos = false;
					break;
			}
		}
		
#region   ---  一般界面操作
		//文本框的字符格式验证
		public void ValidateForSingle(object sender, EventArgs e)
		{
			TextBox ctrl = (TextBox) sender;
			string T = ctrl.Text;
			try
			{
				float Coordinate = float.Parse(T);
				ctrl.Text = T.TrimStart(new[] {'0'});
			}
			catch (Exception)
			{
				ctrl.Text = "0.0";
			}
		}
		public void ValidateForInteger(object sender, EventArgs e)
		{
			TextBox ctrl = (TextBox) sender;
			string T = ctrl.Text;
			try
			{
				int Coordinate = int.Parse(T);
				ctrl.Text = T.TrimStart(new[] {'0'});
			}
			catch (Exception)
			{
				ctrl.Text = "0";
			}
		}
		//数据的导入与导出
		public void Btn_Import_Click(object sender, EventArgs e)
		{
			//这里要再索引一次，是为了避免在此窗口被打开的过程中，更新了项目文件，
			//而如果这里不再次索引，那Me.ProjectFilePath就还是原来的那个未更新的项目文件。
			this.ProjectFilePath = GlobalApplication.Application.ProjectFile.FilePath;
			ImportFromXmlFile(this.ProjectFilePath);
		}
		public void Btn_Export_Click(object sender, EventArgs e)
		{
			XmlDocument xmlDoc = new XmlDocument();
			if (this.ProjectFilePath == null)
			{
				//用打开文件对话框选择 .ame 项目文件
			}
			xmlDoc.Load(this.ProjectFilePath);
			XmlElement eleRoot = (XmlElement) (xmlDoc.SelectSingleNode(System.Convert.ToString(My.Settings.Default.ProjectName)));
			XmlElement ElePtInfo = (XmlElement) (eleRoot.SelectSingleNode(Nd1_MonitorPoints));
			if (ElePtInfo == null)
			{
				ElePtInfo = eleRoot.AppendChild(xmlDoc.CreateElement(Nd1_MonitorPoints));
			}
			//
			this.MonitorPointinfos = UIToPointsInfo();
			ExportToXmlFile(ElePtInfo, this.MonitorPointinfos);
			//
			xmlDoc.Save(this.ProjectFilePath);
			MessageBox.Show("导出到项目文件成功！", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.None);
		}
		//
		/// <summary>
		/// 将监测点位信息的属性值显示在窗口界面中
		/// </summary>
		/// <param name="PointsInfo"></param>
		/// <remarks></remarks>
		private void PointsInfoToUI(ClsDrawing_PlanView.MonitorPointsInformation PointsInfo)
		{
			ClsDrawing_PlanView.MonitorPointsInformation with_1 = PointsInfo;
			this.txtbx_ShapeName_MonitorPointTag.Text = with_1.ShapeName_MonitorPointTag;
			this.txtbx_Pt_BL_ShapeID.Text = System.Convert.ToString(with_1.pt_Visio_BottomLeft_ShapeID);
			this.txtbx_Pt_UR_ShapeID.Text = System.Convert.ToString(with_1.pt_Visio_UpRight_ShapeID);
			this.txtbx_Pt_BL_CAD_X.Text = System.Convert.ToString(with_1.pt_CAD_BottomLeft.X);
			this.txtbx_Pt_BL_CAD_Y.Text = System.Convert.ToString(with_1.pt_CAD_BottomLeft.Y);
			this.txtbx_Pt_UR_CAD_X.Text = System.Convert.ToString(with_1.pt_CAD_UpRight.X);
			this.txtbx_Pt_UR_CAD_Y.Text = System.Convert.ToString(with_1.pt_CAD_UpRight.Y);
		}
		/// <summary>
		/// 根据窗口界面中输入的监测点位数据，来返回对应的结构体属性。
		/// </summary>
		/// <returns></returns>
		/// <remarks></remarks>
		private ClsDrawing_PlanView.MonitorPointsInformation UIToPointsInfo()
		{
			ClsDrawing_PlanView.MonitorPointsInformation PointsInfo = new ClsDrawing_PlanView.MonitorPointsInformation();
			try
			{
				PointsInfo.ShapeName_MonitorPointTag = this.txtbx_ShapeName_MonitorPointTag.Text;
				PointsInfo.pt_CAD_BottomLeft = new PointF(float.Parse(this.txtbx_Pt_BL_CAD_X.Text), float.Parse(this.txtbx_Pt_BL_CAD_Y.Text));
				PointsInfo.pt_CAD_UpRight = new PointF(float.Parse(this.txtbx_Pt_UR_CAD_X.Text), float.Parse(this.txtbx_Pt_UR_CAD_Y.Text));
				PointsInfo.pt_Visio_BottomLeft_ShapeID = int.Parse(this.txtbx_Pt_BL_ShapeID.Text);
				PointsInfo.pt_Visio_UpRight_ShapeID = int.Parse(this.txtbx_Pt_UR_ShapeID.Text);
			}
			catch (Exception)
			{
				MessageBox.Show("测点绘制与定位的数据格式不正确，请调整", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return new ClsDrawing_PlanView.MonitorPointsInformation();
			}
			return PointsInfo;
		}
		
#endregion
		
#region   ---  Visio平面图中，测点的位置与标签信息的导入与导出
		//导入
		/// <summary>
		/// 从项目文件中导入测点的数据
		/// </summary>
		/// <param name="Path_xmlDocFile"></param>
		/// <remarks></remarks>
		private void ImportFromXmlFile(string Path_xmlDocFile)
		{
			XmlDocument xmlDoc = new XmlDocument();
			xmlDoc.Load(Path_xmlDocFile);
			//
			XmlElement eleRoot = (XmlElement) (xmlDoc.SelectSingleNode(System.Convert.ToString(My.Settings.Default.ProjectName)));
			XmlElement ElePtInfo = (XmlElement) (eleRoot.SelectSingleNode(Nd1_MonitorPoints));
			if (ElePtInfo != null)
			{
				this.MonitorPointinfos = ImportFromXmlElement(ElePtInfo);
			}
			else
			{
				MessageBox.Show("项目文件中没有监测点位布置信息的数据");
			}
			//
			PointsInfoToUI(this.MonitorPointinfos);
		}
		/// <summary>
		/// 从xml的节点中导入其子节点中保存的数据
		/// </summary>
		/// <param name="EleParent"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private MonitorPointsInformation ImportFromXmlElement(XmlElement EleParent)
		{
			MonitorPointsInformation PointsInfo = new MonitorPointsInformation();
			try
			{
				PointsInfo.ShapeName_MonitorPointTag = EleParent.SelectSingleNode(Nd2_ShapeName_MonitorPointTag).InnerText;
				PointsInfo.pt_Visio_BottomLeft_ShapeID = (int) (EleParent.SelectSingleNode(Nd2_pt_Visio_BottomLeft_ShapeID).InnerText);
				PointsInfo.pt_Visio_UpRight_ShapeID = (int) (EleParent.SelectSingleNode(Nd2_pt_Visio_UpRight_ShapeID).InnerText);
				PointsInfo.pt_CAD_BottomLeft = new PointF(float.Parse(EleParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_X).InnerText), float.Parse(
					EleParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_Y).InnerText));
				PointsInfo.pt_CAD_UpRight = new PointF(System.Convert.ToSingle(EleParent.SelectSingleNode(Nd2_pt_CAD_UpRight_X).InnerText), System.Convert.ToSingle(
					EleParent.SelectSingleNode(Nd2_pt_CAD_UpRight_Y).InnerText));
			}
			catch (Exception ex)
			{
				MessageBox.Show("从文件导入出错" + "\r\n" + ex.Message + "\r\n" + "报错位置：" +
					ex.TargetSite.Name, "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return PointsInfo;
		}
		//导出
		/// <summary>
		/// 将测点的绘制与定位数据保存到xml文档的某节点中。
		/// </summary>
		/// <param name="xmlParent"></param>
		/// <param name="Pointinfos"></param>
		/// <remarks></remarks>
		private void ExportToXmlFile(XmlElement xmlParent, MonitorPointsInformation Pointinfos)
		{
			try
			{
				XmlDocument Doc = xmlParent.OwnerDocument;
				MonitorPointsInformation with_1 = Pointinfos;
				//
				XmlElement Nd_ShapeName_MonitorPointTag = default(XmlElement);
				XmlElement Nd_pt_CAD_BottomLeft_X = default(XmlElement);
				XmlElement Nd_pt_CAD_BottomLeft_Y = default(XmlElement);
				XmlElement Nd_pt_CAD_UpRight_X = default(XmlElement);
				XmlElement Nd_pt_CAD_UpRight_Y = default(XmlElement);
				XmlElement Nd_pt_Visio_BottomLeft_ShapeID = default(XmlElement);
				XmlElement Nd_pt_Visio_UpRight_ShapeID = default(XmlElement);
				//
				Nd_ShapeName_MonitorPointTag = xmlParent.SelectSingleNode(Nd2_ShapeName_MonitorPointTag);
				bool blnHasChildNode = false;
				if (Nd_ShapeName_MonitorPointTag != null)
				{
					blnHasChildNode = true;
				}
				//后面认为：如果节点中没有子节点“ShapeName_MonitorPointTag”，则没有其他的子节点；
				//而如果有此节点, 则认为其他的子节点也都存在
				if (blnHasChildNode)
				{
					Nd_pt_CAD_BottomLeft_X = xmlParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_X);
					Nd_pt_CAD_BottomLeft_Y = xmlParent.SelectSingleNode(Nd2_pt_CAD_BottomLeft_Y);
					Nd_pt_CAD_UpRight_X = xmlParent.SelectSingleNode(Nd2_pt_CAD_UpRight_X);
					Nd_pt_CAD_UpRight_Y = xmlParent.SelectSingleNode(Nd2_pt_CAD_UpRight_Y);
					Nd_pt_Visio_BottomLeft_ShapeID = xmlParent.SelectSingleNode(Nd2_pt_Visio_BottomLeft_ShapeID);
					Nd_pt_Visio_UpRight_ShapeID = xmlParent.SelectSingleNode(Nd2_pt_Visio_UpRight_ShapeID);
				}
				else
				{
					Nd_ShapeName_MonitorPointTag = xmlParent.AppendChild(Doc.CreateElement(Nd2_ShapeName_MonitorPointTag));
					Nd_pt_CAD_BottomLeft_X = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_BottomLeft_X));
					Nd_pt_CAD_BottomLeft_Y = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_BottomLeft_Y));
					Nd_pt_CAD_UpRight_X = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_UpRight_X));
					Nd_pt_CAD_UpRight_Y = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_CAD_UpRight_Y));
					Nd_pt_Visio_BottomLeft_ShapeID = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_Visio_BottomLeft_ShapeID));
					Nd_pt_Visio_UpRight_ShapeID = xmlParent.AppendChild(Doc.CreateElement(Nd2_pt_Visio_UpRight_ShapeID));
				}
				Nd_ShapeName_MonitorPointTag.InnerText = with_1.ShapeName_MonitorPointTag;
				Nd_pt_CAD_BottomLeft_X.InnerText = System.Convert.ToString(with_1.pt_CAD_BottomLeft.X);
				Nd_pt_CAD_BottomLeft_Y.InnerText = System.Convert.ToString(with_1.pt_CAD_BottomLeft.Y);
				Nd_pt_CAD_UpRight_X.InnerText = System.Convert.ToString(with_1.pt_CAD_UpRight.X);
				Nd_pt_CAD_UpRight_Y.InnerText = System.Convert.ToString(with_1.pt_CAD_UpRight.Y);
				Nd_pt_Visio_BottomLeft_ShapeID.InnerText = System.Convert.ToString(with_1.pt_Visio_BottomLeft_ShapeID);
				Nd_pt_Visio_UpRight_ShapeID.InnerText = System.Convert.ToString(with_1.pt_Visio_UpRight_ShapeID);
				//
			}
			catch (Exception ex)
			{
				MessageBox.Show("导出到文件出错。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" +
					ex.TargetSite.Name, "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		
#endregion
		
	}
}
