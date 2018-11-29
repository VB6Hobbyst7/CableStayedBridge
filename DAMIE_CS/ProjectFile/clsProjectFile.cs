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

using System.Xml;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
using Microsoft.Office.Interop.Excel;
using CableStayedBridge.Constants;
//using DAMIE.Constants.xmlNodeNames;


namespace CableStayedBridge
{
	namespace DataBase
	{
		
		/// <summary>
		/// 项目文件类，对应于每一个本地的项目文件。
		/// 它主要实现XML文档中的内容与程序中的FileContents对象的交互。
		/// </summary>
		/// <remarks>此类并不实现与界面的UI交互。</remarks>
		public class clsProjectFile
		{
			
#region   ---  属性值定义
			
			/// <summary>
			/// 项目文件的路径
			/// </summary>
			/// <remarks></remarks>
			private string P_FilePath;
			/// <summary>
			/// 项目文件的路径
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public string FilePath
			{
				get
				{
					return this.P_FilePath;
				}
				set
				{
					this.P_FilePath = value;
				}
			}
			
			/// <summary>
			/// 项目文件中记录的内容的实际对象
			/// </summary>
			/// <remarks></remarks>
			private clsData_FileContents P_FileContents;
			/// <summary>
			/// ！项目文件中记录的内容的实际对象，即其中的Workbook对象、Worksheet对象等
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public clsData_FileContents Contents
			{
				get
				{
					return this.P_FileContents;
				}
				set
				{
					this.P_FileContents = value;
				}
				
			}
			
#endregion
			
#region   ---  字段值定义
			
			/// <summary>
			/// 整个程序中用来放置各种隐藏的Excel数据文档的Application对象
			/// </summary>
			/// <remarks></remarks>
			private Application F_Application;
			
			/// <summary>
			/// 指示此项目文件是否是有效文件，即文件中的数据是否正常，文件中索引的工作簿或者工作表是否正常
			/// </summary>
			/// <remarks>只要有一个不正常，则为False</remarks>
			private bool F_blnFileValid;
			
			private List<string> F_lstErrorMessage = new List<string>();
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="xmlFilePath">与程序进行交互的XML文档的路径，如果不指定，则为空字符</param>
			/// <remarks></remarks>
			public clsProjectFile(string xmlFilePath = "")
			{
				this.P_FilePath = xmlFilePath;
				this.F_Application = GlobalApplication.Application.ExcelApplication_DB;
			}
			
			//将项目中的内容写入XML文档
			/// <summary>
			/// 将设置好的项目写入文件
			/// </summary>
			/// <remarks></remarks>
			public void SaveToXmlFile()
			{
				clsData_FileContents FileContents = this.P_FileContents;
				XmlDocument xmlDoc = new XmlDocument();
				Worksheet sheet = default(Worksheet);
				// --- 写入根节点
				XmlNode eleRoot = xmlDoc.CreateElement(System.Convert.ToString(My.Settings.Default.ProjectName));
				xmlDoc.AppendChild(eleRoot);
				XmlElement eleDataBase = xmlDoc.CreateElement(DataBasePath.Nd1_DataBasePaths);
				eleRoot.AppendChild(eleDataBase);
				// ------------ 写入整个项目的所有工作簿的绝对路径
				foreach (Workbook wkbk in FileContents.lstWkbks)
				{
					XmlElement eleWkbks = xmlDoc.CreateElement(DataBasePath.Nd2_WorkbooksInProject);
					eleWkbks.InnerText = wkbk.FullName;
					eleDataBase.AppendChild(eleWkbks);
				}
				
				//--------- 写入施工进度工作表
				short iProgress = (short) 1;
				foreach (Worksheet tempLoopVar_sheet in FileContents.lstSheets_Progress)
				{
					sheet = tempLoopVar_sheet;
					XmlElement eleProgress = xmlDoc.CreateElement(DataBasePath.Nd2_Progress);
					WriteChildNodes(xmlDoc, eleProgress, sheet);
					eleDataBase.AppendChild(eleProgress);
				}
				
				//--------- 写入开挖剖面工作表
				XmlElement eleSecional = xmlDoc.CreateElement(DataBasePath.Nd2_SectionalView);
				eleDataBase.AppendChild(eleSecional);
				sheet = FileContents.Sheet_Elevation;
				WriteChildNodes(xmlDoc, eleSecional, sheet);
				
				//-------- 写入测点坐标工作表
				XmlElement elePoint = xmlDoc.CreateElement(DataBasePath.Nd2_PointCoordinates);
				eleDataBase.AppendChild(elePoint);
				sheet = FileContents.Sheet_PointCoordinates;
				WriteChildNodes(xmlDoc, elePoint, sheet);
				
				
				//-------- 写入测点坐标工作表
				XmlElement eleWorkingStage = xmlDoc.CreateElement(DataBasePath.Nd2_WorkingStage);
				eleDataBase.AppendChild(eleWorkingStage);
				sheet = FileContents.Sheet_WorkingStage;
				WriteChildNodes(xmlDoc, eleWorkingStage, sheet);
				
				
				//-------- 写入开挖分块平面图
				XmlElement elePlan = xmlDoc.CreateElement(DataBasePath.Nd2_PlanView);
				eleDataBase.AppendChild(elePlan);
				sheet = FileContents.Sheet_PlanView;
				WriteChildNodes(xmlDoc, elePlan, sheet);
				
				//保存文档
				xmlDoc.Save(this.P_FilePath);
			}
			/// <summary>
			/// 将每一个工作表项目写入XML文档中，此方法在ParentElement下创建两个子节点
			/// </summary>
			/// <param name="xmlDoc">写入节点的xml文档</param>
			/// <param name="ParentElement">节点元素，要写入的子节点就是在此节点之下的</param>
			/// <param name="sheet">要写入的Excel工作表</param>
			/// <remarks>在此方法中，将指定工作表所在的工作簿的绝对路径，与此工作表的名称，作为两个子节点，
			/// 写入到父节点ParentElement中。</remarks>
			private void WriteChildNodes(XmlDocument xmlDoc, XmlElement ParentElement, Worksheet sheet)
			{
				if (sheet != null)
				{
					XmlDocument with_1 = xmlDoc;
					
					//节点：工作簿路径
					Workbook wkbk = default(Workbook);
					XmlElement eleFilePath1 = with_1.CreateElement(DataBasePath.Nd3_FilePath);
					wkbk = sheet.Parent;
					eleFilePath1.InnerText = wkbk.FullName;
					
					//节点：工作表名称
					XmlElement eleShtName = with_1.CreateElement(DataBasePath.Nd3_SheetName);
					eleShtName.InnerText = sheet.Name;
					
					//文件写入
					ParentElement.AppendChild(eleFilePath1);
					ParentElement.AppendChild(eleShtName);
				}
			}
			//从XML文件读取并检测文件中的成员是否存在
			/// <summary>
			/// 从项目文件中读取数据，并打开相应的Excel程序与工作簿
			/// </summary>
			/// <remarks></remarks>
			public void LoadFromXmlFile()
			{
				//载入文档
				XmlDocument xmlDoc = new XmlDocument();
				xmlDoc.Load(this.P_FilePath);
				//
				clsData_FileContents FC = new clsData_FileContents();
				XMLNode eleRoot = xmlDoc.SelectSingleNode(System.Convert.ToString(My.Settings.Default.ProjectName));
				//这里可以尝试用GetElementById
				XmlElement Node_DataBase = eleRoot.SelectSingleNode(DataBasePath.Nd1_DataBasePaths);
				if (Node_DataBase == null)
				{
					return;
				}
				// ---------------------- 读取文档 ------------------------
				// ---------------------- 读取文档 ------------------------
				
				XmlNodeList eleWkbks = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_WorkbooksInProject);
				foreach (XmlElement eleWkbk in eleWkbks)
				{
					string strWkbkPath = eleWkbk.InnerText;
					Workbook wkbk = ExcelFunction.MatchOpenedWkbk(wkbkPath: ref strWkbkPath, Application: ref this.F_Application, OpenIfNotOpened: true);
					if (wkbk != null)
					{
						FC.lstWkbks.Add(wkbk);
					}
					else //此工作簿不存在，或者是没有成功赋值
					{
						this.F_blnFileValid = false;
						this.F_lstErrorMessage.Add("The Specified Workbook is not found : " + strWkbkPath);
					}
				}
				
				// ---------------- 施工进度工作表
				bool blnNodeForWorksheetValidated = false;
				XmlNodeList eleSheetsProgress = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_Progress);
				foreach (XmlElement eleSheetProgress in eleSheetsProgress)
				{
					Worksheet shtProgress = ValidateNodeForWorksheet(eleSheetProgress, FC, ref blnNodeForWorksheetValidated);
					if (blnNodeForWorksheetValidated)
					{
						FC.lstSheets_Progress.Add(shtProgress);
					}
				}
				
				// ---------------- 开挖平面图工作表
				
				XmlNodeList eleSheetPlanView = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_PlanView);
				var shtPlanView = ValidateNodeForWorksheet((System.Xml.XmlElement) (eleSheetPlanView.Item(0)), FC, ref blnNodeForWorksheetValidated);
				FC.Sheet_PlanView = shtPlanView;
				
				// ---------------- 开挖剖面图工作表
				
				XmlNodeList eleSheetSectionalView = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_SectionalView);
				Worksheet shtSectionalView = ValidateNodeForWorksheet((System.Xml.XmlElement) (eleSheetSectionalView.Item(0)), FC, ref blnNodeForWorksheetValidated);
				FC.Sheet_Elevation = shtSectionalView;
				
				// ---------------- 测点坐标工作表
				
				XmlNodeList eleSheetPointCoordinates = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_PointCoordinates);
				Worksheet shtPoint = ValidateNodeForWorksheet((System.Xml.XmlElement) (eleSheetPointCoordinates.Item(0)), FC, ref blnNodeForWorksheetValidated);
				FC.Sheet_PointCoordinates = shtPoint;
				
				// ---------------- 开挖工况工作表
				
				XmlNodeList eleWorkingStage = Node_DataBase.GetElementsByTagName(DataBasePath.Nd2_WorkingStage);
				Worksheet shtWorkingStage = ValidateNodeForWorksheet((System.Xml.XmlElement) (eleWorkingStage.Item(0)), FC, ref blnNodeForWorksheetValidated);
				FC.Sheet_WorkingStage = shtWorkingStage;
				//
				this.P_FileContents = FC;
				//刷新主程序界面显示
				APPLICATION_MAINFORM.MainForm.MainUI_ProjectOpened();
			}
			/// <summary>
			/// 检测工作表节点的有效性，此类节点中包含了两个子节点，一个是此工作表所在的工作簿的路径，一个是此工作表的名称；
			/// 如果检测通过，则返回此Worksheet对象，否则返回Nothing
			/// </summary>
			/// <param name="WorksheetNode"></param>
			/// <param name="FileContents">用来放置项目文件中记录的工作簿或者工作表对象的变量</param>
			/// <param name="blnNodeForWorksheetValidated"></param>
			/// <returns>要返回的工作表对象，如果验证不通过，则返回Nothing</returns>
			/// <remarks></remarks>
			private Worksheet ValidateNodeForWorksheet(XmlElement WorksheetNode, clsData_FileContents FileContents, ref bool blnNodeForWorksheetValidated)
			{
				//要返回的工作表对象，如果验证不通过，则返回Nothing
				Worksheet ValidSheet = null;
				//
				blnNodeForWorksheetValidated = false;
				//节点中记录的工作簿路径
				string strWkbkPath = "";
				XmlNode ndWkbkPath = WorksheetNode.SelectSingleNode(DataBasePath.Nd3_FilePath);
				if (ndWkbkPath == null) //说明此节点中没有记录工作表所在的工作簿信息，也就是说，此节点中没有记录值
				{
					return ValidSheet;
				}
				else
				{
					strWkbkPath = ndWkbkPath.InnerText;
				}
				
				//节点中记录的工作表名称
				string strSheetName = "";
				XmlNode ndShetName = WorksheetNode[DataBasePath.Nd3_SheetName];
				if (ndShetName == null)
				{
					return ValidSheet;
				}
				else
				{
					strSheetName = ndShetName.InnerText;
				}
				
				
				//---先检测工作表所在的工作簿是否在有效的并成功打开和返回的工作簿列表中
				Workbook ValidWkbk = null;
				foreach (Workbook Wkbk in FileContents.lstWkbks)
				{
					if (string.Compare(strWkbkPath, Wkbk.FullName, true) == 0)
					{
						ValidWkbk = Wkbk;
						break;
					}
				}
				
				//---- 根据工作簿的有效性与否执行相应的操作
				if (ValidWkbk != null) //说明工作簿有效
				{
					
					//开始检测工作表的有效性
					
					ValidSheet = ExcelFunction.MatchWorksheet(ValidWkbk, strSheetName);
					if (ValidSheet != null)
					{
						blnNodeForWorksheetValidated = true;
						//
					}
					else //此工作簿不存在，或者是没有成功赋值
					{
						this.F_blnFileValid = false;
						this.F_lstErrorMessage.Add("The Specified Worksheet" + strSheetName + "is not found in workbook: " + strWkbkPath);
					}
					
				}
				else //说明节点中记录的工作簿无效
				{
					this.F_blnFileValid = false;
					this.F_lstErrorMessage.Add("The Specified Workbook for worksheet " + strSheetName + " is not found : " + strWkbkPath);
				}
				
				//返回检测结果
				return ValidSheet;
			}
			
		}
	}
}
