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
using CableStayedBridge.Constants;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
//using DAMIE.Constants.Data_Drawing_Format;
using CableStayedBridge.DataBase;


namespace CableStayedBridge
{
	namespace DataBase
	{
		
		/// <summary>
		/// 最原始的数据库类
		/// </summary>
		/// <remarks></remarks>
		public class ClsData_DataBase
		{
			
#region   ---  定义与声明
			
#region   ---  Fields
			
			/// <summary>
			/// 数据库中的剖面标高工作表
			/// </summary>
			/// <remarks></remarks>
			private Worksheet F_shtElevation;
			
			
#endregion
			
#region   ---  Events
			
			/// <summary>
			/// 在DataBase中的表示剖面标高的基坑ID属性被整体修改赋值时触发。
			/// </summary>
			/// <param name="dic_Data"></param>
			/// <remarks></remarks>
			public delegate void dic_IDtoComponentsChangedEventHandler(Dictionary<string, clsData_ExcavationID> dic_Data);
			private static dic_IDtoComponentsChangedEventHandler dic_IDtoComponentsChangedEvent;
			
			public static event dic_IDtoComponentsChangedEventHandler dic_IDtoComponentsChanged
			{
				add
				{
					dic_IDtoComponentsChangedEvent = (dic_IDtoComponentsChangedEventHandler) System.Delegate.Combine(dic_IDtoComponentsChangedEvent, value);
				}
				remove
				{
					dic_IDtoComponentsChangedEvent = (dic_IDtoComponentsChangedEventHandler) System.Delegate.Remove(dic_IDtoComponentsChangedEvent, value);
				}
			}
			
			
			/// <summary>
			/// 在DataBase中的表示基坑的施工进度属性被整体修改赋值时触发。
			/// </summary>
			/// <param name="ProcessRange"></param>
			/// <remarks></remarks>
			public delegate void ProcessRangeChangedEventHandler(List<clsData_ProcessRegionData> ProcessRange);
			private static ProcessRangeChangedEventHandler ProcessRangeChangedEvent;
			
			public static event ProcessRangeChangedEventHandler ProcessRangeChanged
			{
				add
				{
					ProcessRangeChangedEvent = (ProcessRangeChangedEventHandler) System.Delegate.Combine(ProcessRangeChangedEvent, value);
				}
				remove
				{
					ProcessRangeChangedEvent = (ProcessRangeChangedEventHandler) System.Delegate.Remove(ProcessRangeChangedEvent, value);
				}
			}
			
			
			/// <summary>
			/// 在DataBase中的表示基坑中的开挖工况汇总信息被整体修改赋值时触发。
			/// </summary>
			/// <param name="NewWorkingStage">以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
			/// 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高</param>
			/// <remarks></remarks>
			public delegate void WorkingStageChangedEventHandler(Dictionary<string, List<clsData_WorkingStage>> NewWorkingStage);
			private static WorkingStageChangedEventHandler WorkingStageChangedEvent;
			
			public static event WorkingStageChangedEventHandler WorkingStageChanged
			{
				add
				{
					WorkingStageChangedEvent = (WorkingStageChangedEventHandler) System.Delegate.Combine(WorkingStageChangedEvent, value);
				}
				remove
				{
					WorkingStageChangedEvent = (WorkingStageChangedEventHandler) System.Delegate.Remove(WorkingStageChangedEvent, value);
				}
			}
			
			
#endregion
			
#region   ---  Properties
			
			
			private Worksheet F_sheet_Points_Coordinates;
			/// <summary>
			/// 工作表：监测点编号与对应坐标(在CAD中的坐标)
			/// </summary>
			/// <remarks></remarks>
public Worksheet sheet_Points_Coordinates
			{
				get
				{
					return this.F_sheet_Points_Coordinates;
				}
			}
			
			/// <summary>
			/// 以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
			/// </summary>
			/// <remarks>返回一个字典，其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</remarks>
			private Dictionary<string, clsData_ExcavationID> F_ID_Components; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
			/// <summary>
			/// 以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
			/// </summary>
			/// <value></value>
			/// <returns>返回一个字典，其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</returns>
			/// <remarks></remarks>
public Dictionary<string, clsData_ExcavationID> ID_Components
			{
				get
				{
					return F_ID_Components;
				}
				set
				{
					this.F_ID_Components = value;
					if (dic_IDtoComponentsChangedEvent != null)
						dic_IDtoComponentsChangedEvent(value);
				}
			}
			
			/// <summary>
			/// 整个工程项目中的所有的基坑区域的集合
			/// </summary>
			/// <remarks></remarks>
			private List<clsData_ProcessRegionData> P_lst_ProcessRegion = new List<clsData_ProcessRegionData>();
			/// <summary>
			/// 整个工程项目中的所有的基坑区域的集合，其中包含了与此基坑区域相关的各种信息，比如区域名称，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public List<clsData_ProcessRegionData> ProcessRegion
			{
				get
				{
					return this.P_lst_ProcessRegion;
				}
				set
				{
					this.P_lst_ProcessRegion = value;
					if (ProcessRangeChangedEvent != null)
						ProcessRangeChangedEvent(value);
				}
			}
			
			/// <summary>
			/// 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据
			/// </summary>
			/// <remarks></remarks>
			private clsData_ShapeID_FinishedDate F_ShapeIDAndFinishedDate;
			/// <summary>
			/// 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据
			/// </summary>
			/// <value></value>
			/// <returns>返回一个数组，数组中有两个元素，其中每一个元素都是以向量的形式记录了这两列数据。</returns>
			/// <remarks>向量的第一个元素的下标值为0。</remarks>
public clsData_ShapeID_FinishedDate ShapeIDAndFinishedDate
			{
				get
				{
					return F_ShapeIDAndFinishedDate;
				}
				private set
				{
					F_ShapeIDAndFinishedDate = value;
				}
			}
			
			/// <summary>
			/// 以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
			/// 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高
			/// </summary>
			/// <remarks></remarks>
			private Dictionary<string, List<clsData_WorkingStage>> F_WorkingStage;
			/// <summary>
			/// 以基坑区域（可以是某一个大基坑，而可以是基坑中的一个小分区）的名称，来索引它的开挖工况，
			/// 开挖工况中包括每一个特征日期所对应的开挖工况描述以及开挖标高
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Dictionary<string, List<clsData_WorkingStage>> WorkingStage
			{
				get
				{
					return this.F_WorkingStage;
				}
				private set
				{
					this.F_WorkingStage = value;
					if (WorkingStageChangedEvent != null)
						WorkingStageChangedEvent(value);
				}
			}
			
#endregion
			
#endregion
			
			/// <summary>
			/// 构造函数
			/// </summary>
			/// <param name="FileContents">根据项目文件来提取相应的工作簿和工作表对象</param>
			/// <remarks></remarks>
			public ClsData_DataBase(clsData_FileContents FileContents)
			{
				if (FileContents != null)
				{
					
					clsData_FileContents with_1 = FileContents;
					//开挖剖面工作表及其数据提取
					this.F_shtElevation = with_1.Sheet_Elevation;
					try
					{
						if (this.F_shtElevation != null)
						{
							//Me.ID_Components的赋值一定要在施工进度工作表的赋值之前，因为后面在赋值时可能会用到Me.ID_Components的值。
							this.ID_Components = ExtractElevation(F_shtElevation);
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show("基坑ID及剖面标高信息提取出错，请检查\"剖面标高\"工作表的数据格式是否正确。" +
							"\r\n" + ex.Message + "\r\n" + "报错位置：" +
							ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ID_Components = null;
					}
					
					//施工进度工作表及其数据提取
					if (with_1.lstSheets_Progress != null)
					{
						try
						{
							this.ProcessRegion = GetDataForProcessRegion(with_1.lstSheets_Progress);
						}
						catch (Exception ex)
						{
							MessageBox.Show("基坑群的施工进度信息提取出错，请检查\"施工进度\"工作表的数据格式是否正确。" +
								"\r\n" + ex.Message + "\r\n" + "报错位置：" +
								ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ProcessRegion = null;
						}
					}
					
					//记录测点坐标的工作表
					this.F_sheet_Points_Coordinates = with_1.Sheet_PointCoordinates;
					
					//Excel中记录Visio中开挖分块的形状的相关信息的工作表
					try
					{
						if (with_1.Sheet_PlanView != null)
						{
							this.F_ShapeIDAndFinishedDate = GetShapeID_FinishedDatePair(with_1.Sheet_PlanView);
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show("开挖分块信息提取出错，请检查\"开挖分块\"工作表的数据格式是否正确。" +
							"\r\n" + ex.Message + "\r\n" + "报错位置：" +
							ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.F_ShapeIDAndFinishedDate = null;
					}
					
					//开挖工况信息
					try
					{
						//Dim wkbk As Workbook = Me.F_shtElevation.Parent
						//Dim shtWorkingStage As Worksheet = wkbk.Worksheets.Item("开挖工况")
						if (with_1.Sheet_WorkingStage != null)
						{
							this.WorkingStage = GetWorkingStage(with_1.Sheet_WorkingStage);
						}
						else
						{
							this.WorkingStage = null;
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show("提取开挖工况信息出错！" + "\r\n" + ex.Message + "\r\n" +
							"报错位置：" + ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.WorkingStage = null;
					}
				}
			}
			
			//剖面标高工作表的数据提取
			/// <summary>
			/// 创建字典：dic_IDtoElevationAndData
			/// </summary>
			/// <param name="shtElevation"></param>
			/// <returns> 创建一个字典，以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
			/// 其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</returns>
			/// <remarks></remarks>
			public Dictionary<string, clsData_ExcavationID> ExtractElevation(Worksheet shtElevation)
			{
				Dictionary<string, clsData_ExcavationID> IDtoElevationAndData = new Dictionary<string, clsData_ExcavationID>();
				//创建一个字典dic_IDtoElevationAndData，以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组
				//其key为基坑的ID，value为[此基坑的标高项目,对应的标高]的数组
				Worksheet with_1 = shtElevation;
				const byte cstRowID = Data_Drawing_Format.DB_Sectional.RowNum_ID;
				const byte cstColFirstID = Data_Drawing_Format.DB_Sectional.ColNum_FirstID;
				//
				int ColumnCount = with_1.UsedRange.Columns.Count;
				foreach (Range icell in with_1.Range(with_1.Cells[cstRowID, cstColFirstID], with_1.Cells[cstRowID, ColumnCount]))
				{
					//如果单元格是位于合并单元格内
					if (icell.MergeCells)
					{
						
						// ----------------- 确定工作表中第一行的基坑名称所在的单元格
						if (string.Compare(icell.MergeArea.Address, System.Convert.ToString(icell.Offset(0, -1).MergeArea.Address)) != 0)
						{
							//说明此单元格不是位于两个单元格所形成的合并单元格中的第二个，所以它是位于第一个
							
							//合并单元格中的第一个单元格，其所在的列即为此基坑ID下的结构构件数据所在的列，
							//而其后面一列即为此基坑ID下的结构构件的标高数据所在的列。
							Range firstInMerge = icell.MergeArea.Cells[1, 1];
							//基坑ID
							string ID = System.Convert.ToString(firstInMerge.Value);
							//基坑底部标高
							float ExcavationBottom = 0;
							//底板顶部标高
							float BottomFloor = 0;
							
							//将基坑名称与这个基坑的结构和标高相对应，并放在一个集合中
							var firstCellColNum = firstInMerge.Column;
							Component[] ArrItemAndData = ColNumToComponents(firstCellColNum, ref ExcavationBottom, ref BottomFloor);
							//
							clsData_ExcavationID ExcavationID = new clsData_ExcavationID(ID, ExcavationBottom, BottomFloor, ArrItemAndData);
							if (ExcavationID == null)
							{
								return default(Dictionary<string, clsData_ExcavationID>);
							}
							IDtoElevationAndData.Add(ID, ExcavationID);
						}
					}
				}
				return IDtoElevationAndData;
			}
			/// <summary>
			/// 由基坑ID所在的列号返回此基坑ID对应的每一个构件项目的名称与标高
			/// </summary>
			/// <param name="firstCellColNum">第一个单元格的列号，即为此基坑ID下的结构构件数据所在的列</param>
			/// <param name="ExcavationBottom">作为输出变量，表示基坑底的标高</param>
			/// <param name="BottomFloor">作为输出变量，表示底板顶的标高</param>
			/// <returns>数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值</returns>
			/// <remarks></remarks>
			private Component[] ColNumToComponents(int firstCellColNum, ref float ExcavationBottom, ref float BottomFloor)
			{
				//提取数据
				Range rg = default(Range);
				rg = this.F_shtElevation.Range(this.F_shtElevation.Cells[3, firstCellColNum], this.F_shtElevation.Cells[DB_Sectional.RowNum_EndRowInElevation, firstCellColNum + 1]);
				object[,] AllData = rg.Value;
				//
				try
				{
					UInt16 index = 0;
					byte lb = (byte) 0;
					Component[] Components = new Component[(AllData.Length - 1) - lb + 1];
					for (var i = lb; i <= (AllData.Length - 1); i++)
					{
						if (AllData[(int) i, 2] != null)
						{
							string tag = System.Convert.ToString(AllData[(int) i, 1]);
							float Elevation = System.Convert.ToSingle(AllData[(int) i, 2]);
							
							// ----------------- 判断每一个构件的类型 -------------------------------------
							//"Contains" method performs an ordinal (case-sensitive and culture-insensitive) comparison.
							ComponentType CompontType = ComponentType.Others;
							if (tag.IndexOf(DB_Sectional.identifier_TopOfBottomSlab, System.StringComparison.OrdinalIgnoreCase) >= 0)
							{
								BottomFloor = Elevation;
								CompontType = ComponentType.TopOfBottomSlab;
								
							}
							else if (tag.IndexOf(DB_Sectional.identifier_ExcavationBottom, System.StringComparison.OrdinalIgnoreCase) >= 0)
							{
								ExcavationBottom = Elevation;
								CompontType = ComponentType.ExcavationBottom;
								
							}
							else if (tag.IndexOf(DB_Sectional.identifier_Floor, System.StringComparison.OrdinalIgnoreCase) >= 0)
							{
								CompontType = ComponentType.Floor;
								
							}
							else if (tag.IndexOf(DB_Sectional.identifier_struts, System.StringComparison.OrdinalIgnoreCase) >= 0)
							{
								CompontType = ComponentType.Strut;
								
							}
							else if (tag.IndexOf(DB_Sectional.identifier_Ground, System.StringComparison.OrdinalIgnoreCase) >= 0)
							{
								CompontType = ComponentType.Ground;
							}
							// -------------------------------------------------
							Components[index] = new Component(CompontType, tag, Elevation);
							index++;
						}
					}
					//剔除数据库中为空的构件
					Component[] arrComponentsWithoutEmpty = new Component[index - 1 + 1];
					for (var i2 = 0; i2 <= index - 1; i2++)
					{
						arrComponentsWithoutEmpty[(int) i2] = Components[(int) i2];
					}
					return arrComponentsWithoutEmpty;
				}
				catch (Exception ex)
				{
					MessageBox.Show("提取数据库中的基坑ID及其构件的信息出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, 
						"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return null;
				}
			}
			
			//开挖分块工作表的数据提取
			/// <summary>
			/// 获取工作表“开挖分块”中的“形状名”与“完成日期”的组合。
			/// </summary>
			/// <param name="shtRegion">开挖分块工作表</param>
			/// <returns></returns>
			/// <remarks>数组的第一个元素的下标值为0</remarks>
			private clsData_ShapeID_FinishedDate GetShapeID_FinishedDatePair(Worksheet shtRegion)
			{
				
				const byte cstColNum_ShapeID = DB_ExcavRegionForVisio.ColNum_ShapeID; //工作表“开挖分块”中的“形状名”所在的列号
				const byte cstColNum_FinishedDate = DB_ExcavRegionForVisio.ColNum_FinishedDate; //工作表“开挖分块”中的“完成日期”所在的列号
				const byte cstRowNum_FirstShape = DB_ExcavRegionForVisio.RowNum_FirstShape; //第一个形状数据所在的行号
				
				//从工作表的range中提取数据
				Range rgFinishedDate = default(Range);
				Range rgShapeID = default(Range);
				Worksheet with_1 = shtRegion;
				int endrow = with_1.UsedRange.Rows.Count;
				rgFinishedDate = with_1.Columns[cstColNum_FinishedDate].range(with_1.Cells[cstRowNum_FirstShape, 1], with_1.Cells[endrow, 1]);
				rgShapeID = with_1.Columns[cstColNum_ShapeID].range(with_1.Cells[cstRowNum_FirstShape, 1], with_1.Cells[endrow, 1]);
				//将Object类型的二维数组转换为Integer类型与Date类型的一维数组，而且让数组的第一个元素的下标值为0
				int[] intShapeID = ExcelFunction.ConvertRangeDataToVector<int>(rgShapeID);
				DateTime[] dateFinishedDate = ExcelFunction.ConvertRangeDataToVector<DateTime>(rgFinishedDate);
				return new clsData_ShapeID_FinishedDate(intShapeID, dateFinishedDate);
			}
			
			//开挖工况工作表的数据提取
			/// <summary>
			/// 获取工作表“开挖工况”中的基坑区域名称，以及每一个基坑区域中的开挖工况信息
			/// </summary>
			/// <param name="shtWorkingStage"></param>
			/// <returns></returns>
			/// <remarks></remarks>
			private Dictionary<string, List<clsData_WorkingStage>> GetWorkingStage(Worksheet shtWorkingStage)
			{
				//
				object[,] allData = shtWorkingStage.UsedRange.Value;
				byte lb = (byte) (Information.LBound((System.Array) allData, 2)); //对于Excel中返回的数组，第一个元素的下标值为1，而不是0
				UInt16 ColumnsCount = Information.UBound((System.Array) allData, 2) - lb + 1;
				UInt16 RowsCount = (allData.Length - 1) - lb + 1;
				//如果数据列的列数是3的整数，说明工作表的列的数据排布是规范的。
				if ((Information.UBound((System.Array) allData, 2) - lb + 1) % DB_WorkingStage.ColCount_EachRegion != 0)
				{
					MessageBox.Show("开挖工况信息工作表中的数据列的排版不规范，请检查。", 
						"Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return default(Dictionary<string, List<clsData_WorkingStage>>);
				}
				//基坑区域的个数
				UInt16 RegionsCount = ColumnsCount / DB_WorkingStage.ColCount_EachRegion;
				//
				Dictionary<string, List<clsData_WorkingStage>> dicWS = new Dictionary<string, List<clsData_WorkingStage>>();
				//
				for (UInt16 index_Region = 0; index_Region <= RegionsCount - 1; index_Region++) //第几个基坑区域
				{
					string strRegionName = System.Convert.ToString(allData[DB_WorkingStage.RowNum_RegionName - 1 + lb, index_Region * DB_WorkingStage.ColCount_EachRegion + lb]);
					List<clsData_WorkingStage> listWS = new List<clsData_WorkingStage>();
					//
					for (UInt16 RN_WS = DB_WorkingStage.RowNum_FirstStage - 1 + lb; RN_WS <= RowsCount - 1 + lb; RN_WS++) //一个基坑区域中每一个工况在数组中的行号
					{
						UInt16 CN_FirstReference = index_Region * DB_WorkingStage.ColCount_EachRegion + lb;
						try
						{
							//提取数组中的每一行开挖工况，并将其转换为指定的类型。
							//如果数据提取或者转换出错，说明此基坑区域的施工工况信息已经提取完毕。
							string strDescription = System.Convert.ToString(allData[RN_WS, CN_FirstReference + DB_WorkingStage.Index_Description - lb]);
							float Elevation = System.Convert.ToSingle(allData[RN_WS, CN_FirstReference + DB_WorkingStage.Index_Elevation - lb]);
							DateTime ConstructionDate = System.Convert.ToDateTime(allData[RN_WS, CN_FirstReference + DB_WorkingStage.Index_ConstructionDate - lb]);
							if (ConstructionDate.Ticks == 0)
							{
								goto endOfForLoop;
							}
							//
							listWS.Add(new clsData_WorkingStage(strDescription, ConstructionDate, Elevation));
						}
						catch (Exception)
						{
							goto endOfForLoop;
						}
					}
endOfForLoop:
					dicWS.Add(strRegionName, listWS);
				}
				return dicWS;
			}
			
#region   ---  施工进度工作表
			
			//施工进度工作表的操作
			/// <summary>
			/// 施工进度工作表的操作：得到所有基坑区域的施工进度信息
			/// </summary>
			/// <returns></returns>
			/// <remarks></remarks>
			private List<clsData_ProcessRegionData> GetDataForProcessRegion(List<Worksheet> lst_Sheets_Progress)
			{
				//--------------------------------------------
				const byte cstColNum_FirstRegionInProgress = DB_Progress.ColNum_theFirstRegion;
				const byte cstColNum_DateList = DB_Progress.ColNum_DateList;
				const byte cstRowNum_FirstDate = DB_Progress.RowNum_TheFirstDay;
				const byte cstRowNum_ExcavTag = DB_Progress.RowNum_ExcavTag;
				const byte cstRowNum_ExcavPosition = DB_Progress.RowNum_ExcavPosition;
				//--------------------------------------------
				List<clsData_ProcessRegionData> lstRegion = new List<clsData_ProcessRegionData>();
				//Dim dicExcavTagToColRange As New Dictionary(Of String, Range)(StringComparer.OrdinalIgnoreCase)
				foreach (Worksheet sht in lst_Sheets_Progress)
				{
					Microsoft.Office.Interop.Excel.Range UsedRg = sht.UsedRange;
					UInt16 btRegionCount = UsedRg.Columns.Count;
					UInt16 btRowsCount = UsedRg.Rows.Count;
					//以日期所在的数据列作为数据库的主键的列
					Microsoft.Office.Interop.Excel.Range Rg_DateColumn = UsedRg.Columns[cstColNum_DateList];
					//日期列中真正保存日期数据的区域
					Microsoft.Office.Interop.Excel.Range rg_Date = Rg_DateColumn.Range(sht.Cells[cstRowNum_FirstDate, 1], sht.Cells[btRowsCount, 1]);
					object[,] arrDate = rg_Date.Value;
					//--------------------------------------------
					for (byte iCol = cstColNum_FirstRegionInProgress; iCol <= btRegionCount; iCol++)
					{
						try
						{
							Range Rg_Process = UsedRg.Columns[iCol];
							//基坑名：A1/B/C1等
							Range mergecell = Rg_Process.Cells[cstRowNum_ExcavTag, 1].MergeArea(1, 1);
							string strExcavName = System.Convert.ToString(mergecell.Value);
							//基坑分块：普遍区域、东南侧、西侧等
							string strExcavPosition = System.Convert.ToString(Rg_Process.Cells[cstRowNum_ExcavPosition, 1].Value);
							
							//构造此基坑区域的开挖标高的索引
							string strExcavID = System.Convert.ToString(Rg_Process.Cells[Data_Drawing_Format.DB_Progress.RowNum_ExcavID, 1].value);
							clsData_ExcavationID ExcavID = null;
							if (this.ID_Components.Keys.Contains(strExcavID))
							{
								ExcavID = this.ID_Components.Item(strExcavID);
							}
							
							//构造此基坑区域在每一天的施工标高
							object[,] arrElevation = rg_Date.Offset(0, iCol - cstColNum_DateList).Value;
							SortedList<DateTime, Single> Date_Process = GetDate_Elevatin(sht, arrDate, arrElevation);
							
							//
							string strBottomDate = System.Convert.ToString(Rg_Process.Cells[Data_Drawing_Format.DB_Progress.RowNum_BottomDate, 1].value);
							
							//--------------------------------------------
							
							clsData_ProcessRegionData RD = new clsData_ProcessRegionData(strExcavName, 
								strExcavPosition, 
								ExcavID, strBottomDate, 
								Rg_Process, Rg_DateColumn, Date_Process);
							
							//--------------------------------------------
							lstRegion.Add(RD);
						}
						catch (Exception)
						{
							
							return new List<clsData_ProcessRegionData>();
						}
					}
					
				}
				return lstRegion;
			}
			
			/// <summary>
			/// 创建每一天的日期到当天的开挖标高的索引
			/// </summary>
			/// <param name="arrDate">在施工进度工作表中，保存日期的那一列数组，数组中只有日期的数据，而没有前几行的表头信息数据。
			/// 从标高数据的搜索与处理方式上，要求此列日期数据必须是从早到晚进行排列的！</param>
			/// <param name="arrElevation">在施工进度工作表中，保存日期的那一列数组，数组中只有开挖标高的数据，而没有前几行的表头信息数据</param>
			/// <returns>此基坑区域的每一天所对应的开挖标高</returns>
			/// <remarks>如果在搜索的某一天天没有找到对应的开挖标高的数据，那么就开始向更早的日期进行搜索，
			/// 在数据库的设计中认为，如果今天没有开挖标高的数据，说明今天的开挖标高位于早于今天的最近的那一天的有效标高与晚于今天的最近
			/// 的那一天的标高之间，在处理时，即认为今天的开挖标高还是位于早于今天的最近的那一天的有效标高。</remarks>
			private SortedList<DateTime, float> GetDate_Elevatin(Microsoft.Office.Interop.Excel.Worksheet sheetData, object[,] arrDate, object[,] arrElevation)
				{
				SortedList<DateTime, Single> Sorted_Date_Elevation = new SortedList<DateTime, Single>();
				float Ground = Project_Expo.Elevation_GroundSurface;
				//
				byte lb = (byte) 0;
				DateTime Dt = default(DateTime);
				object Elevation = null;
				float ReferenceElevation = Ground;
				//
				UInt16 index_Sortedlist = 0;
				int i = lb;
				try
				{
					for (i = lb; i <= (arrDate.Length - 1); i++)
					{
						Dt = System.Convert.ToDateTime(arrDate[i, lb]);
						Elevation = arrElevation[i, lb];
						//------- 确定具体的开挖标高 "Elevation" 的值 ---------------
						if (Elevation == null || Information.VarType(Elevation) == VariantType.Empty) //说明在今天没有找到对应的开挖标高的数据，
						{
							//那么就要开始向更早的日期进行搜索，在数据库的设计中认为，如果今天没有开挖标高的数据，
							//说明今天的开挖标高位于早于今天的最近的那一天的有效标高与晚于今天的最近的那一天的标高之间，
							//在处理时，即认为今天的开挖标高还是位于早于今天的最近的那一天的有效标高。
							Elevation = ReferenceElevation;
						}
						ReferenceElevation = System.Convert.ToSingle(Elevation);
						//--------------------------------------------
						Sorted_Date_Elevation.Add(Dt, Elevation);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("施工进度工作表\" " + sheetData.Name + " \"中，第 " + (i - lb + DB_Progress.RowNum_TheFirstDay).ToString() +
						" 行的数据出错！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return Sorted_Date_Elevation;
			}
			
			//此方法可删！！！由监测点所在基坑区域，以及当天的日期，来得到此点当日的开挖标高
			/// <summary>
			/// 此方法可删！！！ 由监测点所在基坑区域，以及当天的日期，来得到此点当日的开挖标高
			/// </summary>
			/// <param name="ExcavationRegion">在施工进度工作表中，考察点所在的基坑区域Region，其中每一个Range对象，
			/// 都代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
			/// <param name="dateThisday">考察的日期</param>
			/// <returns>此点当天的开挖标高</returns>
			/// <remarks></remarks>
			public static string GetElevation(Range ExcavationRegion, DateTime dateThisday)
			{
				string Elevation = "";
				int rowByDate = 0;
				//每一张进度表中，日期的第一天的数据所在的行
				const byte cstRowNum_TheFirstDay = DB_Progress.RowNum_TheFirstDay;
				//每一个区域（可能分布于不同的基坑开挖进度的工作表）都有其特有的初始日期
				DateTime dateTheFirstDay = System.Convert.ToDateTime(ExcavationRegion.Worksheet.Cells[cstRowNum_TheFirstDay, DB_Progress.ColNum_DateList].value);
				rowByDate = System.Convert.ToInt32(dateThisday.Subtract(dateTheFirstDay).Days + cstRowNum_TheFirstDay);
				if (rowByDate >= cstRowNum_TheFirstDay) //正常操作
				{
					Elevation = System.Convert.ToString(ExcavationRegion.Cells[rowByDate, 1].Value);
					var i1 = 0;
					//任意一天的基坑开挖深度的确定机制： 如果这一天没有挖深的数据，则向上进行索引，一直索引到此基坑区域开挖的第一天。
					while (!(!string.IsNullOrEmpty(Elevation)|| rowByDate - i1 < cstRowNum_TheFirstDay))
					{
						Elevation = System.Convert.ToString(ExcavationRegion.Cells[rowByDate - i1, 1].Value);
						i1++;
					}
				}
				else //说明这一天的日期小于此基坑区域中记录的最早日期，说明还处于未开挖状态，此时应该将其Depth(i)的值设定为自然地面的标高。
				{
					Elevation = System.Convert.ToString(Project_Expo.Elevation_GroundSurface);
				}
				return Elevation;
			}
			
			/// <summary>
			/// 从SortedList的日期集合中，搜索距离指定的日期最近的那一天。
			/// </summary>
			/// <param name="Keys"></param>
			/// <param name="WantedDate"></param>
			/// <returns></returns>
			/// <remarks></remarks>
			public static DateTime FindTheClosestDateInSortedList(IList<DateTime> Keys, DateTime WantedDate)
			{
				System.Collections.IList with_1 = Keys;
				DateTime FirstDay = System.Convert.ToDateTime(with_1.First);
				DateTime LastDay = System.Convert.ToDateTime(with_1.Last);
				// ------------------------------
				//先考察搜索的日期是否是超出了集合中的界限()
				if (WantedDate.CompareTo(FirstDay) <= 0)
				{
					return FirstDay;
				}
				else if (WantedDate.CompareTo(LastDay) >= 0)
				{
					return LastDay;
				}
				// ------------------------------
				
				DateTime ClosestDate = WantedDate;
				DateTime SearchingDate = WantedDate;
				DateTime Date_Earlier = default(DateTime);
				DateTime Date_Later = default(DateTime);
				bool blnItemFound = false;
				int Index = 0;
				// ----------
				//先向上搜索离之最近的那一天DA
				while (!blnItemFound)
				{
					Index = Keys.IndexOf(SearchingDate);
					if (Index >= 0)
					{
						Date_Earlier = SearchingDate;
						blnItemFound = true;
					}
					else
					{
						//将日期提前一天进行搜索
						SearchingDate = SearchingDate.Subtract(TimeSpan.FromDays(1));
						blnItemFound = false;
					}
				}
				Date_Later = Keys[Index + 1];
				if (Date_Later.Subtract(WantedDate) < WantedDate.Subtract(Date_Earlier))
				{
					ClosestDate = Date_Later;
				}
				else
				{
					ClosestDate = Date_Earlier;
				}
				return ClosestDate;
			}
			
#endregion
			
		}
		
	}
	
}
