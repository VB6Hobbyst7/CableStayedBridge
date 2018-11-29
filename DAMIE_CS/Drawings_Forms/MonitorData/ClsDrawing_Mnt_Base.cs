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

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.GlobalApp_Form;


namespace CableStayedBridge
{
    namespace All_Drawings_In_Application
    {
        public abstract class ClsDrawing_Mnt_Base : IAllDrawings, Dictionary_AutoKey<ClsDrawing_Mnt_Base>.I_Dictionary_AutoKey
        {

            #region  --- Properties

            private Excel.Application F_Application;
            /// <summary>
            /// 图表所在的Excel的Application对象
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Excel.Application Application
            {
                get
                {
                    return this.F_Application;
                }
            }

            /// <summary>
            /// 监测数据的工作表
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Excel.Worksheet Sheet_Data { get; set; }

            /// <summary>
            /// 工作表“标高图”
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Excel.Worksheet Sheet_Drawing { get; set; }

            private Excel.Chart F_myChart;
            /// <summary>
            /// 监测曲线图
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Excel.Chart Chart
            {
                get
                {
                    return F_myChart;
                }
                set
                {
                    F_myChart = value;
                }
            }

            /// <summary>
            /// 绘图界面与画布的尺寸
            /// </summary>
            /// <value></value>
            /// <returns>此结构有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量</returns>
            /// <remarks></remarks>
            protected abstract ChartSize ChartSize_sugested { get; set; }

            /// <summary>
            /// 记录数据信息的文本框
            /// </summary>
            /// <remarks></remarks>
            private Excel.TextFrame2 F_textbox_Info;
            /// <summary>
            /// 记录数据信息的文本框
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Excel.TextFrame2 Information
            {
                get
                {
                    return F_textbox_Info;
                }
                set
                {
                    F_textbox_Info = value;
                }
            }

            private Cls_ExcelForMonitorDrawing F_Class_ParentApp;
            /// <summary>
            /// 此画布所有的Class_ExcelForMonitorDrawing实例
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public Cls_ExcelForMonitorDrawing Parent
            {
                get
                {
                    return F_Class_ParentApp;
                }
            }

            private bool F_blnCanRoll;
            /// <summary>
            /// 指示此图表是否可以按时间进行同步滚动
            /// </summary>
            /// <value></value>
            /// <returns>如果为True，则可以同步滚动，为动态图；如果为False，则为静态图</returns>
            /// <remarks></remarks>
            public bool CanRoll
            {
                get
                {
                    return F_blnCanRoll;
                }
            }

            #region   ---  绘图的各种标签信息

            private DrawingType F_DrawingType;
            /// <summary>
            /// 此图表所属的类型，由枚举DrawingType提供
            /// </summary>
            /// <value></value>
            /// <returns>DrawingType枚举类型，指示此图形所属的类型</returns>
            /// <remarks></remarks>
            public DrawingType Type
            {
                get
                {
                    return F_DrawingType;
                }
            }

            /// <summary>
            /// 监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等
            /// </summary>
            /// <remarks></remarks>
            private MntType P_MntType;
            /// <summary>
            /// 监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public MntType MonitorType
            {
                get
                {
                    return this.P_MntType;
                }
            }

            /// <summary>
            /// 绘图图表的相关标签信息
            /// </summary>
            /// <remarks></remarks>
            private MonitorInfo P_Tags;
            /// <summary>
            /// 绘图图表的相关标签信息
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks>包括其项目名称、基坑编号、监测点位名称</remarks>
            public MonitorInfo Tags
            {
                get
                {
                    return this.P_Tags;
                }
                private set
                {
                    this.P_Tags = value;
                    string TT = value.MonitorItem + "-" + value.ExcavationRegion;
                    if ((this.Type == DrawingType.Monitor_Incline_Dynamic) || (this.Type == DrawingType.Monitor_Incline_MaxMinDepth))
                    {
                        TT = TT + "-" + value.PointName;
                    }
                    //
                    this.Chart_App_Title = TT;
                }
            }

            /// <summary>
            /// 此对象在它所在的集合中的键。用来在集合中索引到此对象：me=集合.item(me.key)
            /// </summary>
            /// <remarks></remarks>
            private int P_Key;
            /// <summary>
            /// 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
            /// 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
            /// 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
            /// 在集合中索引此元素：集合.item(me.key)
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public int Key
            {
                get
                {
                    return this.P_Key;
                }
            }

            private long F_UniqueID;
            /// <summary>
            /// 图表的全局独立ID
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks>以当时时间的Tick属性来定义</remarks>
            public long UniqueID
            {
                get
                {
                    return F_UniqueID;
                }
            }

            /// <summary>
            /// 这个属性值的变化会同步到监测曲线的曲线标题，以及绘图程序的窗口标题中。
            /// </summary>
            /// <remarks></remarks>
            private string P_Chart_App_Title;
            /// <summary>
            /// 这个属性值的变化会同步到监测曲线的曲线标题，以及绘图程序的窗口标题中。
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Chart_App_Title
            {
                get
                {
                    return this.P_Chart_App_Title;
                }
                set
                {
                    this.P_Chart_App_Title = value;
                    this.Chart.ChartTitle.Text = value;
                    if (this.Application != null)
                    {
                        this.Application.Caption = value;
                    }
                }
            }

            #endregion

            #endregion

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="DataSheet">图表对应的数据工作表</param>
            /// <param name="DrawingChart">Excel图形所在的Chart对象</param>
            /// <param name="ParentApp">此图表所在的Excel类的实例对象</param>
            /// <param name="type">此图表所属的类型，由枚举drawingtype提供</param>
            /// <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
            /// <param name="Info">图表中用来显示相关信息的那个文本框对象</param>
            /// <param name="DrawingTag">每一个监测曲线图的相关信息</param>
            /// <param name="MonitorType">监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等</param>
            /// <remarks></remarks>
            public ClsDrawing_Mnt_Base(Excel.Worksheet DataSheet,
                Microsoft.Office.Interop.Excel.Chart DrawingChart, Cls_ExcelForMonitorDrawing ParentApp,
                DrawingType type, bool CanRoll, Excel.TextFrame2 Info,
                MonitorInfo DrawingTag, MntType MonitorType)
            {
                try
                {
                    //设置Excel窗口与Chart的尺寸

                    this.F_Application = ParentApp.Application;
                    ExcelFunction.SetLocation_Size(this.ChartSize_sugested, DrawingChart, this.Application, true);
                    //
                    this.Sheet_Data = DataSheet;
                    this.F_myChart = DrawingChart;
                    this.F_textbox_Info = Info;
                    this.Sheet_Drawing = DrawingChart.Parent.Parent;
                    this.F_blnCanRoll = CanRoll;
                    this.F_DrawingType = type;
                    this.P_MntType = MonitorType;
                    //将此对象添加进其所属的集合中
                    F_Class_ParentApp = ParentApp;
                    //
                    this.P_Key = System.Convert.ToInt32(F_Class_ParentApp.Mnt_Drawings.Add(this));
                    this.F_UniqueID = GeneralMethods.GetUniqueID();
                    //
                    this.Tags = DrawingTag;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("创建基本监测曲线图出错。" + "\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            //
            /// <summary>
            /// 将自己从所在的集合中删除
            /// </summary>
            /// <remarks></remarks>
            public void RemoveFormCollection()
            {
                Thread thd = new Thread(new System.Threading.ThreadStart(RemoveMyselfFromCollection));
                thd.Name = "删除工作表";
                thd.Start();
            }
            private void RemoveMyselfFromCollection()
            {
                try
                {
                    this.Application.ScreenUpdating = false;
                    this.Sheet_Drawing.Delete();
                    this.Parent.Mnt_Drawings.Remove(this.Key);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("删除工作表出错。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
            //
            /// <summary>
            /// 关闭绘图的Excel文档以及其所在的Application程序
            /// </summary>
            /// <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
            /// <remarks></remarks>
            public void Close(bool SaveChanges = false)
            {
                try
                {
                    ClsDrawing_Mnt_Base with_1 = this;
                    Workbook wkbk = with_1.Sheet_Drawing.Parent;
                    wkbk.Close(SaveChanges);
                    with_1.Application.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("关闭监测曲线图出错！" + "\r\n" + ex.Message,
                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            /// <summary>
            /// 按不同的监测数据类型返回对应的坐标轴标签
            /// </summary>
            /// <param name="Drawing_Type"></param>
            /// <param name="MonitorType"></param>
            /// <param name="AxisType"></param>
            /// <param name="AxisGroup">此选项仅对绘制双Y轴或者双X轴的Chart时有效，否则设置了也不会进行处理。</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static string GetAxisLabel(DrawingType Drawing_Type, MntType MonitorType,
                XlAxisType AxisType, XlAxisGroup AxisGroup = XlAxisGroup.xlPrimary)
            {
                string strAxisLabel = "";
                switch (Drawing_Type)
                {
                    // ---------------------  动态测斜曲线图  ----------------------------------------------------
                    case DrawingType.Monitor_Incline_Dynamic:
                        switch (AxisType)
                        {
                            case XlAxisType.xlCategory:
                                strAxisLabel = AxisLabels.Displacement_mm;
                                break;
                            case XlAxisType.xlValue:
                                strAxisLabel = AxisLabels.Depth;
                                break;
                        }
                        break;
                    // ------------------------  测斜位移最值及对应深度  -------------------------------------------------
                    case DrawingType.Monitor_Incline_MaxMinDepth:
                        switch (AxisType)
                        {
                            case XlAxisType.xlCategory:
                                strAxisLabel = AxisLabels.ConstructionDate;
                                break;
                            case XlAxisType.xlValue:
                                switch (AxisGroup)
                                {
                                    case XlAxisGroup.xlPrimary:
                                        strAxisLabel = AxisLabels.Displacement_mm;
                                        break;

                                    case XlAxisGroup.xlSecondary:
                                        strAxisLabel = AxisLabels.Depth;
                                        break;
                                }
                                break;
                        }
                        break;
                    // -----------------  测斜以外的动态监测曲线曲线图  --------------------------------------------------------
                    case DrawingType.Monitor_Dynamic:
                        switch (AxisType)
                        {
                            case XlAxisType.xlCategory:
                                strAxisLabel = AxisLabels.Points;
                                break;
                            case XlAxisType.xlValue:
                                switch (MonitorType)
                                {
                                    case MntType.Struts:
                                        strAxisLabel = AxisLabels.AxialForce;
                                        break;
                                    case MntType.WaterLevel:
                                        strAxisLabel = AxisLabels.Displacement_m;
                                        break;
                                    default:
                                        strAxisLabel = AxisLabels.Displacement_mm;
                                        break;
                                }
                                break;
                        }
                        break;
                    // -------------------------  测斜以外的静态曲线图  ------------------------------------------------
                    case DrawingType.Monitor_Static:
                        switch (AxisType)
                        {
                            case XlAxisType.xlCategory:
                                strAxisLabel = AxisLabels.ConstructionDate;
                                break;
                            case XlAxisType.xlValue:
                                switch (MonitorType)
                                {
                                    case MntType.Struts:
                                        strAxisLabel = AxisLabels.AxialForce;
                                        break;
                                    case MntType.WaterLevel:
                                        strAxisLabel = AxisLabels.Displacement_m;
                                        break;
                                    default:
                                        strAxisLabel = AxisLabels.Displacement_mm;
                                        break;
                                }
                                break;
                        }
                        break;
                    case DrawingType.Xls_SectionalView:
                        switch (AxisType)
                        {
                            case XlAxisType.xlCategory:
                                strAxisLabel = AxisLabels.Excavation;
                                break;
                            case XlAxisType.xlValue:
                                strAxisLabel = AxisLabels.Elevation;
                                break;
                        }
                        break;
                }
                return strAxisLabel;
            }

        }
    }
}
