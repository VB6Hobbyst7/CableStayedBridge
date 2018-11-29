// VBConversions Note: VB project level imports

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using CableStayedBridge.Constants;
using CableStayedBridge.DataBase;
using CableStayedBridge.Miscellaneous;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using TextFrame2 = Microsoft.Office.Interop.Excel.TextFrame2;
using XlAxisType = Microsoft.Office.Interop.Excel.XlAxisType;
// End of VB project level imports


namespace CableStayedBridge
{
    namespace All_Drawings_In_Application
    {
        /// <summary>
        /// 测斜数据的曲线图，此
        /// </summary>
        /// <remarks></remarks>
        public class ClsDrawing_Mnt_Incline : clsDrawing_Mnt_RollingBase
        {
            #region   ---  Types

            /// <summary>
            /// 数据列series对象所对应的一些信息对象
            /// </summary>
            /// <remarks>它包括每一条数据系列的对象，及其对应的日期、开挖深度的标志线、记录开挖深度值的文本框</remarks>
            public class SeriesTag_Incline : clsDrawing_Mnt_RollingBase.SeriesTag
            {
                /// <summary>
                /// 与数据系列相关联的挖深直线
                /// </summary>
                private Shape P_DepthLine;

                /// <summary>
                /// 与数据系列相关联的挖深直线
                /// </summary>
                /// <value></value>
                /// <returns></returns>
                /// <remarks></remarks>
                public Shape DepthLine
                {
                    get { return P_DepthLine; }
                    set { P_DepthLine = value; }
                }

                /// <summary>
                /// 与数据系列相关联的文本框
                /// </summary>
                /// 开挖深度信息的文本框
                private Shape P_DepthText;

                /// <summary>
                /// 与数据系列相关联的文本框。设置文本框中的文字： DepthTextbox.TextFrame2.TextRange.Text
                /// </summary>
                /// <value></value>
                /// <returns></returns>
                /// <remarks></remarks>
                public Shape DepthTextbox
                {
                    get { return P_DepthText; }
                    set { P_DepthText = value; }
                }

                /// <summary>
                /// 构造函数
                /// </summary>
                /// <param name="Series"></param>
                /// <param name="ConstructionDate"></param>
                /// <param name="DepthLine">与数据系列相关联的挖深直线</param>
                /// <param name="DepthText">与数据系列相关联的文本框</param>
                /// <remarks></remarks>
                public SeriesTag_Incline(Excel.Series Series, DateTime ConstructionDate,
                    Shape DepthLine = null, TextFrame2 DepthText = null) : base(Series, ConstructionDate)
                {
                    P_DepthLine = DepthLine;
                    P_DepthText = DepthText as Excel.Shape;
                }
            }

            #endregion

            #region  --- Properties

            /// <summary>
            /// 绘图界面与画布的尺寸
            /// </summary>
            /// <value></value>
            /// <returns>此数组有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量</returns>
            /// <remarks></remarks>
            protected override ChartSize ChartSize_sugested
            {
                get
                {
                    return new ChartSize(Data_Drawing_Format.Drawing_Incline.ChartHeight,
                        Data_Drawing_Format.Drawing_Incline.ChartWidth,
                        Data_Drawing_Format.Drawing_Incline.MarginOut_Height,
                        Data_Drawing_Format.Drawing_Incline.MarginOut_Width);
                }
                set { ExcelFunction.SetLocation_Size(ChartSize_sugested, Chart, Application); }
            }

            /// <summary>
            /// 重写基类的属性：图例框的尺寸与位置
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            protected override clsDrawing_Mnt_RollingBase.LegendLocation Legend_Location
            {
                get
                {
                    return new clsDrawing_Mnt_RollingBase.LegendLocation(Data_Drawing_Format.Drawing_Incline.Legend_Height,
                        Data_Drawing_Format.Drawing_Incline.Legend_Width);
                }
            }

            /// <summary>
            /// 监测曲线的数据范围返回的Range对象中，包括此工作表的UsedRange中的第一列，
            /// 即深度的数据；但是不包括第一行的日期数据
            /// </summary>
            /// <remarks></remarks>
            private Excel.Range P_rgMntData;

            /// <summary>
            /// 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
            /// </summary>
            /// <remarks></remarks>
            private clsData_ProcessRegionData P_ProcessRegionData;

            /// <summary>
            /// 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public clsData_ProcessRegionData ProcessRegionData
            {
                get { return P_ProcessRegionData; }
            }

            /// <summary>
            /// 指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。
            /// </summary>
            public bool ShowLabelsWhileRolling
            {
                get { return false; }
            }

            /// <summary> 测斜管的顶部的绝对标高。在测斜管的监测数据中，深度值是相对于测斜管顶部的深度，
            /// 而在监测曲线图中绘制开挖深度或者构件深度时，其中的深度是按绝对标高给出来的，所以需要此属性的值来进行二者之间的转换。
            /// </summary>
            private float _inclineTopElevaion { get; set; }

            /// <summary> 测斜管的顶部的绝对标高。在测斜管的监测数据中，深度值是相对于测斜管顶部的深度，
            /// 而在监测曲线图中绘制开挖深度或者构件深度时，其中的深度是按绝对标高给出来的，所以需要此属性的值来进行二者之间的转换。
            /// </summary>
            public float InclineTopElevaion
            {
                get { return _inclineTopElevaion; }
            }

            #endregion

            #region   ---  Fields

            /// <summary>
            /// 此测斜点的监测数据工作表中的每一天与其在工作表中对应的列号
            /// </summary>
            /// <remarks>返回一个字典，以监测数据的日期key索引数据所在列号item</remarks>
            private Dictionary<DateTime, int> F_dicDateAndColumnNumber;

            /// <summary>
            /// 记录开挖深度的直线与文本框
            /// </summary>
            /// <remarks>返回一个数组，数组中有两个元素，第一个为开挖深度的直线；
            /// 第二个为写有相应文字的文本框</remarks>
            private SeriesTag_Incline _ExcavationDepth_lineAndTextbox;

            private float[] F_YValues;

            #endregion

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="DataSheet">图表对应的数据工作表</param>
            /// <param name="DrawingChart">Excel图形所在的Chart对象</param>
            /// <param name="ParentApp">此图表所在的Excel类的实例对象</param>
            /// <param name="DateSpan">此图表的TimeSpan跨度</param>
            /// <param name="type">此图表的类型，则枚举DrawingType提供</param>
            ///  <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
            /// <param name="date_ColNum">此测斜点的监测数据工作表中的每一天与其在工作表中对应的列号，
            /// 以监测数据的日期key索引数据所在列号item</param>
            /// <param name="usedRg">监测曲线的数据范围，此Range对象中，
            /// 包括此工作表的UsedRange中的第一列，即深度的数据；但是不包括第一行的日期数据</param>
            /// <param name="Info">记录数据信息的文本框</param>
            /// <param name="FirstSeriesTag">第一条数据系列对应的Tag信息</param>
            /// <param name="ProcessRegionData">在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，
            /// 区域数据的Range对象，区域所属的基坑ID及其ID的数据等</param>
            /// <remarks></remarks>
            public ClsDrawing_Mnt_Incline(Excel.Worksheet DataSheet, Excel.Chart DrawingChart, Cls_ExcelForMonitorDrawing ParentApp,
                DateSpan DateSpan, DrawingType type, bool CanRoll, TextFrame2 Info, MonitorInfo DrawingTag,
                MntType MonitorType, Dictionary<DateTime, int> date_ColNum, Excel.Range usedRg,
                SeriesTag_Incline FirstSeriesTag, clsData_ProcessRegionData ProcessRegionData = null)
                : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType,
                    DateSpan, new clsDrawing_Mnt_RollingBase.SeriesTag(FirstSeriesTag.series, FirstSeriesTag.ConstructionDate))
            {
                //
                // --------------------------------------------
                try
                {
                    ClsDrawing_Mnt_Incline with_1 = this;
                    with_1.P_rgMntData = usedRg; //'包括第一列，但是不包括第一行的日期。
                    Excel.Range colRange = usedRg.Columns[1] as Excel.Range;
                    with_1.F_YValues = ExcelFunction.ConvertRangeDataToVector<Single>(colRange);
                    with_1.Information = Info;
                    with_1._ExcavationDepth_lineAndTextbox = FirstSeriesTag;
                    with_1.F_dicDateAndColumnNumber = date_ColNum;
                    with_1.P_ProcessRegionData = ProcessRegionData;

                    with_1._inclineTopElevaion = Project_Expo.InclineTopElevaion;
                    // ----- 集合数据的记录
                    with_1.F_DicSeries_Tag[(int)LowIndexOfObjectsInExcel.SeriesInSeriesCollection] = FirstSeriesTag;

                    // -----对图例进行更新---------
                    //Call LegendRefresh(List_HasCurve)
                }
                catch (Exception ex)
                {
                    MessageBox.Show("构造测斜曲线图出错。" + "\r\n" + ex.Message + "\r\n" +
                                    ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            #region   ---  Rolling事件

            /// <summary>
            /// 图形滚动的Rolling方法
            /// </summary>
            /// <param name="dateThisday">施工当天的日期</param>
            /// <remarks></remarks>
            public override void Rolling(DateTime dateThisday)
            {
                F_RollingDate = dateThisday;
                object lockobject = new object();
                lock (lockobject)
                {
                    Excel.Application app = Chart.Application;
                    app.ScreenUpdating = false;
                    // ------------------- 绘制监测曲线图
                    var Allday = F_dicDateAndColumnNumber.Keys;
                    //考察选定的日期是否有数据
                    TodayState State = default(TodayState);
                    DateTime closedDay = default(DateTime);
                    //
                    if (DateTime.Compare(dateThisday, DateSpan.StartedDate) < 0)
                    {
                        State = TodayState.BeforeStartDay;
                    }
                    else if (DateTime.Compare(dateThisday, DateSpan.FinishedDate) > 0)
                    {
                        State = TodayState.AfterFinishedDay;
                    }
                    else if (Allday.Contains(dateThisday)) //如果搜索的那一天有数据
                    {
                        State = TodayState.DateMatched;
                        closedDay = dateThisday;
                    }
                    else //搜索的那一天没有数据，则查找选定的日期附近最近的一天，并绘制其监测曲线
                    {
                        State = TodayState.DateNotFound;
                        SortedSet<DateTime> sortedlist_AllDays = new SortedSet<DateTime>(Allday);
                        closedDay = GetClosestDay(sortedlist_AllDays, dateThisday);
                    }
                    //
                    CurveRolling(dateThisday, State, closedDay);

                    // -------------------- 移动开挖深度的直线与文本框
                    if (P_ProcessRegionData != null)
                    {
                        float excavationElevation = 0;
                        try
                        {
                            excavationElevation = Convert.ToSingle(P_ProcessRegionData.Date_Elevation[dateThisday]);
                            //ClsData_DataBase.GetElevation(P_rgExcavationProcess, dateThisday)
                        }
                        catch (Exception)
                        {
                            Debug.Print("上面的Exception已经被Try...Catch语句块捕获，不用担心。出错代码位于ClsDrawing_Mnt_Incline.vb中。");
                            DateTime ClosestDate =
                                ClsData_DataBase.FindTheClosestDateInSortedList(
                                    P_ProcessRegionData.Date_Elevation.Keys, dateThisday);
                            excavationElevation = Convert.ToSingle(P_ProcessRegionData.Date_Elevation[ClosestDate]);
                        }
                        MoveExcavation(_ExcavationDepth_lineAndTextbox, excavationElevation, dateThisday, State);
                    }
                    app.ScreenUpdating = true;
                }
            }

            /// <summary>
            /// 根据每天不同的挖深情况，移动挖深线的位置
            /// </summary>
            /// <param name="ExcavationLineAndTextBox">表示挖深深度的直线，文本框中记录有“挖深”二字</param>
            /// <param name="excavElevation"> 当天开挖标高 </param>
            /// <param name="dateThisday">施工当天的日期</param>
            /// <remarks></remarks>
            private void MoveExcavation(SeriesTag_Incline ExcavationLineAndTextBox, float excavElevation,
                DateTime dateThisday, TodayState State)
            {
                Shape iline = ExcavationLineAndTextBox.DepthLine;
                Shape itextbox = ExcavationLineAndTextBox.DepthTextbox;
                //
                if (State == TodayState.DateMatched | State == TodayState.DateNotFound)
                {
                    // 将当天的开挖标高转换为相对于测斜管顶部的深度值
                    float relativedepth = _inclineTopElevaion - excavElevation;
                    //Project_Expo.Elevation_GroundSurface - excavElevation
                    //
                    Excel.Axis axisY = Chart.Axes(Excel.XlAxisType.xlValue) as Excel.Axis;
                    float scalemax = (float)axisY.MaximumScale;
                    float scalemin = (float)axisY.MinimumScale;
                    //
                    float linetop = 0;
                    Excel.PlotArea plotA = Chart.PlotArea;
                    //将相对深度值转换与图表坐标系中的深度坐标值！！！！！
                    linetop =
                        (float)(plotA.InsideTop + plotA.InsideHeight * (relativedepth - scalemin) / (scalemax - scalemin));
                    //
                    iline.Visible = MsoTriState.msoCTrue;
                    iline.Top = linetop;

                    //
                    if (itextbox != null)
                    {
                        itextbox.Visible = MsoTriState.msoCTrue;
                        itextbox.Top = iline.Top - itextbox.Height;

                        //指示此基坑当前的状态，为向下开挖还是向上建造。
                        string strStatus = "";
                        if (P_ProcessRegionData == null || (!P_ProcessRegionData.HasBottomDate))
                        {
                            strStatus = "挖深/施工深度";
                        }
                        else
                        {
                            int CompareIndex = DateTime.Compare(dateThisday, P_ProcessRegionData.BottomDate);
                            if (CompareIndex < 0) //说明还未开挖到底
                            {
                                strStatus = "挖深";
                            }
                            else if (CompareIndex > 0) //说明已经开挖到底，正在施工上部结构
                            {
                                strStatus = "施工深度";
                            }
                            else if (CompareIndex == 0) //说明刚好开挖到基坑底部
                            {
                                strStatus = "开挖到底";
                            }
                        }
                        itextbox.TextFrame2.TextRange.Text = strStatus + Convert.ToString(relativedepth);
                    }
                }
                else
                {
                    iline.Visible = MsoTriState.msoFalse;
                    //
                    if (itextbox != null)
                    {
                        itextbox.Visible = MsoTriState.msoFalse;
                    }
                }
            }

            /// <summary>
            /// 绘制距选定的日期（包括选择的日期）最近的日期的监测数据
            /// </summary>
            /// <param name="dateThisDay">选定的施工日期</param>
            /// <param name="State">指示选定的日期所处的状态</param>
            /// <param name="Closestday">距离选定日期最近的日期，如果选定的日期在工作表中有数据，则等效于dateThisDay</param>
            /// <remarks></remarks>
            private void CurveRolling(DateTime dateThisDay, TodayState State, DateTime Closestday)
            {
                const string strDateFormat = AMEApplication.DateFormat;
                F_DicSeries_Tag[1].ConstructionDate = dateThisDay; //刷新滚动的曲线所代表的日期
                bool blnHasCurve = false;
                //
                //进行监测曲线的滚动
                MovingSeries.Values = F_YValues; //Y轴的数据

                switch (State)
                {
                    case TodayState.BeforeStartDay:
                        MovingSeries.XValues = new object[] { null }; //不能设置为Series.Value=vbEmpty，因为这会将x轴标签中的某一个值设置为0.0。
                        MovingSeries.Name = dateThisDay.ToString(strDateFormat) + " :早于" +
                                            DateSpan.StartedDate.ToString(strDateFormat);
                        blnHasCurve = false;
                        break;

                    case TodayState.AfterFinishedDay:
                        MovingSeries.XValues = new object[] { null }; // VB语法： MovingSeries.XValues = {Nothing}
                        MovingSeries.Name = dateThisDay.ToString(strDateFormat) + " :晚于" +
                                            DateSpan.FinishedDate.ToString(strDateFormat);
                        blnHasCurve = false;
                        break;

                    case TodayState.DateNotFound:
                        MovingSeries.Name = Closestday.ToString(strDateFormat) + "(" +
                                            dateThisDay.ToString(strDateFormat) + "无数据" + ")";
                        blnHasCurve = true;
                        break;

                    case TodayState.DateMatched:
                        MovingSeries.Name = dateThisDay.ToString(strDateFormat);
                        blnHasCurve = true;
                        break;
                }

                if (blnHasCurve)
                {
                    //---------------------------  设置信息文本框中的信息
                    double max = 0;
                    double min = 0;
                    string StrdepthForMax = "";
                    string StrdepthForMin = "";
                    //
                    var strcolumn = F_dicDateAndColumnNumber[Closestday];
                    //当天的监测数据的Range对象
                    var rgDisplacement = P_rgMntData.Columns[strcolumn] as Excel.Range; //只包含对应施工日期的监测数据的那一列Range对象
                    //当天的监测数据，这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。
                    object[] MonitorData = ExcelFunction.ConvertRangeDataToVector<object>(rgDisplacement);
                    MovingSeries.XValues = MonitorData; //X轴的数据
                    //

                    //find the maximun/Minimum displacement
                    max = Sheet_Data.Application.WorksheetFunction.Max(MonitorData);
                    min = Sheet_Data.Application.WorksheetFunction.Min(MonitorData);
                    try
                    {
                        // find the corresponding depth of the maximun displacement
                        //如果 MATCH 函数查找匹配项不成功，在Excel中它会返回错误值 #N/A。 而在VB.NET中，它会直接报错。
                        int Row_Max =
                            Convert.ToInt32(rgDisplacement.Cells[1, 1].Row - 1 +
                                            Sheet_Data.Application.WorksheetFunction.Match(max, MonitorData, 0));
                        float sngDepthForMax =
                            Convert.ToSingle(
                                Sheet_Data.Cells[Row_Max, Data_Drawing_Format.Mnt_Incline.ColNum_Depth].Value);
                        StrdepthForMax = sngDepthForMax.ToString("0.0");
                    }
                    catch (Exception)
                    {
                        StrdepthForMax = " Null ";
                        Debug.Print("搜索最大位移所对应的深度失败！");
                    }
                    try
                    {
                        // find the corresponding depth of the mininum displacement
                        //如果 MATCH 函数查找匹配项不成功，在Excel中它会返回错误值 #N/A。 而在VB.NET中，它会直接报错。
                        int Row_Min =
                            Convert.ToInt32(rgDisplacement.Cells[1, 1].Row - 1 +
                                            Sheet_Data.Application.WorksheetFunction.Match(min, MonitorData, 0));
                        float sngDepthForMin =
                            Convert.ToSingle(
                                Sheet_Data.Cells[Row_Min, Data_Drawing_Format.Mnt_Incline.ColNum_Depth].Value);
                        StrdepthForMin = sngDepthForMin.ToString("0.0");
                    }
                    catch (Exception)
                    {
                        StrdepthForMin = " Null ";
                        Debug.Print("搜索最小位移所对应的深度失败！");
                    }
                    //在信息文本框中输出文本，记录当前曲线的最大与最小位移，以及出现的深度
                    Information.TextRange.Text = "Max:" + max.ToString("0.00") + "mm" + "\t"
                                                 + "In depth of " + StrdepthForMax + "m"
                                                 + "\r\n" + "Min:" + min.ToString("0.00") + "mm" + "\t"
                                                 + "In depth of " + StrdepthForMin + "m";
                }
                else
                {
                    Information.TextRange.Text = "Max:" + "\t"
                                                 + "In depth of "
                                                 + "\r\n" + "Min:" + "\t"
                                                 + "In depth of ";
                }
            }

            #endregion

            #region   ---  数据列的锁定与删除

            //删除
            public override void DeleteSeries(int DeletingSeriesIndex)
            {
                //删除数据系列以及对应的开挖深度直线与开挖深度文本框
                SeriesTag_Incline seriesTag_incline = (SeriesTag_Incline)(F_DicSeries_Tag[DeletingSeriesIndex]);
                if (seriesTag_incline.DepthLine != null)
                {
                    seriesTag_incline.DepthLine.Delete();
                }
                if (seriesTag_incline.DepthTextbox != null)
                {
                    seriesTag_incline.DepthTextbox.Delete();
                }
                //
                base.DeleteSeries(DeletingSeriesIndex);
            }

            //添加
            public override KeyValuePair<int, clsDrawing_Mnt_RollingBase.SeriesTag> CopySeries(int SourceSeriesIndex)
            {
                // 调用基类中的方法
                KeyValuePair<int, clsDrawing_Mnt_RollingBase.SeriesTag> NewSeriesIndex_Tag = base.CopySeries(SourceSeriesIndex);
                //
                // ------------ 为新的数据列创建其对应的Tag
                // ------------ 这是这个类与其基类所不同的地方，将下面这个函数删除，即与其基类中的这个方法相同了。  ------------
                int NewSeriesIndex = NewSeriesIndex_Tag.Key; //新数据列的索引下标，第一条曲线的下标为1；
                SeriesTag_Incline newTag = default(SeriesTag_Incline); //新数据列的标签Tag
                newTag = new SeriesTag_Incline(NewSeriesIndex_Tag.Value.series,
                    NewSeriesIndex_Tag.Value.ConstructionDate);

                try
                {
                    newTag = this.ConstructNewTag(F_DicSeries_Tag[SourceSeriesIndex], newTag.series);
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.Message);
                    MessageBox.Show("在设置\"表示开挖深度的标志线与文本框\"时出错", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //----------------------
                //将基类中的F_DicSeries_Tag.Item(NewSeriesIndex)的值由SeriesTag修改为SeriesTag_Incline
                F_DicSeries_Tag[NewSeriesIndex] = newTag;
                return default(KeyValuePair<int, clsDrawing_Mnt_RollingBase.SeriesTag>);
            }

            /// <summary>
            /// 为新的数据列创建其对应的Tag，并且在Excel的绘图中复制出对应的表示开挖深度的深度线与文本框。
            /// </summary>
            /// <param name="SourceTag">用以参考的数据系列的Tag信息</param>
            /// <param name="newSeries">复制出来的新的数据系列</param>
            /// <returns></returns>
            /// <remarks></remarks>
            private SeriesTag_Incline ConstructNewTag(SeriesTag_Incline SourceTag, Excel.Series newSeries)
            {
                SeriesTag_Incline newTag = new SeriesTag_Incline(newSeries, F_RollingDate);

                // ------------ 为新的数据列创建其对应的Tag
                //数据系列的颜色
                int seriesColor = newSeries.Format.Line.ForeColor.RGB;

                //复制一个新的开挖深度线
                try
                {
                    if (SourceTag.DepthLine != null)
                    {
                        newTag.DepthLine = SourceTag.DepthLine.Duplicate();
                        //线条的位置
                        newTag.DepthLine.Top = SourceTag.DepthLine.Top;
                        //.Left = SourceTag.DepthLine.Left
                        newTag.DepthLine.Left = 0;

                        //线条的颜色
                        newTag.DepthLine.Line.ForeColor.RGB = seriesColor;
                        //线条的阴影
                        newTag.DepthLine.Shadow.Type = MsoShadowType.msoShadow21;
                        //线条的缩放
                        newTag.DepthLine.ScaleWidth(Factor: (float)0.5, RelativeToOriginalSize: MsoTriState.msoFalse, Scale: MsoScaleFrom.msoScaleFromTopLeft);
                    }
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.Message);
                    MessageBox.Show("在设置\"表示开挖深度的标志线\"时出错", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //复制一个新的开挖深度文本框
                try
                {
                    if (SourceTag.DepthTextbox != null)
                    {
                        Shape oldTxtShp = SourceTag.DepthTextbox;
                        newTag.DepthTextbox = oldTxtShp.Duplicate();

                        //设置文本框的位置
                        Shape newTxtShp = newTag.DepthTextbox;
                        //设置文本框的位置
                        newTxtShp.Top = oldTxtShp.Top;
                        //.Left = oldTxtShp.Left
                        newTxtShp.Left = 0;

                        //设置文本框中字体的格式
                        newTxtShp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = seriesColor;
                        newTxtShp.TextFrame2.TextRange.Font.Size = 10;
                        newTxtShp.TextFrame2.TextRange.Font.Bold = MsoTriState.msoFalse;
                    }
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.Message);
                    MessageBox.Show("在设置\"表示开挖深度的文本框\"时出错", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return newTag;
            }

            #endregion
        }
    }
}