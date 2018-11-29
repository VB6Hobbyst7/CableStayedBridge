// VBConversions Note: VB project level imports

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.Constants;
using eZstd.eZAPI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using TextFrame2 = Microsoft.Office.Interop.Excel.TextFrame2;
using XlAxisType = Microsoft.Office.Interop.Excel.XlAxisType;
// End of VB project level imports

//using System.Math;


namespace CableStayedBridge
{
    namespace Miscellaneous
    {

        #region   ---  全局模块:一些杂项的方法

        /// <summary>
        /// 一些杂项的方法
        /// </summary>
        /// <remarks>记录了一些用来处理零碎问题的共享方法</remarks>
        public sealed class GeneralMethods
        {
            /// <summary>
            /// 进行CAD坐标与Visio坐标的坐标变换
            /// </summary>
            /// <param name="CAD_x1"></param>
            /// <param name="CAD_y1"></param>
            /// <param name="CAD_x2"></param>
            /// <param name="CAD_y2"></param>
            /// <param name="Visio_x1"></param>
            /// <param name="Visio_y1"></param>
            /// <param name="Visio_x2"></param>
            /// <param name="Visio_y2"></param>
            /// <returns>返回x与y方向的线性变换的斜率与截距</returns>
            /// <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
            public static Cdnt_Cvsion Coordinate_Conversion(double CAD_x1, double CAD_y1, double
                CAD_x2, double CAD_y2, double
                    Visio_x1, double Visio_y1, double
                        Visio_x2, double Visio_y2)
            {
                Cdnt_Cvsion conversion = new Cdnt_Cvsion();
                conversion.kx = (Visio_x1 - Visio_x2) / (CAD_x1 - CAD_x2);
                conversion.cx = Visio_x1 - conversion.kx * CAD_x1;
                conversion.ky = (Visio_y1 - Visio_y2) / (CAD_y1 - CAD_y2);
                conversion.cy = Visio_y1 - conversion.ky * CAD_y1;
                return conversion;
            }

            /// <summary>
            /// CAD坐标系与Visio坐标系进行线性转换的斜率与截距
            /// </summary>
            /// <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
            public struct Cdnt_Cvsion
            {
                public double kx;
                public double cx;
                public double ky;
                public double cy;
            }

            /// <summary>
            /// 将两个TimeSpan向量进行比较，以扩展为二者区间的并集
            /// </summary>
            /// <param name="DateSpan_Larger">timespan1</param>
            /// <param name="DateSpan_Shorter">timespan2</param>
            /// <returns>返回扩展后的timespan</returns>
            /// <remarks>将两个timespan的区间进行组合，取组合后的最小值与最大值</remarks>
            public static DateSpan ExpandDateSpan(DateSpan DateSpan_Larger, DateSpan DateSpan_Shorter)
            {
                if (DateTime.Compare(DateSpan_Larger.StartedDate, DateSpan_Shorter.StartedDate) == 1)
                {
                    DateSpan_Larger.StartedDate = DateSpan_Shorter.StartedDate;
                }
                if (DateTime.Compare(DateSpan_Larger.FinishedDate, DateSpan_Shorter.FinishedDate) == -1)
                {
                    DateSpan_Larger.FinishedDate = DateSpan_Shorter.FinishedDate;
                }
                return DateSpan_Larger;
            }

            /// <summary>
            /// 返回一个全局的唯一的ID值，用来定义每一个绘图的图表对象
            /// </summary>
            /// <returns></returns>
            /// <remarks>这里先暂且就用当前时间的Tick值来定义</remarks>
            public static long GetUniqueID()
            {
                return DateTime.Now.Ticks;
            }

            /// <summary>
            /// 委托：更新：基坑ID列表框、施工进度列表框
            /// </summary>
            /// <remarks></remarks>
            private delegate void RefreshComoboxHandler(ListControl comboBox, LstbxDisplayAndItem[] DataSource);

            /// <summary>
            /// 在主UI线程中更新：基坑ID列表框、施工进度列表框
            /// </summary>
            /// <remarks></remarks>
            public static void RefreshCombobox(ListControl comboBox, LstbxDisplayAndItem[] DataSource)
            {
                ListControl with_1 = comboBox;
                if (with_1.InvokeRequired) //非UI线程，再次封送该方法到UI线程
                {
                    //对于有输入参数的方法，要将其参数放置在一个数组中，统一传递给BeginInvoke方法。
                    object[] args = new object[] { comboBox, DataSource };
                    with_1.BeginInvoke(new RefreshComoboxHandler(RefreshCombobox), args);
                }
                else
                {
                    //方式一：
                    with_1.DataSource = DataSource;
                    //方式二：
                    //.Items.Clear()
                    //.Items.AddRange(DataSource)
                    //
                    with_1.DisplayMember = LstbxDisplayAndItem.DisplayMember;
                    if (DataSource.Count() > 0)
                    {
                        with_1.SelectedIndex = 0;
                    }
                }
            }


            /// <summary>
            /// 通过释放窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
            /// </summary>
            /// <param name="hWnd">要释放大小的窗口的句柄</param>
            /// <remarks></remarks>
            public static void FixWindow(IntPtr hWnd)
            {
                int hStyle = APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE);
                //禁用最大化的标头及拖拽视窗
                APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE,
                    new IntPtr(hStyle & ~(int)WindowStyle.WS_MAXIMIZEBOX & ~(int)WindowStyle.WS_EX_APPWINDOW));
            }

            /// <summary>
            /// 通过禁用窗口的"最大化"按钮及"拖拽窗口"的功能，来达到释放应用程序窗口大小的效果
            /// </summary>
            /// <param name="hWnd">要固定大小的窗口的句柄</param>
            /// <remarks></remarks>
            public static void unFixWindow(IntPtr hWnd)
            {
                int hStyle = APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE);
                APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE,
                    new IntPtr(hStyle | (int)WindowStyle.WS_MAXIMIZEBOX | (int)WindowStyle.WS_EX_APPWINDOW));
            }

            /// <summary>
            /// 计算数组中的最大值，以及最大值在数组中对应的下标（第一个元素的下标为0）。
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="arr"></param>
            /// <param name="Index">最大值在数组中的下标位置，第一个元素的下标值为0</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static T Max_Array<T>(T[] arr, ref int Index) where T : IComparable
            {
                Index = 0;
                int i = 0;
                T max = arr[Index];
                T refer = arr[Index];
                foreach (T v in arr)
                {
                    if (v.CompareTo(refer) > 0)
                    {
                        max = v;
                        refer = max;
                        Index = i;
                    }
                    i++;
                }
                return max;
            }

            /// <summary>
            /// 计算数组中的最小值，以及最小值在数组中对应的下标（第一个元素的下标为0）。
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="arr"></param>
            /// <param name="Index">最小值在数组中的下标位置，第一个元素的下标值为0</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static T Min_Array<T>(T[] arr, ref int Index) where T : IComparable
            {
                Index = 0;
                T min = arr[Index];
                int i = 0;
                T refer = arr[Index];
                foreach (T v in arr)
                {
                    if (v.CompareTo(refer) < 0)
                    {
                        min = v;
                        refer = min;
                        Index = i;
                    }
                    i++;
                }
                return min;
            }

            /// <summary>
            /// 设置列表控件中显示监测数据的类型，并索引到具体的枚举项
            /// </summary>
            /// <param name="listControl"></param>
            /// <remarks></remarks>
            public static void SetMonitorType(ListControl listControl)
            {
                LstbxDisplayAndItem[] arrMntTypes = null;
                arrMntTypes = new[]
                {
                    new LstbxDisplayAndItem("通用", MntType.General), new LstbxDisplayAndItem("墙体测斜", MntType.Incline),
                    new LstbxDisplayAndItem("立柱垂直位移", MntType.Column), new LstbxDisplayAndItem("支撑轴力", MntType.Struts),
                    new LstbxDisplayAndItem("墙顶水平位移", MntType.WallTop_Horizontal),
                    new LstbxDisplayAndItem("墙顶竖直位移", MntType.WallTop_Vertical),
                    new LstbxDisplayAndItem("地表垂直位移", MntType.EarthSurface),
                    new LstbxDisplayAndItem("水位", MntType.WaterLevel)
                };
                ListControl with_1 = listControl;
                with_1.DisplayMember = LstbxDisplayAndItem.DisplayMember;
                with_1.DataSource = arrMntTypes;
            }
        }

        #endregion

        #region   ---  枚举值

        /// <summary>
        /// 各种与RGB函数返回的值相同的颜色值
        /// </summary>
        /// <remarks>RGB函数返回的整数值与RGB参量的换算关系为：Color属性值=R + 256*G + 256^2*B</remarks>
        public enum Color
        {
            /// <summary>
            /// 表示在开挖剖面图中，已经开挖到基坑底部，正在向上建时的颜色：RGB(20, 200, 230)
            /// </summary>
            /// <remarks>为偏暗的青色</remarks>
            Color_BuildindUp = 15124500,

            /// <summary>
            /// 表示在开挖剖面图中，还未开挖到基坑底部标高时的颜色：RGB(220, 160, 0)
            /// </summary>
            /// <remarks>为偏暗的黄色</remarks>
            Color_DiggingDown = 41180

            // ''' <summary>
            // ''' 表示测斜曲线图中，第1条监测曲线的颜色：RGB(0, 0, 0)
            // ''' </summary>
            // ''' <remarks>对于Chart中的数据系列的颜色进行特别的设定，是因为调用模板中数据系列的颜色时，
            // ''' 它返回的是“自动”，但是在赋值时却不能将颜色值赋值为“自动”。</remarks>
            //Inline1 = 0 ' RGB(0 ,0 ,0)
            //Inline2 = 65280 ' RGB(0 ,255 ,0)
            //Inline3 = 65535 ' RGB(255 ,255 ,0)
            //Inline4 = 3305658   ' RGB(186 ,112 ,50)
            //Inline5 = 3685008   ' RGB(144 ,58 ,56)
            //Inline6 = 4295796   ' RGB(116 ,140 ,65)
            //Inline7 = 4803071   ' RGB(255 ,73 ,73)
            //Inline8 = 4841471   ' RGB(255 ,223 ,73)
            //Inline9 = 4849545   ' RGB(137 ,255 ,73)
            //Inline10 = 7817053  ' RGB(93 ,71 ,119)
            //Inline11 = 9606097  ' RGB(209 ,147 ,146)
            //Inline12 = 16711680 ' RGB(0 ,0 ,255)
            //Inline13 = 16711935 ' RGB(255 ,0 ,255)
            //Inline14 = 16713537 ' RGB(65 ,7 ,255)
            //Inline15 = 16733769 ' RGB(73 ,86 ,255)
            //Inline16 = 16776960 ' RGB(0 ,255 ,255)
        }

        /// <summary>
        /// 枚举Excel中各种对象在集合中的第一个元素的索引下标值
        /// </summary>
        /// <remarks>从概念上来说，这里不应该用枚举类型，而应该用常数类型。但是由于这里的所有项的值都是整数值，
        /// 而且值的范围大概都是非0即1，所以这里用枚举值来表示也是完全可以的。
        /// 一定要注意的是，这里的每一个枚举项都必须要给出对应的下标数值来！！！</remarks>
        public enum LowIndexOfObjectsInExcel
        {
            /// <summary>
            /// 在Excel中用Range.Value返回的二维数组的第一个元素的下标值
            /// </summary>
            /// <remarks></remarks>
            ObjectArrayFromRange_Value = 1,

            /// <summary>
            /// Chart图表中的数据列集合中，第一条曲线对应的下标值
            /// </summary>
            /// <remarks></remarks>
            SeriesInSeriesCollection = 1
        }

        /// <summary>
        /// 对项目文件执行的操作方式
        /// </summary>
        /// <remarks></remarks>
        public enum ProjectState
        {
            /// <summary>
            /// 新建项目
            /// </summary>
            /// <remarks></remarks>
            NewProject,

            /// <summary>
            /// 打开项目
            /// </summary>
            /// <remarks></remarks>
            OpenProject,

            /// <summary>
            /// 编辑项目
            /// </summary>
            /// <remarks></remarks>
            EditProject
        }

        /// <summary>
        /// .NET中的数据类型枚举
        /// </summary>
        /// <remarks></remarks>
        public enum DataType
        {
            /// <summary>
            /// 整数
            /// </summary>
            /// <remarks></remarks>
            Type_Integer,

            /// <summary>
            /// 单精度
            /// </summary>
            /// <remarks></remarks>
            Type_Single,

            /// <summary>
            /// 双精度
            /// </summary>
            /// <remarks></remarks>
            Type_Double,

            /// <summary>
            /// 日期
            /// </summary>
            /// <remarks></remarks>
            Type_Date,

            /// <summary>
            /// 字符串
            /// </summary>
            /// <remarks></remarks>
            Type_String,

            /// <summary>
            /// 对象型
            /// </summary>
            /// <remarks></remarks>
            Type_Object
        }

        #endregion

        #region   ---  类或接口

        /// <summary>
        /// 用来作为ListControl类的.Add方法中的Item参数的类。通过指定ListControl类的DisplayMember属性，来设置列表框中显示的文本。
        /// </summary>
        /// <remarks>
        /// 保存数据时：
        ///  With ListBoxWorksheetsName
        ///       .DisplayMember = LstbxDisplayAndItem.DisplayMember
        ///       .ValueMember = LstbxDisplayAndItem.ValueMember
        ///       .DataSource = arrSheetsName   '  Dim arrSheetsName(0 To sheetsCount - 1) As LstbxDisplayAndItem
        ///  End With
        /// 提取数据时：
        ///  Try
        ///      Me.F_shtMonitorData = DirectCast(Me.ListBoxWorksheetsName.SelectedValue, Worksheet)
        ///  Catch ex As Exception
        ///      Me.F_shtMonitorData = Nothing
        ///  End Try
        /// 或者是：
        ///  Dim lst As LstbxDisplayAndItem = Me.ComboBoxOpenedWorkbook.SelectedItem
        ///  Try
        ///     Dim Wkbk As Workbook = DirectCast(lst.Value, Workbook)
        ///  Catch ex ...
        /// </remarks>
        public class LstbxDisplayAndItem
        {
            /// <summary>
            /// 在列表框中进行显示的文本
            /// </summary>
            /// <remarks>此常数的值代表此类中代表要在列表框中显示的文本的属性名，即"DisplayedText"</remarks>
            public const string DisplayMember = "DisplayedText";

            /// <summary>
            /// 列表框中每一项对应的值（任何类型的值）
            /// </summary>
            /// <remarks>此常数的值代表此类中代表列表框中的每一项绑定的数据的属性名，即"Value"</remarks>
            public const string ValueMember = "Value";

            private readonly object _objValue;

            public dynamic Value
            {
                get { return _objValue; }
            }

            private readonly string _DisplayedText;

            public string DisplayedText
            {
                get { return _DisplayedText; }
            }

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="DisplayedText">用来显示在列表的UI界面中的文本</param>
            /// <param name="Value">列表项对应的值</param>
            /// <remarks></remarks>
            public LstbxDisplayAndItem(string DisplayedText, object Value)
            {
                _objValue = Value;
                _DisplayedText = DisplayedText;
            } //New

            //表示“什么也没有”的枚举值
            /// <summary>
            /// 列表框中用来表示“什么也没有”。
            /// 1、在声明时：listControl控件.Items.Add(New LstbxDisplayAndItem(" 无", NothingInListBox.None))
            /// 2、在选择列表项时：listControl控件.SelectedValue = NothingInListBox.None
            /// 3、在读取列表中的数据时，作出判断：If Not LstbxItem.Value.Equals(NothingInListBox.None) Then ...
            /// </summary>
            /// <remarks></remarks>
            public enum NothingInListBox
            {
                /// <summary>
                /// 什么也没有选择
                /// </summary>
                /// <remarks></remarks>
                None
            }
        } //LstbxItem

        /// <summary>
        /// 此集合的最大特点就是：它的每一个元素的“键”都是一个Integer的值，
        /// 而这个键是在元素添加到集合中时自动生成的。此集合中的值的类型中，都含有一个代表此元素在集合中的键的属性（比如ID），
        /// 由此可以直接通过此元素来索引到它在集合中的位置。
        /// </summary>
        /// <typeparam name="TValue">集合中元素的值，值的类型必须要实现“I_Dictionary_AutoKey”接口</typeparam>
        /// <remarks>  </remarks>
        public class Dictionary_AutoKey<TValue> : Dictionary<int, TValue>
        {
            /// <summary>
            /// 集合类Dictionary_AutoKey中的元素所必须要实现的接口。
            /// </summary>
            /// <remarks></remarks>
            interface I_Dictionary_AutoKey
            {
                //ReadOnly Property Parent As Dictionary_AutoKey(Of TValue)

                /// <summary>
                /// 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
                /// 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
                /// 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
                /// 在集合中索引此元素：集合.item(me.key)
                /// </summary>
                /// <value></value>
                /// <returns></returns>
                /// <remarks></remarks>
                int Key { get; }
            }

            /// <summary>
            /// 构造函数：判断此集合中的元素类型是否实现了接口I_Dictionary_AutoKey。
            /// </summary>
            /// <remarks></remarks>
            public Dictionary_AutoKey()
            {
                Type T = typeof(TValue);
                if (T.GetInterface(typeof(I_Dictionary_AutoKey).Name) == null)
                {
                    //说明此集合中的元素没有实现接口：I_Dictionary_AutoKey
                    MessageBox.Show("集合中的类没有实现接口，请检查!!!" + "\r\n" + "程序编写完成后将此判断删除！！！", "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// 用法：Me.Key=Me所在的集合.Add(Me)
            /// </summary>
            /// <param name="Value"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public new int Add(TValue Value)
            {
                int ItemID = NewID();
                base.Add(ItemID, Value);
                return ItemID;
            }

            /// <summary>
            /// 在集合中添加对象时，自动执行此方法，以返回对象元素在集合中的键。
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            // VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
            private int NewID_myID = 0;

            public int NewID()
            {
                // static int myID = 0; VBConversions Note: Static variable moved to class level and renamed NewID_myID. Local static variables are not supported in C#.
                try
                {
                    NewID_myID++;
                }
                catch (OverflowException)
                {
                    MessageBox.Show("集合中的元素数量溢出");
                }
                return NewID_myID;
            }
        } //Collection_Generic

        /// <summary>
        /// 表示一段施工日期的跨度区间
        /// </summary>
        /// <remarks></remarks>
        public struct DateSpan
        {
            /// <summary>
            /// 施工段的起始日期
            /// </summary>
            /// <remarks></remarks>
            public DateTime StartedDate { get; set; }

            /// <summary>
            /// 施工段的结束日期
            /// </summary>
            /// <remarks></remarks>
            public DateTime FinishedDate { get; set; }

            /// <summary>
            /// 检查DateSpan的跨度中，是否包含指定的日期
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Contains(DateTime dt)
            {
                if (dt.CompareTo(StartedDate) < 0 || dt.CompareTo(FinishedDate) > 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        } //DateSpan

        /// <summary>
        /// 利用ADO.NET连接Excel数据库，并执行相应的操作：
        /// 创建表格，读取数据，写入数据，获取工作簿中的所有工作表名称。
        /// </summary>
        /// <remarks></remarks>
        public class AdoForExcel
        {
            /// <summary>
            /// 创建对Excel工作簿的连接
            /// </summary>
            /// <param name="ExcelWorkbookPath">要进行连接的Excel工作簿的路径</param>
            /// <returns>一个OleDataBase的Connection连接，此连接还没有Open。</returns>
            /// <remarks></remarks>
            public static OleDbConnection ConnectToExcel(string ExcelWorkbookPath)
            {
                string strConn = string.Empty;
                if (ExcelWorkbookPath.EndsWith("xls"))
                {
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                              "Data Source=" + ExcelWorkbookPath + "; " +
                              "Extended Properties=\'Excel 8.0;IMEX=1\'";
                }
                else if (ExcelWorkbookPath.EndsWith("xlsx") || ExcelWorkbookPath.EndsWith("xlsb"))
                {
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                              "Data Source=" + ExcelWorkbookPath + ";" +
                              "Extended Properties=\"Excel 12.0;HDR=YES\"";
                }
                OleDbConnection conn = new OleDbConnection(strConn);
                return conn;
            }

            /// <summary>
            /// 从对于Excel的数据连接中获取Excel工作簿中的所有工作表
            /// </summary>
            /// <param name="conn"></param>
            /// <returns>如果此连接不是连接到Excel数据库，则返回Nothing</returns>
            /// <remarks></remarks>
            public static string[] GetSheetsName(OleDbConnection conn)
            {
                //如果连接已经关闭，则先打开连接
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                if (ConnectionSourceValidated(conn))
                {
                    //获取工作簿连接中的每一个工作表，
                    //注意下面的Rows属性返回的并不是Excel工作表中的每一行，而是Excel工作簿中的所有工作表。
                    DataRowCollection Tables = conn.GetSchema("Tables").Rows;
                    //
                    string[] sheetNames = new string[Tables.Count - 1 + 1];
                    for (int i = 0; i <= Tables.Count - 1; i++)
                    {
                        //注意这里的表格Table是以DataRow的形式出现的。
                        DataRow Tb = Tables[i];
                        object Tb_Name = Tb["TABLE_NAME"];
                        sheetNames[i] = Tb_Name.ToString();
                    }
                    return sheetNames;
                }
                else
                {
                    MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }

            /// <summary>
            /// 创建一个新的Excel工作表，并向其中插入一条数据
            /// </summary>
            /// <param name="conn"></param>
            /// <param name="TableName">要新创建的工作表名称</param>
            /// <remarks></remarks>
            public static void CreateNewTable(OleDbConnection conn, string TableName)
            {
                //如果连接已经关闭，则先打开连接
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                if (ConnectionSourceValidated(conn))
                {
                    using (OleDbCommand ole_cmd = conn.CreateCommand())
                    {
                        //----- 生成Excel表格 --------------------
                        //要新创建的表格不能是在Excel工作簿中已经存在的工作表。
                        ole_cmd.CommandText = "CREATE TABLE CustomerInfo ([" + TableName +
                                              "] VarChar,[Customer] VarChar)";
                        try
                        {
                            //在工作簿中创建新表格时，Excel工作簿不能处于打开状态
                            ole_cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("创建Excel文档 " + TableName + "失败，错误信息： " + ex.Message, "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            /// <summary>
            /// 向Excel工作表中插入一条数据，此函数并不通用，不建议使用
            /// </summary>
            /// <param name="conn"></param>
            /// <param name="TableName">要插入数据的工作表名称</param>
            /// <param name="FieldName">要插入到的字段</param>
            /// <param name="Value">实际插入的数据</param>
            /// <remarks></remarks>
            public static void InsertToTable(OleDbConnection conn, string TableName, string FieldName, object Value)
            {
                //如果连接已经关闭，则先打开连接
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                if (ConnectionSourceValidated(conn))
                {
                    using (OleDbCommand ole_cmd = conn.CreateCommand())
                    {
                        //在插入数据时，字段名必须是数据表中已经有的字段名，插入的数据类型也要与字段下的数据类型相符。

                        try
                        {
                            ole_cmd.CommandText = "insert into [" + TableName + ("$](" + FieldName) + ") values(\'" +
                                                  Convert.ToString(Value) + "\')";
                            //这种插入方式在Excel中的实时刷新的，也就是说插入时工作簿可以处于打开的状态，
                            //而且这里插入后在Excel中会立即显示出插入的值。
                            ole_cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("数据插入失败，错误信息： " + ex.Message);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            /// <summary>
            /// 读取Excel工作簿中的数据
            /// </summary>
            /// <param name="conn">OleDB的数据连接</param>
            /// <param name="SheetName">要读取的数据所在的工作表</param>
            /// <param name="FieldName">在读取的字段</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static string[] GetDataFromExcel(OleDbConnection conn, string SheetName, string FieldName)
            {
                //如果连接已经关闭，则先打开连接
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                if (ConnectionSourceValidated(conn))
                {
                    //创建向数据库发出的指令
                    OleDbCommand olecmd = conn.CreateCommand();
                    //类似SQL的查询语句这个[Sheet1$对应Excel文件中的一个工作表]
                    //如果要提取Excel中的工作表中的某一个指定区域的数据，可以用："select * from [Sheet3$A1:C5]"
                    olecmd.CommandText = "select * from [" + SheetName + "$]";

                    //创建数据适配器——根据指定的数据库指令
                    OleDbDataAdapter Adapter = new OleDbDataAdapter(olecmd);
                    //创建一个数据集以保存数据
                    DataSet dtSet = new DataSet();
                    //将数据适配器按指令操作的数据填充到数据集中的某一工作表中（默认为“Table”工作表）
                    Adapter.Fill(dtSet);
                    //其中的数据都是由 "select * from [" + SheetName + "$]"得到的Excel中工作表SheetName中的数据。
                    int intTablesCount = dtSet.Tables.Count;
                    //索引数据集中的第一个工作表对象
                    DataTable DataTable = dtSet.Tables[0]; // conn.GetSchema("Tables")
                    //工作表中的数据有8列9行(它的范围与用Worksheet.UsedRange所得到的范围相同。
                    //不一定是写有数据的单元格才算进行，对单元格的格式，如底纹，字号等进行修改的单元格也在其中。)
                    int intRowsInTable = DataTable.Rows.Count;
                    int intColsInTable = DataTable.Columns.Count;
                    //提取每一行数据中的“成绩”数据
                    string[] Data = new string[intRowsInTable - 1 + 1];
                    for (int i = 0; i <= intRowsInTable - 1; i++)
                    {
                        Data[i] = Convert.ToString(DataTable.Rows[i][FieldName].ToString());
                    }
                    return Data;
                }
                else
                {
                    MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }

            //私有函数
            /// <summary>
            /// 验证连接的数据源是否是Excel数据库
            /// </summary>
            /// <param name="conn"></param>
            /// <returns>如果是Excel数据库，则返回True，否则返回False。</returns>
            /// <remarks></remarks>
            private static bool ConnectionSourceValidated(OleDbConnection conn)
            {
                //考察连接是否是针对于Excel文档的。
                string strDtSource = conn.DataSource; //"C:\Users\Administrator\Desktop\测试Visio与Excel交互\数据.xlsx"
                string strExt = Path.GetExtension(strDtSource);
                if (strExt == ".xlsx" || strExt == "xls" || strExt == "xlsb")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        } //AdoForExcel

        public class ExcelFunction
        {
            #region   ---  Range到数组的转换

            /// <summary>
            /// 将Excel的Range对象的数据转换为指定数据类型的一维向量,
            /// Range.Value返回一个二维的表格，此函数将其数据按列行拼接为一维数组。
            /// （即按(0,0),(0,1),(1,0),(1,1),(2,0)...的顺序排列）
            /// </summary>
            /// <param name="rg">用于提取数据的Range对象</param>
            /// <returns>返回一个指定类型的一维向量，如“Single()”</returns>
            /// <remarks>直接用Range.Value来返回的数据，其类型只能是Object，
            /// 而其中的数据是一个元素类型为Object的二维数据（即使此Range对象只有一行或者一列）。
            /// 所以要进行显式的转换，将其转换为指定类型的向量或者二维数组，以便于后面的数据操作。</remarks>
            public static T[] ConvertRangeDataToVector<T>(Excel.Range rg)
            {
                //Range中的数据，这是以向量的形式给出的，其第一个元素的下标值很有可能是1，而不是0
                object[,] RangeData = rg.Value as object[,];
                int elementCount = RangeData.Length;
                //
                T[] Value_Out = new T[elementCount - 1 + 1];
                //获取输入的数据类型
                Type DestiType = typeof(T);
                TypeCode TC = Type.GetTypeCode(DestiType);
                //判断此类型的值
                switch (TC)
                {
                    case TypeCode.DateTime:
                        int i_1 = 0;
                        foreach (object V in RangeData)
                        {
                            try
                            {
                                Value_Out[i_1] = (T)V;
                            }
                            catch (Exception ex)
                            {
                                Debug.Print("数据：" + V.ToString() + " 转换为日期出错！将其处理为日期的初始值。" + "\r\n" + ex.Message);
                                //如果输入的数据为double类型，则将其转换为等效的Date
                                object O = DateTime.FromOADate(Convert.ToDouble(V));
                                Value_Out[i_1] = (T)O;
                            }
                            finally
                            {
                                i_1++;
                            }
                        }
                        break;
                    default:
                        int i = 0;
                        foreach (object V in RangeData)
                        {
                            Value_Out[i] = (T)V;
                            i++;
                        }
                        break;
                }
                return Value_Out;
            }

            /// <summary>
            /// 将Excel的Range对象的数据转换为指定数据类型的二维数组
            /// </summary>
            /// <param name="rg">用于提取数据的Range对象</param>
            /// <returns>返回一个指定类型的二维数组，如“Single(,)”</returns>
            /// <remarks>直接用Range.Value来返回的数据，其类型只能是Object，
            /// 而其中的数据是一个元素类型为Object的二维数据（即使此Range对象只有一行或者一列）。
            /// 所以要进行显式的转换，将其转换为指定类型的向量或者二维数组，以便于后面的数据操作。</remarks>
            public static T[,] ConvertRangeDataToMatrix<T>(Range rg)
            {
                object[,] RangeData = rg.Value as object[,];
                //由Range.Value返回的二维数组的第一个元素的下标值，这里应该为1，而不是0
                byte LB_Range = (byte)0;
                //
                int intRowsCount = Convert.ToInt32((RangeData.Length - 1) - LB_Range + 1);
                int intColumnsCount = Convert.ToInt32(Information.UBound((Array)RangeData, 2) - LB_Range + 1);
                //
                T[,] OutputMatrix = new T[intRowsCount - 1 + 1, intColumnsCount - 1 + 1];
                //
                switch (Type.GetTypeCode(typeof(T)))
                {
                    case TypeCode.DateTime:
                        for (int row = 0; row <= intRowsCount - 1; row++)
                        {
                            for (int col = 0; col <= intColumnsCount - 1; col++)
                            {
                                try
                                {
                                    OutputMatrix[row, col] = (T)(RangeData[row + LB_Range, col + LB_Range]);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("数据：" + RangeData[row + LB_Range, col + LB_Range].ToString() +
                                                    " 转换为日期出错！将其处理为日期的初始值。" + "\r\n" + ex.Message);
                                }
                            }
                        }
                        break;
                    default:
                        for (int row = 0; row <= intRowsCount - 1; row++)
                        {
                            for (int col = 0; col <= intColumnsCount - 1; col++)
                            {
                                OutputMatrix[row, col] = (T)RangeData[row + LB_Range, col + LB_Range];
                            }
                        }
                        break;
                }
                return OutputMatrix;
            }

            #endregion

            /// <summary>
            /// 将Excel表中的列的数值编号转换为对应的字符
            /// </summary>
            /// <param name="ColNum">Excel中指定列的数值序号</param>
            /// <returns>以字母序号的形式返回指定列的列号</returns>
            /// <remarks>1对应A；26对应Z；27对应AA</remarks>
            public static string ConvertColumnNumberToString(int ColNum)
            {
                // 关键算法就是：连续除以基，直至商为0，从低到高记录余数！
                // 其中value必须是十进制表示的数值
                //intLetterIndex的位数为digits=fix(log(value)/log(26))+1
                //本来算法很简单，但是要解决一个问题：当value为26时，其26进制数为[1 0]，这样的话，
                //以此为下标索引其下面的strAlphabetic时就会出错，因为下标0无法索引。实际上，这种特殊情况下，应该让所得的结果成为[26]，才能索引到字母Z。
                //处理的方法就是，当所得余数remain为零时，就将其改为26，然后将对应的商的值减1.

                if (ColNum < 1)
                {
                    MessageBox.Show("列数不能小于1");
                }

                List<byte> intLetterIndex = new List<byte>();
                //
                int quotient = 0; //商
                byte remain = 0; //余数
                //
                byte i = (byte)0;
                do
                {
                    quotient = (int)(Conversion.Fix((double)ColNum / 26)); //商
                    remain = (byte)(ColNum % 26); //余数
                    if (remain == 0)
                    {
                        intLetterIndex.Add(26);
                        quotient--;
                    }
                    else
                    {
                        intLetterIndex.Add(remain);
                    }
                    i++;
                    ColNum = quotient;
                } while (!(quotient == 0));
                string alphabetic = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                string ans = "";
                for (i = 0; i <= intLetterIndex.Count - 1; i++)
                {
                    ans = alphabetic[Convert.ToInt32(intLetterIndex[i] - 1)] + ans;
                }
                return ans;
            }

            /// <summary>
            /// 将Excel表中的字符编号转换为对应的数值
            /// </summary>
            /// <param name="colString">以A1形式表示的列的字母序号，不区分大小写</param>
            /// <returns>以整数的形式返回指定列的数值编号，A列对应数值1</returns>
            /// <remarks></remarks>
            public static int ConvertColumnStringToNumber(string colString)
            {
                colString = colString.ToUpper();
                byte ASC_A = (byte)(Strings.Asc("A"));
                int ans = 0;
                for (byte i = 0; i <= colString.Length - 1; i++)
                {
                    string Chr = colString.Substring(i, 1);
                    ans = ans + (Strings.Asc(Chr) - ASC_A + 1) * (int)Math.Pow(26, colString.Length - i - 1);
                }
                return ans;
            }

            /// <summary>
            /// 获取指定Range范围中的左下角的那一个单元格
            /// </summary>
            /// <param name="RangeForSearch">要进行搜索的Range区域</param>
            /// <returns>指定Range区域中的左下角的单元格</returns>
            /// <remarks></remarks>
            public static Range GetBottomRightCell(Range RangeForSearch)
            {
                int RowsCount = 0;
                int ColsCount = 0;
                Range LeftTopCell = RangeForSearch.Cells[1, 1];
                Range with_1 = RangeForSearch;
                RowsCount = with_1.Rows.Count;
                ColsCount = with_1.Columns.Count;
                return
                    RangeForSearch.Worksheet.Cells[LeftTopCell.Row + RowsCount - 1, LeftTopCell.Column + ColsCount - 1];
            }

            #region   ---  工作簿或工作表的匹配

            /// <summary>
            /// 比较两个工作表是否相同。
            /// 判断的标准：先判断工作表的名称是否相同，如果相同，再判断工作表所属的工作簿的路径是否相同，
            /// 如果二者都相同，则认为这两个工作表是同一个工作表
            /// </summary>
            /// <param name="sheet1">进行比较的第1个工作表对象</param>
            /// <param name="sheet2">进行比较的第2个工作表对象</param>
            /// <returns></returns>
            /// <remarks>不用 blnSheetsMatched = sheet1.Equals(sheet2)，是因为这种方法并不能正确地返回True。</remarks>
            public static bool SheetCompare(Worksheet sheet1, Worksheet sheet2)
            {
                bool blnSheetsMatched = false;
                //先比较工作表名称
                if (string.Compare(sheet1.Name, sheet2.Name) == 0)
                {
                    Workbook wb1 = sheet1.Parent;
                    Workbook wb2 = sheet2.Parent;
                    //再比较工作表所在工作簿的路径
                    if (string.Compare(wb1.FullName, wb2.FullName) == 0)
                    {
                        blnSheetsMatched = true;
                    }
                }
                return blnSheetsMatched;
            }

            /// <summary>
            /// 检测工作簿是否已经在指定的Application程序中打开。
            /// 如果最后此工作簿在程序中被打开（已经打开或者后期打开），则返回对应的Workbook对象，否则返回Nothing。
            /// 这种比较方法的好处是不会额外去打开已经打开过了的工作簿。
            /// </summary>
            /// <param name="wkbkPath">进行检测的工作簿</param>
            /// <param name="Application">检测工作簿所在的Application程序</param>
            /// <param name="blnFileHasBeenOpened">指示此Excel工作簿是否已经在此Application中被打开</param>
            /// <param name="OpenIfNotOpened">如果此Excel工作簿并没有在此Application程序中打开，是否要将其打开。</param>
            /// <param name="OpenByReadOnly">是否以只读方式打开</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static Workbook MatchOpenedWkbk(string wkbkPath, Application Application,
                ref bool blnFileHasBeenOpened, bool OpenIfNotOpened = false, bool OpenByReadOnly = true)
            {
                Workbook wkbk = null;
                if (Application != null)
                {
                    //进行具体的检测
                    if (File.Exists(wkbkPath)) //此工作簿存在
                    {
                        //如果此工作簿已经打开
                        foreach (Workbook WkbkOpened in Application.Workbooks)
                        {
                            if (string.Compare(WkbkOpened.FullName, wkbkPath, true) == 0)
                            {
                                wkbk = WkbkOpened;
                                blnFileHasBeenOpened = true;
                                break;
                            }
                        }

                        //如果此工作簿还没有在主程序中打开，则将此工作簿打开
                        if (!blnFileHasBeenOpened)
                        {
                            if (OpenIfNotOpened)
                            {
                                wkbk = Application.Workbooks.Open(Filename: wkbkPath, UpdateLinks: false,
                                    ReadOnly: OpenByReadOnly);
                            }
                        }
                    }
                }
                //返回结果
                return wkbk;
            }

            /// <summary>
            /// 检测指定工作簿内是否有指定的工作表，如果存在，则返回对应的工作表对象，否则返回Nothing
            /// </summary>
            /// <param name="wkbk">进行检测的工作簿对象</param>
            /// <param name="sheetName">进行检测的工作表的名称</param>
            /// <returns></returns>
            /// <remarks></remarks>
            public static Worksheet MatchWorksheet(Workbook wkbk, string sheetName)
            {
                //工作表是否存在
                Worksheet ValidWorksheet = null;
                foreach (Worksheet sht in wkbk.Worksheets)
                {
                    if (string.Compare(sht.Name, sheetName) == 0)
                    {
                        ValidWorksheet = sht;
                        return ValidWorksheet;
                    }
                }
                //返回检测结果
                return ValidWorksheet;
            }

            #endregion

            #region   ---  几何绘图

            /// <summary>
            /// 设置Chart与Applicatin的尺寸，Application的尺寸默认是随着Chart的尺寸而自动变化的。
            /// </summary>
            /// <param name="ChartSize_sugested"></param>
            /// <param name="Chart">进行尺寸设置的Chart对象</param>
            /// <param name="App">进行尺寸设置的Excel程序</param>
            /// <param name="Fixed">是否要将Excel应用程序的窗口固定</param>
            /// <remarks></remarks>
            public static void SetLocation_Size(ChartSize ChartSize_sugested, Excel.Chart Chart = null, Application App = null,
                bool Fixed = true)
            {
                ChartSize with_1 = ChartSize_sugested;
                if (Chart != null)
                {
                    try
                    {
                        Chart.ChartArea.Top = 0;
                        Chart.ChartArea.Left = 0;
                        Chart.ChartArea.Height = ChartSize_sugested.ChartHeight;
                        Chart.ChartArea.Width = ChartSize_sugested.ChartWidth;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("设置Chart尺寸失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                            "Error", MessageBoxButtons.OK);
                    }
                }
                if (App != null)
                {
                    try
                    {
                        //在可以在隐藏的情况下设置程序界面的尺寸，但是在设置之前，一定要确保其WindowState不能为xlMaximized(可以为xlMinimized)
                        if (App.WindowState == Excel.XlWindowState.xlMaximized)
                        {
                            App.WindowState = Excel.XlWindowState.xlNormal;
                        }
                        App.Height = with_1.ChartHeight + with_1.MarginOut_Height;
                        App.Width = with_1.ChartWidth + with_1.MarginOut_Width;
                        if (Fixed)
                        {
                            FixWindow(App.Hwnd);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "设置Excel的窗口尺寸失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Error",
                            MessageBoxButtons.OK);
                    }
                }
            }

            /// <summary>
            /// 将任意形状以指定的值定位在Chart的某一坐标轴中。
            /// </summary>
            /// <param name="ShapeToLocate">要进行定位的形状</param>
            /// <param name="Ax">此形状将要定位的轴</param>
            /// <param name="Value">此形状在Chart中所处的值</param>
            /// <param name="percent">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
            /// 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
            /// 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
            /// <remarks></remarks>
            public static void setPositionInChart(Shape ShapeToLocate, Axis Ax, double Value, double percent = 0)
            {
                Chart cht = (Chart)Ax.Parent;
                if (cht != null)
                {
                    //Try          '先考察形状是否是在Chart之中

                    //    ShapeToLocate = cht.Shapes.Item(ShapeToLocate.Name)
                    //Catch ex As Exception           '如果形状不在Chart中，则将形状复制进Chart，并将原形状删除
                    //    ShapeToLocate.Copy()
                    //    cht.Paste()
                    //    ShapeToLocate.Delete()
                    //    ShapeToLocate = cht.Shapes.Item(cht.Shapes.Count)
                    //End Try
                    //
                    switch (Ax.Type)
                    {
                        case XlAxisType.xlCategory: //横向X轴
                            double PositionInChartByValue_1 = GetPositionInChartByValue(Ax, Value);
                            Shape with_1 = ShapeToLocate;
                            with_1.Left = (float)(PositionInChartByValue_1 - percent * with_1.Width);
                            break;

                        case XlAxisType.xlValue: //竖向Y轴
                            double PositionInChartByValue = GetPositionInChartByValue(Ax, Value);
                            Shape with_2 = ShapeToLocate;
                            with_2.Top = (float)(PositionInChartByValue - percent * with_2.Height);
                            break;
                        case XlAxisType.xlSeriesAxis:
                            MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            break;
                    }
                }
            }

            /// <summary>
            /// 将一组形状以指定的值定位在Chart的某一坐标轴中。
            /// </summary>
            /// <param name="ShapesToLocate">要进行定位的形状</param>
            /// <param name="Ax">此形状将要定位的轴</param>
            /// <param name="Values">此形状在Chart中所处的值</param>
            /// <param name="percents">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
            /// 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
            /// 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
            /// <remarks></remarks>
            public static void setPositionInChart(Axis Ax, Shape[] ShapesToLocate, double[] Values,
                double[] Percents = null)
            {
                // ------------------------------------------------------
                //检查输入的数组中的元素个数是否相同
                UInt16 Count = ShapesToLocate.Length;
                if (Values.Length != Count)
                {
                    MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (Percents != null)
                {
                    if (Percents.Count() != 1 & Percents.Length != Count)
                    {
                        MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                // ------------------------------------------------------
                Chart cht = (Chart)Ax.Parent;
                //
                double max = Ax.MaximumScale;
                double min = Ax.MinimumScale;
                //
                PlotArea PlotA = cht.PlotArea;
                // ------------------------------------------------------
                Shape shp = default(Shape);
                double Value = 0;
                double Percent = Percents[0];
                double PositionInChartByValue = 0;
                // ------------------------------------------------------

                switch (Ax.Type)
                {
                    case XlAxisType.xlCategory: //横向X轴
                        break;


                    case XlAxisType.xlValue: //竖向Y轴
                        if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                shp = ShapesToLocate[i];
                                Value = Values[i];
                                if (Percents.Count() > 1)
                                {
                                    Percent = Percents[i];
                                }
                                PositionInChartByValue = PlotA.InsideTop + PlotA.InsideHeight * (max - Value) / (max - min);
                                shp.Top = (float)(PositionInChartByValue - Percent * shp.Width);
                            }
                        }
                        else //逆序刻度值，说明Y轴数据为上边小下边大
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                shp = ShapesToLocate[i];
                                Value = Values[i];
                                if (Percents.Count() > 1)
                                {
                                    Percent = Percents[i];
                                }
                                PositionInChartByValue = PlotA.InsideTop + PlotA.InsideHeight * (Value - min) / (max - min);
                                shp.Top = (float)(PositionInChartByValue - Percent * shp.Width);
                            }
                        }
                        break;

                    case XlAxisType.xlSeriesAxis:
                        MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                }
            }

            /// <summary>
            /// 根据在坐标轴中的值，来返回这个值在Chart中的几何位置
            /// </summary>
            /// <param name="Ax"></param>
            /// <param name="Value"></param>
            /// <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
            /// 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
            /// <remarks></remarks>
            public static double GetPositionInChartByValue(Axis Ax, double Value)
            {
                double PositionInChartByValue = 0;
                Chart cht = (Chart)Ax.Parent;
                //
                double max = Ax.MaximumScale;
                double min = Ax.MinimumScale;
                //
                PlotArea PlotA = cht.PlotArea;
                switch (Ax.Type)
                {
                    case XlAxisType.xlCategory: //横向X轴
                        double PositionInPlot_1 = 0;
                        if (Ax.ReversePlotOrder == false) //正向分类，说明X轴数据为左边小右边大
                        {
                            PositionInPlot_1 = PlotA.InsideWidth * (Value - min) / (max - min);
                        }
                        else //逆序类别，说明X轴数据为左边大右边小
                        {
                            PositionInPlot_1 = PlotA.InsideWidth * (max - Value) / (max - min);
                        }
                        PositionInChartByValue = PlotA.InsideLeft + PositionInPlot_1;
                        break;

                    case XlAxisType.xlValue: //竖向Y轴
                        double PositionInPlot = 0;
                        if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                        {
                            PositionInPlot = PlotA.InsideHeight * (max - Value) / (max - min);
                        }
                        else //逆序刻度值，说明Y轴数据为上边小下边大
                        {
                            PositionInPlot = PlotA.InsideHeight * (Value - min) / (max - min);
                        }
                        PositionInChartByValue = PlotA.InsideTop + PositionInPlot;
                        break;
                    case XlAxisType.xlSeriesAxis:
                        Debug.Print("暂时不知道这是什么坐标轴");
                        break;
                }
                return PositionInChartByValue;
            }

            /// <summary>
            /// 根据一组形状在某一坐标轴中的值，来返回这些值在Chart中的几何位置
            /// </summary>
            /// <param name="Ax"></param>
            /// <param name="Values">要在坐标轴中进行定位的多个形状在此坐标轴中的数值</param>
            /// <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
            /// 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
            /// <remarks></remarks>
            public static double[] GetPositionInChartByValue(Axis Ax, double[] Values)
            {
                UInt16 Count = Values.Length;
                double[] PositionInChartByValue = new double[Count - 1 + 1];
                // --------------------------------------------------
                Chart cht = (Chart)Ax.Parent;
                //
                double max = Ax.MaximumScale;
                double min = Ax.MinimumScale;
                double Value = 0;
                //
                PlotArea PlotA = cht.PlotArea;
                switch (Ax.Type)
                {
                    case XlAxisType.xlCategory: //横向X轴
                        if (Ax.ReversePlotOrder == false) //正向分类，说明X轴数据为左边小右边大
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                Value = Values[i];
                                PositionInChartByValue[i] = PlotA.InsideLeft +
                                                            PlotA.InsideWidth * (Value - min) / (max - min);
                            }
                        }
                        else //逆序类别，说明X轴数据为左边大右边小
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                Value = Values[i];
                                PositionInChartByValue[i] = PlotA.InsideLeft +
                                                            PlotA.InsideWidth * (max - Value) / (max - min);
                            }
                        }
                        break;

                    case XlAxisType.xlValue: //竖向Y轴
                        if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                Value = Values[i];
                                PositionInChartByValue[i] = PlotA.InsideTop +
                                                            PlotA.InsideHeight * (max - Value) / (max - min);
                            }
                        }
                        else //逆序刻度值，说明Y轴数据为上边小下边大
                        {
                            for (UInt16 i = 0; i <= Count - 1; i++)
                            {
                                Value = Values[i];
                                PositionInChartByValue[i] = PlotA.InsideTop +
                                                            PlotA.InsideHeight * (Value - min) / (max - min);
                            }
                        }
                        break;
                    case XlAxisType.xlSeriesAxis:
                        Debug.Print("暂时不知道这是什么坐标轴");
                        break;
                }
                return PositionInChartByValue;
            }

            /// <summary>
            /// 设置Excel中的文本框格式：无边距、正中排列
            /// </summary>
            /// <param name="TextFrame">要进行格式设置的文本框对象</param>
            /// <param name="TextSize">字体的大小</param>
            /// <param name="VerticalAnchor">文本的竖向排列方式</param>
            /// <param name="HorizontalAlignment">文本的水平排列方式</param>
            /// <param name="Text">文本框中的文本</param>
            /// <remarks></remarks>
            public static void FormatTextbox_Tag(TextFrame2 TextFrame, float TextSize = 8,
                MsoVerticalAnchor VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle,
                MsoParagraphAlignment HorizontalAlignment = MsoParagraphAlignment.msoAlignCenter, string Text = null)
            {
                TextFrame2 with_1 = TextFrame;
                //文本框的边距
                with_1.MarginLeft = 0;
                with_1.MarginRight = 0;
                with_1.MarginTop = 0;
                with_1.MarginBottom = 0;
                //文本的垂直对齐
                with_1.VerticalAnchor = VerticalAnchor;
                //.WordWrap = Microsoft.Office.Core.MsoTriState.msoCTrue
                //.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone
                with_1.TextRange.Font.Size = TextSize;
                with_1.TextRange.Font.Name = AMEApplication.FontName_TNR;
                with_1.TextRange.ParagraphFormat.Alignment = HorizontalAlignment;
                with_1.TextRange.Text = Text;
            }

            #endregion
        }

        #endregion
    }
}