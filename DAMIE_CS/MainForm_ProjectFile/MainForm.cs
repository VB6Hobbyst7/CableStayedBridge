// VBConversions Note: VB project level imports

using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.Constants;
using CableStayedBridge.DataBase;
using CableStayedBridge.Miscellaneous;
using CableStayedBridge.My;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using My.Resources;


namespace CableStayedBridge
{
    namespace GlobalApp_Form
    {
        /// <summary>
        /// 程序的主界面
        /// </summary>
        /// <remarks></remarks>
        public partial class APPLICATION_MAINFORM
        {
            #region Default Instance

            private static APPLICATION_MAINFORM m_uniqueInstance;

            /// <summary>
            /// Added by the VB.Net to C# Converter to support default instance behavour in C#
            /// </summary>
            public static APPLICATION_MAINFORM UniqueInstance
            {
                get
                {
                    if (m_uniqueInstance == null)
                    {
                        m_uniqueInstance = new APPLICATION_MAINFORM();
                        m_uniqueInstance.FormClosed += MUniqueInstanceFormClosed;
                    }

                    return m_uniqueInstance;
                }
                set { m_uniqueInstance = value; }
            }

            private static void MUniqueInstanceFormClosed(object sender, FormClosedEventArgs e)
            {
                m_uniqueInstance = null;
            }

            #endregion

            #region   ---  定义与声明

            #region   ---  字段定义

            /// <summary>
            /// 全局的主程序
            /// </summary>
            /// <remarks></remarks>
            private readonly GlobalApplication GlbApp;

            #endregion

            #region   ---  属性值的定义

            private static APPLICATION_MAINFORM F_main_Form;

            /// <summary>
            /// 共享属性，用来索引启动窗口对象
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks>这一操作并不多余，对于多线程操作，有可能会出现在其他线程不能正确地
            /// 调用到这个唯一的主程序对象，此时可以用这个属性来返回其实例对象。</remarks>
            public static APPLICATION_MAINFORM MainForm
            {
                get { return F_main_Form; }
            }

            #region   ---  操作窗口

            /// <summary>
            /// 图形滚动窗口
            /// </summary>
            /// <remarks></remarks>
            private frmRolling frmRolling;

            /// <summary>
            /// 图形滚动窗口
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public frmRolling Form_Rolling
            {
                get { return frmRolling; }
                set { frmRolling = value; }
            }

            /// <summary>
            /// 生成剖面标高图窗口
            /// </summary>
            /// <remarks></remarks>
            private frmDrawElevation frmSectionView;

            /// <summary>
            /// 生成剖面标高图窗口
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public frmDrawElevation Form_SectionalView
            {
                get { return frmSectionView; }
                set { frmSectionView = value; }
            }

            /// <summary>
            /// 绘制测斜曲线的窗口
            /// </summary>
            /// <remarks></remarks>
            private frmDrawing_Mnt_Incline frmMnt_Incline;

            /// <summary>
            /// 绘制测斜曲线的窗口
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public frmDrawing_Mnt_Incline Form_Mnt_Incline
            {
                get { return frmMnt_Incline; }
                set { frmMnt_Incline = value; }
            }

            /// <summary>
            /// 绘制其他监测曲线的窗口
            /// </summary>
            /// <remarks></remarks>
            private frmDrawing_Mnt_Others frmMnt_Others;

            /// <summary>
            /// 绘制其他监测曲线的窗口
            /// </summary>
            /// <value></value>
            /// <remarks>从此窗口中可以生成非测斜曲线的其他曲线，并且包括其时间分布与空间分布的形式</remarks>
            public frmDrawing_Mnt_Others Form_Mnt_Others
            {
                get { return frmMnt_Others; }
                set { frmMnt_Others = value; }
            }

            /// <summary>
            /// 对项目文件进行操作的窗口
            /// </summary>
            /// <remarks></remarks>
            private frmProjectFile frmProjectFile;

            /// <summary>
            /// 对项目文件进行操作的窗口
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public frmProjectFile Form_ProjectFile
            {
                get { return frmProjectFile; }
            }

            /// <summary>
            /// 打开Visio的开挖平面图的窗口
            /// </summary>
            /// <remarks></remarks>
            private frmDrawingPlan frmVisioPlanView;

            /// <summary>
            /// 打开Visio的开挖平面图的窗口
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public frmDrawingPlan Form_VisioPlanView
            {
                get
                {
                    if (frmVisioPlanView == null)
                    {
                        frmVisioPlanView = new frmDrawingPlan();
                    }
                    return frmVisioPlanView;
                }
            }

            #endregion

            #region   - - 逻辑标志 布尔值

            /// <summary>
            /// 布尔值，用来指示主程序中的绘图窗口是否有新添加的，或者是否被关闭
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool DrawingWindowChanged { get; set; }

            #endregion

            #endregion

            #endregion

            #region   ---  主程序的加载与关闭

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <remarks></remarks>
            public APPLICATION_MAINFORM()
            {
                // This call is required by the designer.
                InitializeComponent();

                //Added to support default instance behavour in C#
                if (m_uniqueInstance == null)
                    m_uniqueInstance = this;
                // Add any initialization after the InitializeComponent() call.
                //----------------------------
                GlbApp = new GlobalApplication();
                //为关键字段赋初始值()
                F_main_Form = this;
                //----------------------------
                MainUI_ProjectNotOpened();
                //获取与文件或文件夹路径有关的数据
                GetPath();
                //----------------------------
                //创建新窗口，窗口在创建时默认是不隐藏的。
                frmSectionView = new frmDrawElevation();
                frmSectionView.FormClosing += frmSectionView.frmDrawSectionalView_FormClosing;
                frmSectionView.Disposed += frmSectionView.frmDrawElevation_Disposed;
                frmMnt_Incline = new frmDrawing_Mnt_Incline();
                frmMnt_Incline.FormClosing += frmMnt_Incline.frmDrawing_Mnt_Incline_FormClosing;
                frmMnt_Incline.Disposed += frmMnt_Incline.frmDrawing_Mnt_Incline_Disposed;
                frmMnt_Incline.DataWorkbookChanged += frmMnt_Incline.frmDrawing_Mnt_Incline_DataWorkbookChanged;
                frmMnt_Incline.Activated += frmMnt_Incline.frmDrawing_Mnt_Incline_Activated;
                frmMnt_Incline.Deactivate += frmMnt_Incline.frmDrawing_Mnt_Incline_Deactivate;
                frmMnt_Others = new frmDrawing_Mnt_Others();
                frmMnt_Others.Load += frmMnt_Others.frmDrawingMonitor_Load;
                frmMnt_Others.FormClosing += frmMnt_Others.frmDrawing_Mnt_Others_FormClosing;
                frmMnt_Others.Disposed += frmMnt_Others.frmDrawing_Mnt_Others_Disposed;
                frmMnt_Others.DataWorkbookChanged += frmMnt_Others.frmDrawing_Mnt_Incline_DataWorkbookChanged;
                frmMnt_Others.Activated += frmMnt_Others.frmDrawing_Mnt_Incline_Activated;
                frmMnt_Others.Deactivate += frmMnt_Others.frmDrawing_Mnt_Incline_Deactivate;
                frmRolling = new frmRolling();

                // ----------------------- 设置MDI窗体的背景
                foreach (Control C in Controls)
                {
                    if (string.Compare(C.GetType().ToString(), "System.Windows.Forms.MdiClient", true) == 0)
                    {
                        MdiClient MDIC = (MdiClient) C;
                        MDIC.BackgroundImage = Resources.线条背景;
                        MDIC.BackgroundImageLayout = ImageLayout.Tile;
                        break;
                    }
                }
                // ----------------------- 设置主程序窗口启动时的状态
                APPLICATION_MAINFORM with_2 = this;
                mySettings_UI mysettings = new mySettings_UI();
                FormWindowState winState = mysettings.WindowState;
                switch (winState)
                {
                    case FormWindowState.Maximized:
                        with_2.WindowState = winState;
                        break;
                    case FormWindowState.Minimized:
                        with_2.WindowState = FormWindowState.Normal;
                        break;
                    case FormWindowState.Normal:
                        with_2.WindowState = winState;
                        with_2.Location = mysettings.WindowLocation;
                        with_2.Size = mysettings.WindowSize;
                        break;
                }
                //在新线程中进行程序的一些属性的初始值的设置
                Thread thd = new Thread(new ThreadStart(myDefaltSettings));
                thd.Name = "程序的一些属性的初始值的设置";
                thd.Start();
            }

            /// <summary>
            /// 程序的一些属性的初始值的设置，这些属性是与UI线程无关的属性，以在新的非UI线程中进行设置。
            /// </summary>
            /// <remarks></remarks>
            private void myDefaltSettings()
            {
                mySettings_Application setting1 = new mySettings_Application();

                ClsDrawing_PlanView.MonitorPointsInformation @struct =
                    new ClsDrawing_PlanView.MonitorPointsInformation();
                @struct.ShapeName_MonitorPointTag = "Tag";
                @struct.pt_CAD_BottomLeft = new PointF(309598.527F, -119668.436F);
                @struct.pt_CAD_UpRight = new PointF(536642.644F, 201852.14F);
                @struct.pt_Visio_BottomLeft_ShapeID = 197;
                @struct.pt_Visio_UpRight_ShapeID = 217;
                setting1.MonitorPointsInfo = @struct;
                //在下面的Save方法中，不知为何为出现两次相同的报错：System.IO.FileNotFoundException
                //可以明确其于多线程无关，但是好在此异常对于程序的运行无影响。
                setting1.Save();
            }

            /// <summary>
            /// 主程序加载
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks>在主程序界面加载时，先启动Splash Screen，再为关键字段赋值，
            /// 然后用一个伪窗口作为主程序界面的背景，最后关闭Splash Screen窗口。</remarks>
            public void mainForm_Load(object sender, EventArgs e)
            {
                //-----------------------根据程序的启动方式的不同，作出不同的操作
                ReadOnlyCollection<string> s = (new WindowsFormsApplicationBase()).CommandLineArgs;
                if (s.Count == 1) //
                {
                    string StartFilePath = s[0];
                    if (File.Exists(StartFilePath))
                    {
                        if (string.Compare(Path.GetExtension(StartFilePath), AMEApplication.FileExtension, true) == 0)
                        {
                            OpenProjectFile(StartFilePath);
                            MainUI_ProjectOpened();
                        }
                    }
                }
            }

            /// <summary>
            /// 获取与文件或文件夹路径有关的数据,并保存在My.Settings中
            /// </summary>
            /// <remarks>用来保证程序在不同机器或文件夹间迁移时能够正常索引</remarks>
            private void GetPath()
            {
                Settings with_1 = Settings.Default;

                //主程序.exe文件所在的文件夹路径，比如：F:\基坑数据\程序编写\群坑分析\AME\bin\Debug
                with_1.Path_MainForm = (new WindowsFormsApplicationBase()).Info.DirectoryPath;

                //"Templates"
                with_1.Path_Template = Path.Combine(Convert.ToString(with_1.Path_MainForm),
                    FolderOrFileName.Folder.Template);

                //用来进行输出的文件夹
                with_1.Path_Output = Path.Combine(Convert.ToString(with_1.Path_MainForm), FolderOrFileName.Folder.Output);

                with_1.Path_DataBase = Path.Combine(Convert.ToString(with_1.Path_MainForm),
                    FolderOrFileName.Folder.DataBase);
            }

            /// <summary>
            /// 退出主程序
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks>
            /// 这里用FormClosing事件来控制主程序的退出，而不用Dispose事件来控制，
            /// 是为了解决其子窗口在各自的FormClosing事件中只是将其进行隐藏并取消了默认的关闭操作，
            /// 所以这里在主程序最先触发的FormClosing事件中就直接执行Me.Dispose()方法，这样就可以
            /// 跳过子窗口的FormClosing事件，而直接退出主程序了。
            /// </remarks>
            public void mainForm_FormClosing(object sender, FormClosingEventArgs e)
            {
                AmeDrawings AllDrawing = GlbApp.ExposeAllDrawings();
                if (AllDrawing.Count() > 0)
                {
                    DialogResult result = MessageBox.Show("还有图表未处理，是否要关闭所有绘图并退出程序", "tip",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button3);
                    if (result == DialogResult.Yes) //关闭AME主程序，同时也关闭所有的绘图程序。
                    {
                        Hide();
                        //' --- 通过新的工作线程来执行绘图程序的关闭。
                        //Dim thd As New Thread(AddressOf Me.QuitDrawingApplications)
                        //With thd
                        //    .Name = "关闭所有绘图程序"
                        //    .Start(AllDrawing)
                        //    thd.Join()
                        //End With
                        // ---
                        //通过Main Thread来执行绘图程序的关闭。
                        QuitDrawingApplications(AllDrawing);
                    } //不关闭图表，但关闭程序
                    else if (result == DialogResult.No)
                    {
                    } //图表与程序都不关闭
                    else if (result == DialogResult.Cancel)
                    {
                        e.Cancel = true;
                        return;
                    }
                }
                // ---------------------- 先隐藏主界面，以达到更好的UI效果。
                Hide();
                //--------------------------- 断后工作
                try
                {
                    //关闭隐藏的Excel数据库中的所有工作簿
                    foreach (Workbook wkbk in GlbApp.ExcelApplication_DB.Workbooks)
                    {
                        wkbk.Close(false);
                    }
                    GlbApp.ExcelApplication_DB.Quit();
                    //这一步非常重要哟
                    GlbApp.ExcelApplication_DB = null;
                }
                catch (Exception)
                {
                    //有可能会出现数据库文件已经被关闭的情况
                }

                //保存主程序窗口关闭时的界面位置与大小
                APPLICATION_MAINFORM with_1 = this;
                mySettings_UI mysetting = new mySettings_UI();
                mysetting.WindowLocation = with_1.Location;
                mysetting.WindowSize = with_1.Size;
                mysetting.WindowState = with_1.WindowState;
                mysetting.Save();

                //---------------------------
                //这里在主程序最先触发的FormClosing事件中就直接执行Me.Dispose()方法，
                //这样就可以跳过子窗口的FormClosing事件，而直接退出主程序了。
                Dispose();
            }

            /// <summary>
            /// 关闭程序中的所有绘图所在的程序，如Excel或者Visio的程序
            /// </summary>
            /// <remarks></remarks>
            private void QuitDrawingApplications(AmeDrawings AllDrawing)
            {
                AmeDrawings with_1 = AllDrawing;
                //开挖剖面图
                if (with_1.SectionalView != null)
                {
                    with_1.SectionalView.Close(false);
                }
                //Visio开挖平面图
                if (with_1.PlanView != null)
                {
                    with_1.PlanView.Close(false);
                }
                //监测曲线图
                foreach (ClsDrawing_Mnt_Base MntDrawing in with_1.MonitorData)
                {
                    MntDrawing.Close(false);
                }
            }

            /// <summary>
            /// 退出程序
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItemExit_Click(object sender, EventArgs e)
            {
                OnFormClosing(new FormClosingEventArgs(CloseReason.ApplicationExitCall, false));
            }

            #endregion

            #region   ---  一般界面操作

            //项目文件的新建与打开
            public void MenuItem_NewProject_Click(object sender, EventArgs e)
            {
                NewProjectFile();
            }

            public void MenuItem_OpenProject_Click(object sender, EventArgs e)
            {
                string FilePath = "";
                OpenFileDialog1.Title = "选择项目文件";
                string FileExtension = AMEApplication.FileExtension;
                OpenFileDialog1.Filter = FileExtension + "文件(*" + FileExtension + ")|*" + FileExtension;
                OpenFileDialog1.FilterIndex = 1;
                if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FilePath = OpenFileDialog1.FileName;
                }
                if (FilePath.Length > 0)
                {
                    OpenProjectFile(FilePath);
                }
            }

            public void MenuItem_EditProject_Click(object sender, EventArgs e)
            {
                EditProjectFile();
            }

            //拖拽操作
            public void APPLICATION_MAINFORM_DragDrop(object sender, DragEventArgs e)
            {
                string[] FileDrop = e.Data.GetData(DataFormats.FileDrop) as string[];
                // DoSomething with the Files or Directories that are droped in.
                string filepath = FileDrop[0];
                string ext = Path.GetExtension(filepath);

                if (string.Compare(ext, AMEApplication.FileExtension, true) == 0)
                {
                    OpenProjectFile(filepath);
                }
                else
                {
                    MessageBox.Show("Can not open file" + filepath + ". Verify that the file is a(an)"
                                    + AMEApplication.FileExtension + " file.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public void APPLICATION_MAINFORM_DragEnter(object sender, DragEventArgs e)
            {
                // See if the data includes text.
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    // There is text. Allow copy.
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    // There is no text. Prohibit drop.
                    e.Effect = DragDropEffects.None;
                }
            }

            //Arrange,子窗口重排
            /// <summary>
            /// 子窗口水平排列
            /// </summary>
            public void ChildrenFormAlligment_Horizontal(object sender, EventArgs e)
            {
                LayoutMdi(MdiLayout.TileHorizontal);
            }

            /// <summary>
            /// 子窗口竖直排列
            /// </summary>
            public void ChildrenFormAlligment_Vertical(object sender, EventArgs e)
            {
                LayoutMdi(MdiLayout.TileVertical);
            }

            /// <summary>
            /// 子窗口层叠
            /// </summary>
            public void ChildrenFormAlligment_Cascade(object sender, EventArgs e)
            {
                LayoutMdi(MdiLayout.Cascade);
            }

            #endregion

            #region   ---  绘制图表

            /// <summary>
            /// 生成剖面标高图
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItemSectionalView_Click(object sender, EventArgs e)
            {
                //数据库中所有的开挖区域
                if (GlbApp.DataBase.ID_Components.Count == 0)
                {
                    MessageBox.Show("没有发现基坑标高相关的数据，请先在项目中添加相关数据。",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    frmSectionView.MdiParent = this;
                    frmSectionView.Show();
                    frmSectionView.WindowState = FormWindowState.Maximized;
                }
            }

            //
            /// <summary>
            /// 生成Visio的平面开挖图
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void ShowForm_DrawingVisioPlanView(object sender, EventArgs e)
            {
                GlobalApplication.Application.DrawVisioPlanView();
            }

            /// <summary>
            /// 生成测斜曲线图
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItemDataMonitored_Click(object sender, EventArgs e)
            {
                frmMnt_Incline.MdiParent = this;
                frmMnt_Incline.Show();
                frmMnt_Incline.WindowState = FormWindowState.Maximized;
            }

            /// <summary>
            /// 生成其他监测数据曲线图
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItemOtherCurves_Click(object sender, EventArgs e)
            {
                //With Me.frmMnt_Others
                //    myAPI.SetParent(.Handle, Me.PictureBox_Background.Handle)
                //    .WindowState = FormWindowState.Maximized
                //    .Show()
                //End With
                frmMnt_Others.MdiParent = this;
                frmMnt_Others.Show();
                frmMnt_Others.WindowState = FormWindowState.Maximized;
            }

            /// <summary>
            /// 执行Rolling操作，以进行窗口的同步滚动
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void StartRolling(object sender, EventArgs e)
            {
                frmRolling.MdiParent = this;
                frmRolling.Show();
                frmRolling.WindowState = FormWindowState.Maximized;
            }

            #endregion

            #region   ---  对话框式界面

            /// <summary>
            /// 将最终结果输出的Word中的窗口对象
            /// </summary>
            /// <remarks></remarks>
            private Diafrm_Output_Word frm_Output_Word;

            /// <summary>
            /// 将最终结果输出的Word中的
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void ExportToWord(object sender, EventArgs e)
            {
                if (frm_Output_Word == null)
                {
                    frm_Output_Word = new Diafrm_Output_Word();
                    frm_Output_Word.Load += frm_Output_Word.frm_Output_Word_Load;
                    frm_Output_Word.FormClosed += frm_Output_Word.Diafrm_Output_Word_FormClosed;
                }
                frm_Output_Word.ShowDialog();
            }

            #endregion

            #region   ---  项目文件的新建、打开、保存等

            /// <summary>
            /// 新建项目文件
            /// </summary>
            /// <remarks></remarks>
            private void NewProjectFile()
            {
                frmProjectFile = new frmProjectFile();
                frmProjectFile.FormClosed += frmProjectFile.frmProjectFile_FormClosed;
                frmProjectFile.WorkBookInProjectChanged += frmProjectFile._WorkBookInProjectChanged;
                frmProjectFile.ProjectState = ProjectState.NewProject;
                frmProjectFile.ShowDialog();
            }

            /// <summary>
            /// 打开项目文件
            /// </summary>
            /// <param name="FilePath">打开的文件的文件路径（此文件已经确定为项目后缀的文件）</param>
            /// <remarks>单纯的打开项目文件并不要求打开操作项目文件的窗口</remarks>
            private void OpenProjectFile(string FilePath)
            {
                //新开一个线程来执行打开文件的操作
                Thread thd = new Thread(new ThreadStart(this.OpenFile));
                thd.Name = "打开项目文件";
                thd.Start(FilePath);
            }

            /// <summary>
            /// 在工作者线程中执行具体的打开文件的工作
            /// </summary>
            /// <param name="FilePath"></param>
            /// <remarks></remarks>
            private void OpenFile(string FilePath)
            {
                //在主程序界面上显示出进度条
                ShowProgressBar_Marquee();
                //
                GlbApp.ProjectFile = new clsProjectFile(FilePath);
                GlbApp.ProjectFile.LoadFromXmlFile();
                GlbApp.DataBase = new ClsData_DataBase(GlbApp.ProjectFile.Contents);
                //隐藏进度条
                HideProgress("File Opened");
            }

            /// <summary>
            /// 编辑项目文件
            /// </summary>
            /// <remarks></remarks>
            private void EditProjectFile()
            {
                if (frmProjectFile == null)
                {
                    frmProjectFile = new frmProjectFile();
                    frmProjectFile.FormClosed += frmProjectFile.frmProjectFile_FormClosed;
                    frmProjectFile.WorkBookInProjectChanged += frmProjectFile._WorkBookInProjectChanged;
                }
                frmProjectFile.ProjectState = ProjectState.EditProject;
                frmProjectFile.ShowDialog();
            }

            /// <summary>
            /// 保存项目文件
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItem_SaveProject_Click(object sender, EventArgs e)
            {
                //执行保存文件的操作
                if (GlbApp.ProjectFile.FilePath == null)
                {
                    MenuItem_SaveAsProject_Click(sender, e);
                }
                else
                {
                    GlbApp.ProjectFile.SaveToXmlFile();
                }
            }

            /// <summary>
            /// 关闭项目文件
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItem_CloseProject_Click(object sender, EventArgs e)
            {
            }

            /// <summary>
            /// 另存为项目文件
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItem_SaveAsProject_Click(object sender, EventArgs e)
            {
                string FinalPathToSave = "";
                string FileExtension = AMEApplication.FileExtension;
                SaveFileDialog1.Filter = FileExtension + "文件(*" + FileExtension + ")|*" + FileExtension;
                SaveFileDialog1.DefaultExt = AMEApplication.FileExtension;
                SaveFileDialog1.AddExtension = true;
                SaveFileDialog1.OverwritePrompt = true;
                //打开对话框
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //获取选择的路径
                    FinalPathToSave = SaveFileDialog1.FileName;
                    if (FinalPathToSave.Length > 0)
                    {
                        //执行保存文件的操作
                        GlbApp.ProjectFile.FilePath = FinalPathToSave;
                        GlbApp.ProjectFile.SaveToXmlFile();
                    }
                }
            }

            #endregion

            #region   ---  数据的提取与格式化（从Excel或Word）

            public void ExtractDataFromExcel(object sender, EventArgs e)
            {
                frmDeriveData_Excel a = new frmDeriveData_Excel();
                a.ShowDialog();
            }

            public void ExtractDataFromWord(object sender, EventArgs e)
            {
                frmDeriveData_Word a = new frmDeriveData_Word();
                try
                {
                    a.ShowDialog();
                }
                catch (TargetInvocationException)
                {
                    Debug.Print("窗口关闭出错！");
                }
            }

            public void VisioToolStripMenuItem_Click(object sender, EventArgs e)
            {
                Visio_DataRecordsetLinkToShape a = new Visio_DataRecordsetLinkToShape();
                a.Show();
            }

            #endregion

            /// <summary>
            /// 在Visio平面图中绘制监测点位图
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void MenuItemDrawingPoints_Click(object sender, EventArgs e)
            {
                GlobalApplication.Application.DrawingPointsInVisio();
            }
        }
    }
}