using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using CableStayedBridge.Test;
using System.Windows.Controls.Ribbon;


namespace CableStayedBridge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private GlobalApplication m_globalApp;

        private static readonly object locker = new object();
        private static MainWindow _instance = null;
        public static MainWindow UniqueInstance
        {
            get
            {
                //只有当第一次创建实例的时候再执行lock语句
                if (_instance == null)
                {
                    //当一个线程执行的时候，会先检测locker对象是否是锁定状态，
                    //如果不是则会对locker对象锁定，如果是则该线程就会挂起等待
                    //locker对象解锁,lock语句执行运行完之后会对该locker对象解锁
                    lock (locker)
                    {
                        if (_instance == null)
                        {
                            _instance = new MainWindow();
                        }
                    }
                }
                return _instance;
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            m_globalApp = GlobalApplication.UniqueInstance;
            //

            //CableStayedBridge.Test.Window_Ribbon wr = new Window_Ribbon();
            //wr.Show();


            // 监听Ribbon的收缩与扩展
            System.ComponentModel.DependencyPropertyDescriptor.FromProperty(Ribbon.IsMinimizedProperty, typeof(Ribbon)).AddValueChanged(Ribbon_CSB, (o, args) => { AdjustWorkingArea(Ribbon_CSB.IsMinimized); });
            // 禁止双击时Ribbon的收缩
            // System.ComponentModel.DependencyPropertyDescriptor.FromProperty(Ribbon.IsMinimizedProperty, typeof(Ribbon))    .AddValueChanged(Ribbon_CSB, (o, args) => Ribbon_CSB.IsMinimized = false);


            //var winForm = new Form1();
            //winForm.TopLevel = false; // 最关键的一步
            //// f.FormBorderStyle = FormBorderStyle.None;
            //FormsHost.Child = winForm;  // WindowsFormsHost类为WPF中Window中的控件，用来在WPF中装载WinForm控件或窗体
            //  Window w1 = new CableStayedBridge.Test.RRWParameters();
            // w1.Show();

            // this.Close();
            //

        }

        /// <summary>
        /// 在Ribbon的收缩与扩展时，调整窗口的工作区域
        /// </summary>
        private void AdjustWorkingArea(bool isMinimized)
        {
            double ribbonTabHeight = 92;// ribbonTab.Height; RibTab_CSB.Height;

            // System.Windows.MessageBox.Show(ribbonTabHeight.ToString());


            var margin = Grid_MainWindowContent.Margin;
            // Grid_MainWindowContent
            if (isMinimized)
            {
                margin.Top -= ribbonTabHeight;
            }
            else
            {
                margin.Top += ribbonTabHeight;

            }
            Grid_MainWindowContent.Margin = margin;
        }
    
        /// <summary>
        /// 切换到建模参数输入界面
        /// </summary>
        private void SwitchToModeling()
        {

        }

        private void ribtn_Modeling_Click(object sender, RoutedEventArgs e)
        {
            PreModeling frmPreModeling = new PreModeling();
            frmPreModeling.ShowDialog();
        }
    }
}
