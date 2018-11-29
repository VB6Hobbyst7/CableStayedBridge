using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace CableStayedBridge.Test
{
    /// <summary>
    /// Interaction logic for WindowMultiThread.xaml
    /// </summary>
    public partial class WindowMultiThread : Window
    {
        public WindowMultiThread()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Thread trd = new Thread(haltThread);
            trd.Start();
        }

        private void haltThread()
        {
            DateTime t1 = DateTime.Now;
            DateTime t2 = DateTime.Now;
            while ((t2 - t1).Seconds < 5)
            {
                t2 = DateTime.Now;
                
            }
        }
    }
}
