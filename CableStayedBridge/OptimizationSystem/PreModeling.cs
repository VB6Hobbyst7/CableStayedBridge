using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CableStayedBridge
{

        public partial class PreModeling : Form
    {
        public static FileSystemWatcher watcher;
        public PreModeling()
        {
            InitializeComponent();
            //
            Dock = DockStyle.Fill;     
        }
        
        private void btn_UserControl_Click(object sender, EventArgs e)
        {

        }
    }
}
