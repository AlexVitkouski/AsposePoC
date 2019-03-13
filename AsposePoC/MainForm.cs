using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsposePoC
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            button.Text = "Testing...";
            var testRunner = new TestRunner(Common.PathTSrcFolder, pictureBox1);
            testRunner.Run();
            MessageBox.Show("Done.");
            button.Text = "Start";
        }
    }
}
