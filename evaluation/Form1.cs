using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace evaluation
{
    public partial class Form1 : Form
    {
        uc.Data dataControls = new uc.Data();
        uc.Evaluation evaluationControls = new uc.Evaluation();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pbody.Controls.Clear();
            dataControls.Dock = DockStyle.Fill;
            pbody.Controls.Add(dataControls);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pbody.Controls.Clear();
            dataControls.Dock = DockStyle.Fill;
            pbody.Controls.Add(dataControls);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pbody.Controls.Clear();
            evaluationControls.Dock = DockStyle.Top;
            pbody.Controls.Add(evaluationControls);
        }
    }
}
