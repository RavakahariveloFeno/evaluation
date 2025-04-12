using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace evaluation.uc
{
    public partial class Accueil : UserControl
    {
        Panel pbody;
        UserControl uc;
        Button b1,b2,b3;
        public Accueil(Panel pbody, UserControl uc, Button b1, Button b2, Button b3)
        {
            InitializeComponent();
            this.pbody = pbody;
            this.uc = uc;
            this.b1 = b1;
            this.b2 = b2;
            this.b3 = b3;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            b1.Visible = true;
            b2.Visible = true;
            b3.Visible = true;

            b1.ForeColor = Color.Orange;

            pbody.Controls.Clear();
            uc.Dock = DockStyle.Fill;
            pbody.Controls.Add(uc);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
