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
        uc.Accueil accueilControls;
        uc.Data dataControls = new uc.Data();
        uc.Evaluation evaluationControls = new uc.Evaluation();
        uc.Settings settingsControls = new uc.Settings();
        public Form1()
        {
            InitializeComponent();
            accueilControls = new uc.Accueil(pbody,dataControls,button1,button2,button3);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pbody.Controls.Clear();
            accueilControls.Dock = DockStyle.Fill;
            pbody.Controls.Add(accueilControls);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dataControls = new uc.Data();
                // Changer le curseur en mode "attente"
                Cursor = Cursors.WaitCursor;

                button1.ForeColor = Color.Orange;
                button2.ForeColor = Color.White;
                button3.ForeColor = Color.White;

                pbody.Controls.Clear();
                dataControls.Dock = DockStyle.Fill;
                pbody.Controls.Add(dataControls);
            }
            finally
            {
                // Remettre le curseur en mode "normal"
                Cursor = Cursors.Default;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // Changer le curseur en mode "attente"
                Cursor = Cursors.WaitCursor;

                // Modifier la couleur des boutons
                button1.ForeColor = Color.White;
                button2.ForeColor = Color.Orange;
                button3.ForeColor = Color.White;

                // Charger le contrôle
                evaluationControls = new uc.Evaluation();
                pbody.Controls.Clear();
                evaluationControls.Dock = DockStyle.Top;
                pbody.Controls.Add(evaluationControls);
            }
            finally
            {
                // Remettre le curseur en mode "normal"
                Cursor = Cursors.Default;
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // Changer le curseur en mode "attente"
                Cursor = Cursors.WaitCursor;

                button1.ForeColor = Color.White;
                button2.ForeColor = Color.White;
                button3.ForeColor = Color.Orange;

                settingsControls = new uc.Settings();
                pbody.Controls.Clear();
                settingsControls.Dock = DockStyle.Fill;
                pbody.Controls.Add(settingsControls);
            }
            finally
            {
                // Remettre le curseur en mode "normal"
                Cursor = Cursors.Default;
            }
        }
    }
}
