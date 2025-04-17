using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace evaluation.uc
{
    public partial class Data : UserControl
    {
        service.DataService _dataService;
        private DataTable dt;
        public Data()
        {
            InitializeComponent();
            dataGridView1.DefaultCellStyle.Font = new Font("Arial", 11);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 11, FontStyle.Bold);

            _dataService = new service.DataService();
            getAll();
        }

        private void getAll()
        {
            dt = _dataService.getAll();

            if (dt.Rows.Count > 0)
            {
                // Affecter la source au DataGridView
                dataGridView1.DataSource = dt;

                // Liste des colonnes à modifier
                string[] colonnesAMultiplier = new string[]
                {
                    "Taux_d_absentéisme",
                    "R/O en CI",
                    "Qualité",
                    "Taux de réussite",
                    "Quizz",
                    "T2B_Personnalisation",
                    "T2B_Solution",
                    "Appréciation_managériale"
                };

                // Appliquer un format personnalisé sur les colonnes concernées
                foreach (string colName in colonnesAMultiplier)
                {
                    if (dataGridView1.Columns.Contains(colName))
                    {
                        dataGridView1.Columns[colName].DefaultCellStyle.Format = "P2"; // pourcentage avec 2 décimales
                    }
                }

                // Masquer les colonnes non désirées
                string[] colonnesAmasquer = new string[]
                {
                    "PV1", "PV2", "PV3", "Montant_commission", "Observations"
                };

                foreach (string col in colonnesAmasquer)
                {
                    if (dataGridView1.Columns.Contains(col))
                        dataGridView1.Columns[col].Visible = false;
                }
            }
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (dt != null)
            {
                DataView dv = dt.DefaultView;
                dv.RowFilter = "Trigramme LIKE '%" + textBox1.Text + "%'";
                dataGridView1.DataSource = dv;
            }
        }
    }
}
