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
            //dataGridView1.DataSource = _dataService.getAll();
            dt = _dataService.getAll();
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;

                // Masquer les colonnes non désirées
                if (dataGridView1.Columns.Contains("PV1"))
                    dataGridView1.Columns["PV1"].Visible = false;
                if (dataGridView1.Columns.Contains("PV2"))
                    dataGridView1.Columns["PV2"].Visible = false;
                if (dataGridView1.Columns.Contains("PV3"))
                    dataGridView1.Columns["PV3"].Visible = false;
                if (dataGridView1.Columns.Contains("Montant_commission"))
                    dataGridView1.Columns["Montant_commission"].Visible = false;
                if (dataGridView1.Columns.Contains("Observations"))
                    dataGridView1.Columns["Observations"].Visible = false;
                if (dataGridView1.Columns.Contains("Montant_commission"))
                    dataGridView1.Columns["Montant_commission"].Visible = false;
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
