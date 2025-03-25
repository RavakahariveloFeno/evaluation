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
        public Data()
        {
            InitializeComponent();
            _dataService = new service.DataService();
            getAll();
        }

        private void getAll()
        {
            dataGridView1.DataSource = _dataService.getAll();
        }
    }
}
