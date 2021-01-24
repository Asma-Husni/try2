using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace try2
{
    public partial class teacher : Form
    {
        public teacher()
        {
            InitializeComponent();
        }
        public void loadtable() // عرض البيانات
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();

            OleDbCommand cmd = new OleDbCommand("Select * from [ove$] ", cn);
            cn.Open();

            dt.Load(cmd.ExecuteReader());

            cn.Close();
            dataGridView1.DataSource = dt;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadtable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void teacher_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            loadtable();
        }
    }
}

