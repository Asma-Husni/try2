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
    public partial class EditM : Form
    {
        public EditM()
        {
            InitializeComponent();
        }
        public void edittab(string cha, string row, string chro, string roo)
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            OleDbCommand cmd = new OleDbCommand("update  [halls$] set  materials = @cha , period = @row , Date = @chro where Days = @roo;", cn);

            cmd.Parameters.Add("id", OleDbType.VarChar).Value = cha;
            cmd.Parameters.Add("ib", OleDbType.VarChar).Value = row;
            cmd.Parameters.Add("ic", OleDbType.VarChar).Value = chro;
            cmd.Parameters.Add("ik", OleDbType.VarChar).Value = roo;
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();
        }
        public void loadtable() // عرض البيانات
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();

            OleDbCommand cmd = new OleDbCommand("Select * from [pro$] ", cn);
            cn.Open();

            dt.Load(cmd.ExecuteReader());

            cn.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string te1, te2, te3, te4;
            te1 = textBox1.Text; te2 = textBox2.Text; te3 = textBox3.Text; te4 = textBox4.Text;
            edittab(te2, te3, te4, te1);
            loadtable();
            this.Close();
        }

        private void EditM_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string te1, te2, te3, te4;
            te1 = textBox1.Text; te2 = textBox2.Text; te3 = textBox3.Text; te4 = textBox4.Text;
            edittab(te2, te3, te4, te1);
            loadtable();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

