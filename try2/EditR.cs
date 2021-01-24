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
    public partial class EditR : Form
    {
        public EditR()
        {
            InitializeComponent();
        }
        public void loadtable() // عرض البيانات
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();

            OleDbCommand cmd = new OleDbCommand("Select * from [halls$] where Room is not null", cn);
            cn.Open();

            dt.Load(cmd.ExecuteReader());

            cn.Close();


        }

        public void edittab(int cha, int row, int chro, int roo)
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            OleDbCommand cmd = new OleDbCommand("update  [halls$] set  Charis = @cha , rrow = @row , chairsRow = @chro where Room = @roo;", cn);

            cmd.Parameters.Add("id", OleDbType.Integer).Value = cha;
            cmd.Parameters.Add("ib", OleDbType.Integer).Value = row;
            cmd.Parameters.Add("ic", OleDbType.Integer).Value = chro;
            cmd.Parameters.Add("ik", OleDbType.Integer).Value = roo;
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            edittab(Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox3.Text), Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox1.Text));
            loadtable();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void EditR_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            edittab(Convert.ToInt32(textBox2.Text), Convert.ToInt32(textBox3.Text), Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox1.Text));
            loadtable();
            this.Close();
        }
    }
}
