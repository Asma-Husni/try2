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
using System.Data.SqlClient;
namespace try2
{
    public partial class rooms : Form
    {
        public rooms()
        {
            InitializeComponent();
        }
        //public void loadtable() // عرض البيانات
        //{
        //    string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = d:\pp.xlsx;  Extended Properties='Excel 12.0; HDR = YES' ");
        //    OleDbConnection cn = new OleDbConnection(con);

        //    DataTable dt = new DataTable();

        //    OleDbCommand cmd = new OleDbCommand("Select * from [halls$] where Room is not null", cn);
        //    cn.Open();

        //    dt.Load(cmd.ExecuteReader());

        //    cn.Close();

        //    dataGridView1.DataSource = dt;
        //}
        //public void deltab(int roo)
        //{
        //    string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source =d:\pp.xlsx; ;  Extended Properties='Excel 12.0; HDR = YES' ");
        //    OleDbConnection cn = new OleDbConnection(con);

        //    OleDbCommand cmd = new OleDbCommand("update  [halls$] set Room  =null, Charis = null , rrow = null , chairsRow = null where Room = @roo;", cn);
        //    cmd.Parameters.Add("ik", OleDbType.Integer).Value = roo;
        //    cn.Open();
        //    cmd.ExecuteNonQuery();
        //    cn.Close();
        //}

        private void rooms_Load(object sender, EventArgs e)
        {
            //cn.Open();
            //SqlDataAdapter da = new SqlDataAdapter("select id from Rooms ", cn);
            //DataTable dt = new DataTable();
            //da.Fill(dt);
            //comboBox1.DataSource = dt;
            //comboBox1.DisplayMember = "id";

            //cn.Close();
            //comboBox1.Hide();
            //button5.Hide();
            //label1.Hide();

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //loadtable();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox1.Show();
            label1.Show();
            button5.Show();


        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var add = new AddR();
            add.Show();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            var edit = new EditR();
            edit.Show();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            //textBox1.Hide();
            //button5.Hide();
            //label1.Hide();

            //deltab(Convert.ToInt32(textBox1.Text));
            //loadtable();
        }

       

        private void button6_Click(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //loadtable();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var edit = new EditR();
            edit.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Show();
            label1.Show();
            button5.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Hide();
            button5.Hide();
            label1.Hide();

            //deltab(Convert.ToInt32(textBox1.Text));
            //loadtable();
        }
    }
}
