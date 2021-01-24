using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace try2
{
    public partial class subjet : Form
    {

        SqlConnection cn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");
        SqlCommand cmm = new SqlCommand();
        public subjet()
        {
            InitializeComponent();
        }

        public void loadtable() // عرض البيانات
        {
            //string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\subject.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //OleDbConnection cn = new OleDbConnection(con);

            //DataTable dt = new DataTable();

            //OleDbCommand cmd = new OleDbCommand("Select * from [su$] where code <> null", cn);
            //cn.Open();

            //dt.Load(cmd.ExecuteReader());

            //cn.Close();
          

            SqlDataAdapter da = new SqlDataAdapter("select code as رقم, name as اسم from subject ", cn);
            cn.Open();
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            cn.Close();

        }

        public void loadde() // عرض البيانات
        {
             string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\subject.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();

            OleDbCommand cmd = new OleDbCommand("Select * from [su$] where code <> null", cn);
            cn.Open();

            dt.Load(cmd.ExecuteReader());

            cn.Close();

            dataGridView1.DataSource = dt;
        }
        public void insertable(string id, string set)//الإدخال
        {
            cn.Open();
            SqlCommand cmd = new SqlCommand("insert into subject(code,name) values(@id,@set)", cn);
            cmd.ExecuteNonQuery();

            MessageBox.Show(" تم الاضافة");

            cn.Close();

        }

      

            public void edittab(string cha, string roo)
            {
             
               //# OleDbCommand cmd = new OleDbCommand("update  [su$] set  name= @cha where code = @roo;", cn);

               // cmd.Parameters.Add("id", OleDbType.VarChar).Value = cha;
               // cmd.Parameters.Add("ik", OleDbType.VarChar).Value = roo;
               // cn.Open();
               // cmd.ExecuteNonQuery();
               // cn.Close();
            }
            


        private void Form1_Load(object sender, EventArgs e)
        {
            loadtable();
            panel1.Visible = false;
            panel2.Visible = false;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            //panel1.Visible = true;

          AddSubject f2 = new AddSubject();
            f2.Show();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            loadtable();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            edittab(textBox2.Text, textBox1.Text);
            loadtable();



        }
        public void deltab(string roo)
        {



            cmm.Connection = cn;
            cmm.CommandType = CommandType.Text;
            cmm.CommandText = "delete from subject where name ='" + dataGridView1.CurrentRow.Cells[1].Value + "' ";
            cn.Open();

            cmm.ExecuteNonQuery();

            MessageBox.Show("تم حذف  ");

            cn.Close();
            SqlDataAdapter da = new SqlDataAdapter("select code as رقم, name as اسم from subject ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            cn.Close();



        }
        private void button5_Click(object sender, EventArgs e)
        {
            deltab(textBox1.Text);
            loadde();
        }
       

      
    

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
           main f2 = new main();
            f2.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            panel2.Visible =true;
            textBox3.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //insertable(textBox1.Text, textBox2.Text);
            cn.Open();
            SqlCommand cmd = new SqlCommand("insert into subject(code,name) values('" + textBox1.Text + "','" + textBox2.Text + "')", cn);
            cmd.ExecuteNonQuery();

            MessageBox.Show(" تم الاضافة");

            cn.Close();
            loadtable();
            panel1.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
        }

        private void button3_Click_1(object sender, EventArgs e)
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
