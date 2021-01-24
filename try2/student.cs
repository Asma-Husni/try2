using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace try2
{

    public partial class student : Form
    {
        string code;
        SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
        //Excel.Application xlApp = new Excel.Application();
        //Excel.Workbook xlWorkBook;
        //Excel.Worksheet xlWorkSheet;
        //string sFileName;
        //int iRow, iCol = 2;

        public object ComboBox1 { get; private set; }

        public student()
        {
            InitializeComponent();
        }
        public void edittab(string pa, string cha, string roo)
        {
            //string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 12.0; HDR = YES' ");
            //OleDbConnection cn = new OleDbConnection(con);

            //OleDbCommand cmd = new OleDbCommand("update  ["+pa+"$] set  name= @cha where no= @roo;", cn);

            //cmd.Parameters.Add("id", OleDbType.VarChar).Value = cha;
            //cmd.Parameters.Add("ik", OleDbType.VarChar).Value = roo;
            //cn.Open();
            //cmd.ExecuteNonQuery();
            //cn.Close();
        }
        public void insertable(string pa, string id, string set)//الإدخال
        {
            //string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx  ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //OleDbConnection cn = new OleDbConnection(con);

            //OleDbCommand cmd = new OleDbCommand("insert into ["+pa+"$] values (@id,@set)", cn);

            //cmd.Parameters.Add("id", OleDbType.VarChar).Value = id;
            //cmd.Parameters.Add("ib", OleDbType.VarChar).Value = set;


            //cn.Open();

            //cmd.ExecuteNonQuery();
            //cn.Close();

        }

        public void loadtable(string pa) // عرض البيانات
        {
            //    string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //    OleDbConnection cn = new OleDbConnection(con);

            //    DataTable dt = new DataTable();

            //    OleDbCommand cmd = new OleDbCommand("Select * from ["+pa+"$]" , cn);
            //    cn.Open();

            //    dt.Load(cmd.ExecuteReader());

            //    cn.Close();

            //    dataGridView1.DataSource = dt;
        }

        public void number(string pa) // عرض البيانات
        {
            //string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //OleDbConnection cn = new OleDbConnection(con);

            //DataTable dt = new DataTable();
            //OleDbDataAdapter cmd = new OleDbDataAdapter("SELECT COUNT(*) FROM[" + pa + "$]", cn);
            //cn.Open();
            //DataSet ds = new DataSet();
            //cmd.Fill(ds);
            //dt=ds.Tables[0];
            //label1.Text=dt.Rows[0][0].ToString();
            //cn.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            panel1.Visible = false;
            panel3.Visible = false;
            label1.Visible = false;
            label3.Visible = false;
            su.Visible = false;
            SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
            SqlCommand cmm = new SqlCommand();
            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select name from subject ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "name";
            cn.Close();

        }

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    string selected =comboBox1.Text;
        //    label1.Visible = true;
        //    label3.Visible = true;
        //    su.Visible = true;
        //    su.Text = selected;
        //    loadtable(selected);
        //    number(selected);


        //}

        private void button1_Click(object sender, EventArgs e)
        {
            string selected = comboBox1.Text;
            insertable(selected, textBox3.Text, textBox4.Text);
            loadtable(selected);
            su.Text = selected;
            loadtable(selected);
            number(selected);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            panel3.Visible = true;
            textBox7.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string selected = comboBox1.Text;
            insertable(selected, textBox3.Text, textBox4.Text);
            loadtable(selected);
            su.Text = selected;
            loadtable(selected);
            number(selected);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string selected = comboBox1.Text;
            edittab(selected, textBox7.Text, textBox8.Text);
            su.Text = selected;
            loadtable(selected);
            number(selected);
            loadtable(selected);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string selected = comboBox1.Text.ToString();
            OleDbConnection con;
            OleDbDataAdapter da;
            DataTable dt;
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "ALL Files |*.*| Excel Files |*.XLSE";
            if (op.ShowDialog() == DialogResult.OK)
            {
                con = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0 ; Data Source=" + op.FileName + "; Extended Properties = Excel 12.0 ");
                da = new OleDbDataAdapter("select * from [su$]", con);
                dt = new DataTable();
                da.Fill(dt);
                this.dataGridView1.DataSource = dt;
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            //cn.Open();
            //SqlCommand cmd = new SqlCommand("insert into student(ID,Student_name) values(@id,@name)");
            //cmd.Parameters.Add("@ID", SqlDbType.Int);
            //cmd.Parameters.Add("@tudent_name", SqlDbType.VarChar);
            //int n = dataGridView1.Rows.Count - 1;
            //for (int i = 0;i== n; i++)
            //{
            //             command.Parameters(0).Value = dgvServerConfig.Rows(i).Cells(0).Value
            // command.Parameters(1).Value = dgvServerConfig.Rows(i).Cells(1).Value
            //}

            //string StrQuery;
            ////try
            ////{
            ///*using (*/
            //SqlConnection conn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");
            //{
            //    using 




            //comm.Connection = conn;
            //            conn.Open();
            //SqlCommand comm = new SqlCommand();
            //cn.Open();


            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    SqlCommand cmd = new SqlCommand("insert into student(ID,Student_name,code) values ("
            //        + dataGridView1.Rows[i].Cells["ID"].Value + " ,'" + dataGridView1.Rows[i].Cells["name"].Value + "','" + dataGridView1.Rows[i].Cells["ID"].Value + "')", cn);
            //    cmd.ExecuteNonQuery();
            //    MessageBox.Show("Done");

                //StrQuery = @"INSERT INTO student VALUES ("
                //                + dataGridView1.Rows[i].Cells["ID"].Value + ", "
                //               + dataGridView1.Rows[i].Cells["ID"].Value + ", " 
                //               +dataGridView1.Rows[i].Cells["ID"].Value + ");";
                //comm.CommandText = StrQuery;
                //            comm.ExecuteNonQuery();
            ////}
            ////cn.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //string selected = comboBox1.Text.ToString()
            SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
            //SqlCommand cmm = new SqlCommand();
            cn.Open();
            //SqlDataAdapter da = new SqlDataAdapter("select code from subject where name = '" + comboBox1.Text.ToString() + "' ", cn);
            //int result = (int)cmd.ExecuteScalar();
            //cn.Close();
            using (SqlCommand command = new SqlCommand("select code from subject where name = '" + comboBox1.Text.ToString() + "' ", cn))
            {
                //
                // Invoke ExecuteReader method.
                //
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    code = reader.GetString(0);
                }
            }
            cn.Close();
            //SqlCommand cmm = new SqlCommand();
            cn.Open();
            ////SqlCommand comm = new SqlCommand();


            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                SqlCommand cmd = new SqlCommand("insert into student(id,name,code) values ('" + dataGridView1.Rows[i].Cells["id"].Value + "' ,'" + dataGridView1.Rows[i].Cells["name"].Value + "','" + code + "')", cn);
                cmd.ExecuteNonQuery();
            }
            MessageBox.Show("Done");
            cn.Close();
        }
       
        //}
        //catch (Exception ex)
        //{
        //    MessageBox.Show("hi");
        //}
    }
}
    





