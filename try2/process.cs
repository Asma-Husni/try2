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
    public partial class process : Form
    {
        static string code, name, id;
        static char latter;
        public process()
        {
            InitializeComponent();
        }
        public void loadtable(string pa) // عرض البيانات
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();

            OleDbCommand cmd = new OleDbCommand("Select * from [" + pa + "$]", cn);
            cn.Open();

            dt.Load(cmd.ExecuteReader());

            cn.Close();

            dataGridView1.DataSource = dt;
        }

        public void number(string pa) // عرض البيانات
        {
            string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            OleDbConnection cn = new OleDbConnection(con);

            DataTable dt = new DataTable();
            OleDbDataAdapter cmd = new OleDbDataAdapter("SELECT COUNT(*) FROM[" + pa + "$]", cn);
            cn.Open();
            DataSet ds = new DataSet();
            cmd.Fill(ds);
            dt = ds.Tables[0];
            label1.Text = dt.Rows[0][0].ToString();
            cn.Close();
        }
        private void process_Load(object sender, EventArgs e)
        {
            //string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\pp.xlsx ;  Extended Properties='Excel 12.0; HDR = YES' ");
            //OleDbConnection cn = new OleDbConnection(con);
            //DataTable dt = new DataTable();
            //OleDbDataAdapter cmd = new OleDbDataAdapter("Select * from [pro$] ", cn);
            //cn.Open();
            //cmd.Fill(dt);
            //comboBox1.DataSource = dt;
            //comboBox1.DisplayMember = "Days";
            //comboBox2.DataSource = dt;
            //comboBox2.DisplayMember = "period";


            SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
            SqlCommand cmm = new SqlCommand();
            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select * from Finsch ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "daays";
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "priod";
            cn.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //string selected = comboBox1.Text.ToString();
            //string selected2 = comboBox2.Text.ToString();
            ////string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\pp.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            ////OleDbConnection cn = new OleDbConnection(con);
            ////DataTable dt = new DataTable();
            ////OleDbDataAdapter cmd = new OleDbDataAdapter("SELECT materials FROM [pro$] where Days = [" + comboBox1.Text + "] &&  period = [" + comboBox2.Text + "]", cn);
            ////cn.Open();
            ////DataSet ds = new DataSet();
            ////cmd.Fill(ds);
            //SqlConnection cn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");
            //SqlCommand cmm = new SqlCommand();
            //cn.Open();
            //SqlDataAdapter da = new SqlDataAdapter("SELECT mate FROM Finsch where days = '" + comboBox1.Text.ToString() + "' and period= '" + comboBox2.Text.ToString() + "' ", cn);
            //DataTable dt = new DataTable();
            //DataSet ds = new DataSet();
            //da.Fill(dt);
            //dt = ds.Tables[0];
            //label1.Text = dt.Rows[0][0].ToString();
            //cn.Close();
            //label1.Visible = true;
            //label3.Visible = true;



            SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
            //SqlCommand cmm = new SqlCommand();
            cn.Open();
            //SqlDataAdapter da = new SqlDataAdapter("select code from subject where name = '" + comboBox1.Text.ToString() + "' ", cn);
            //int result = (int)cmd.ExecuteScalar();
            //cn.Close();
            using (SqlCommand command = new SqlCommand("SELECT mate FROM Finsch where daays = '" + comboBox1.Text.ToString() + "' and priod= '" + comboBox2.Text.ToString() + "' ", cn))
            {
                //
                // Invoke ExecuteReader method.
                //
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    label1.Text = reader.GetString(0);
                }
                label1.Visible = true;
                label3.Visible = true;
            }
            cn.Close();
            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select count(ID) as عدد ,ss.mate as المادة ,code from student,Finsch ss where daays = '" + comboBox1.Text.ToString() + "' and priod = '" + comboBox2.Text.ToString() + "' and code = (SELECT code  FROM subject where ss.mate = name) group by ss.mate ,code ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            cn.Close();

            int sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                sum += Convert.ToInt32(dataGridView1.Rows[i].Cells["عدد"].Value);
            }
            MessageBox.Show(Convert.ToString(sum));
            label5.Text = sum.ToString();


            cn.Open();
            SqlDataAdapter da1 = new SqlDataAdapter("select name , Num_colum, Num_chire from halls where sum >='" + label5.Text + "' ", cn);
            DataTable dt1 = new DataTable();
            da1.Fill(dt1);
            dataGridView2.DataSource = dt1;
            cn.Close();

            //string conn = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\stu.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //OleDbConnection cnn = new OleDbConnection(conn);
            //DataTable dtt = new DataTable();
            //OleDbDataAdapter cmdd = new OleDbDataAdapter("SELECT COUNT(*) FROM[" + label1.Text + "$]", cnn);
            //cnn.Open();
            //DataSet dss = new DataSet();
            //cmd.Fill(dss);
            //dtt = ds.Tables[0];
            //label2.Text = dtt.Rows[0][0].ToString();
            //cnn.Close();
            ////su.Visible = true;
            ////su.Text = selected;
            ////loadtable(selected);
            ////number(selected);
            //string conn1 = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\pp.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
            //OleDbConnection cnn1 = new OleDbConnection(conn1);
            //DataTable dt1t = new DataTable();
            //OleDbDataAdapter cmdd1 = new OleDbDataAdapter("SELECT sum(Charis) FROM[halls$]", cnn1);
            //cnn.Open();
            //DataSet dss1 = new DataSet();
            //cmd.Fill(dss);
            //dt1t = dss1.Tables[0];
            //label3.Text = dt1t.Rows[0][0].ToString();
            //cnn1.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");

           label1.Text = dataGridView2.SelectedRows[0].Cells["name"].Value.ToString();
            int chair = Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["Num_chire"].Value);
            int row = Convert.ToInt32(dataGridView2.SelectedRows[0].Cells["Num_colum"].Value);
            MessageBox.Show(Convert.ToString(row));
            for (int i = 1; i <= row; i++)

            {
                switch (i)
                {
                    case 1:
                        latter = 'A';
                        break;
                    case 2:
                        latter = 'B';
                        break;
                    case 3:
                        latter = 'C';
                        break;
                    case 4:
                        latter = 'D';
                        break;
                    case 5:
                        latter = 'E';
                        break;
                    case 6:
                        latter = 'F';
                        break;
                    case 7:
                        latter = 'G';
                        break;
                }
                if (i % 2 != 0)
                {
                    code = dataGridView1.Rows[0].Cells["code"].Value.ToString();
                    MessageBox.Show(Convert.ToString(code));

                    cn.Open();
                  
                        SqlCommand command = new SqlCommand("SELECT top(1) s.id ,s.name from student s WHERE s.code='" + code + "' and  NOT EXISTS( SELECT TOP 1 NULL FROM Process WHERE s.id = Stu_id and Subject_Code = '" + code + "')", cn);
                        SqlDataReader reader = command.ExecuteReader();
                        bool istt;
                    try
                    {
                        while (reader.Read())
                        {
                            id = reader.GetString(0);
                            name = reader.GetString(1);
                            MessageBox.Show(id);
                            MessageBox.Show(name);
                        }
                    }
                    catch (Exception ex)
                    {
                        for (int j = 0; j <= 10; j++)
                        {
                            cn.Close();
                            cn.Open();
                            SqlCommand cmd = new SqlCommand("INSERT INTO [dbo].[process]([Stu_id],[Stu_name],[Subject_Code],[Chair],[Halls])VALUES('" + id + "' ,'" + name + "','" + code + "','" + latter+j + "','" + label1.Text + "')", cn);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Done");
                            
                        }
                    }
                    cn.Close();
                }
                else
                {
                    code = dataGridView1.Rows[1].Cells["code"].Value.ToString();

                    cn.Open();

                    SqlCommand command = new SqlCommand("SELECT top(1) s.id ,s.name from student s WHERE s.code='" + code + "' and  NOT EXISTS( SELECT TOP 1 NULL FROM Process WHERE s.id = Stu_id and Subject_Code = '" + code + "' )", cn);

                    SqlDataReader reader = command.ExecuteReader();
                    try
                    {
                        while (reader.Read())
                        {
                            id = reader.GetString(0);
                            name = reader.GetString(1);
                            MessageBox.Show(id);
                            MessageBox.Show(name);
                        }
                    }
                    catch (Exception ex)
                    {

                        cn.Close();
                        cn.Open();
                        for (int k = 0; k <= 10; k++)
                        {
                            SqlCommand cmd = new SqlCommand("INSERT INTO [dbo].[process]([Stu_id],[Stu_name],[Subject_Code],[Chair],[Halls])VALUES('" + id + "' ,'" + name + "','" + code + "','" + latter+k + "','" + label1.Text + "')", cn);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Done");
                        }
                        cn.Close();
                    }
                }
            }
        }

    }
}







