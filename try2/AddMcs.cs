using System;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace try2
{
    public partial class AddMcs : Form
    {
        SqlConnection cn = new SqlConnection("Data Source=PC-WORLD;Initial Catalog=exsam;Integrated Security=True");
        public AddMcs()
        {
            InitializeComponent();
        }
        //public void insertable(string id, string set, string co, string setco)//الإدخال
        //{
        //    string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\pp.xlsx ;  Extended Properties='Excel 8.0; HDR = YES' ");
        //    OleDbConnection cn = new OleDbConnection(con);

        //    OleDbCommand cmd = new OleDbCommand("insert into [pro$] values (@id,@ib,@ic,@ik)", cn);

        //    cmd.Parameters.Add("id", OleDbType.VarChar).Value = id;
        //    cmd.Parameters.Add("ib", OleDbType.VarChar).Value = set;
        //    cmd.Parameters.Add("ic", OleDbType.VarChar).Value = co;
        //    cmd.Parameters.Add("ik", OleDbType.VarChar).Value = setco;

        //    cn.Open();

        //    cmd.ExecuteNonQuery();
        //    cn.Close();

        //}

        //public void loadtable() // عرض البيانات
        //{
        //    string con = (@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = D:\pp.xlsx ;  Extended Properties='Excel 12.0; HDR = YES' ");
        //    OleDbConnection cn = new OleDbConnection(con);

        //    DataTable dt = new DataTable();

        //    OleDbCommand cmd = new OleDbCommand("Select * from [pro$] ", cn);
        //    cn.Open();

        //    dt.Load(cmd.ExecuteReader());

        //    cn.Close();

        //}
        //private void button1_Click(object sender, EventArgs e)
        //{


        //    //string te1, te2, te3, te4;
        //    //te1 = textBox1.Text; te2 = textBox2.Text; te3 = textBox3.Text; te4 = textBox4.Text;
        //    //insertable(te1, te2, te3, te4);
        //    //loadtable();

        //}

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AddMcs_Load(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
            exsam m = new exsam();
            m.Show();
            this.Hide();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            cn.Open();
            SqlCommand cmd = new SqlCommand("insert into finsch(daays,mate,priod,daate) values ( '" + comboBox1.Text + "','" + comboBox2.Text + "','" + comboBox3.Text + "','" + Convert.ToDateTime(dateTimePicker1.Value.ToString()) + "')", cn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Done");
            cn.Close();


            //string te1, te2, te3, te4;
            //te1 = textBox1.Text; te2 = textBox2.Text; te3 = textBox3.Text; te4 = textBox4.Text;
            //insertable(te1, te2, te3, te4);
            //loadtable();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

