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
    public partial class exsam : Form
    {
        SqlConnection cn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");


        public exsam()
        {
            InitializeComponent();
        }


        private void exsam_Load(object sender, EventArgs e)
        {
            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select day,subject,date,period from scheduel ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comboBox6.DataSource = dt;
            comboBox6.DisplayMember = "day";
            comboBox5.DataSource = dt;
            comboBox5.DisplayMember = "subject";
            comboBox4.DataSource = dt;
            comboBox4.DisplayMember = "period";
            cn.Close();

   
          


        }





        private void button1_Click(object sender, EventArgs e)
        {



        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //this.dataGridView1.Rows[0].Cells[0].Value = comboBox1.Text;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void button2_Click_1(object sender, EventArgs e)

        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }


        private void button3_Click(object sender, EventArgs e)
        {
            //  textBox5.Show();
            //  label1.Show();
            //   button3.Show();

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var add = new AddMcs();
            add.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //var edit = new EditM();
            //edit.Show();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            var add = new AddMcs();
            add.Show();
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {

            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select *from scheduel ", cn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            cn.Close();
        }
    }
}




