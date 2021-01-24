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
using System.Data.SqlClient;

namespace try2
{
    public partial class AddR : Form
    {
        SqlConnection cn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");
        public AddR()
        {
            InitializeComponent();

        }






        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AddR_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {


            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int sm = Convert.ToInt32(textBox3.Text) * Convert.ToInt32(textBox2.Text);
            SqlCommand cmd = new SqlCommand("insert into halls(name,Num_colum,Num_chire,sum) values (" + Convert.ToInt32(textBox1.Text) + " ,'" + Convert.ToInt32(textBox2.Text) + "','" + Convert.ToInt32(textBox3.Text) + "','" + sm + "')", cn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Done");
            cn.Close();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            this.Close();

        }
    }
}

