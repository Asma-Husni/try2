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

namespace try2
{
    public partial class AddSubject : Form
    {
        SqlConnection cn = new SqlConnection("Data Source=TOSHIBA;Initial Catalog=exsam;Integrated Security=True");
        SqlCommand cmm = new SqlCommand();
        public AddSubject()
        {
            InitializeComponent();
        }
    

        private void button1_Click(object sender, EventArgs e)
        {
            cn.Open();
            SqlCommand cmd = new SqlCommand("insert into subject(code,name) values('" + code.Text + "','" + name.Text + "')", cn);
            cmd.ExecuteNonQuery();

            MessageBox.Show(" تم الاضافة");

            cn.Close();
            this.Close(); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
