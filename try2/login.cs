using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace try2
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void login_Load(object sender, EventArgs e)
        {
            this.textBox2.AutoSize = false;
            this.textBox2.Size = new System.Drawing.Size(180, 25);
            this.textBox1.AutoSize = false;
            this.textBox1.Size = new System.Drawing.Size(180, 25);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("يرجوا ادخال اسم او كلمة المرور");
                textBox1.Clear();
                textBox2.Clear();
            }
            else
            {
             
                if (textBox2.Text == "it" &&  textBox1.Text == "1234")
                {
                    login n = new login();
                    n.Hide();
                    main m = new main();
                    m.Show();
                    this.Hide();

                }
                else
                {
                    MessageBox.Show("لا يوجد اسم المستخدم او كلمة المرور غير صحيحة ");
                    textBox1.Clear();
                    textBox2.Clear();

                }
            }
        }
    }
}
