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
    public partial class main : Form
    {
        public main()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
           exsam m = new exsam();
            m.Show();
            this.Hide();
        }

        private void button8_Click(object sender, EventArgs e)
        {
           main n = new main();
            n.Hide();
            subjet m = new subjet();
            m.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
            student m = new student();
            m.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
           process m = new process();
            m.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
            report m = new report();
            m.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
           rooms m = new rooms();
            m.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            main n = new main();
            n.Hide();
            teacher m = new teacher();
            m.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
