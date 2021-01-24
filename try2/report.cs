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
    public partial class report : Form
    {
        public report()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            main f2 = new main();
            f2.Show();
            this.Close();
        }
    }
}
