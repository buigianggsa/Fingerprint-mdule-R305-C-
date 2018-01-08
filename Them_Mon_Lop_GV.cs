using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static MHCN_QLSV.Form1;

namespace MHCN_QLSV
{
    public partial class Them_Mon_Lop_GV : Form
    {
        public Them_Mon_Lop_GV()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {

            Form1 Child = new Form1(txt_ThemMon2.Text,1);
            Child.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_themLop_Click(object sender, EventArgs e)
        {
            Form1 Child = new Form1(txt_ThemMon2.Text, 2);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form1 Child = new Form1(txt_ThemMon2.Text, 3);
        }
    }
}
