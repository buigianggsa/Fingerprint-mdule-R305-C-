using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MHCN_QLSV
{
    public partial class fm_Login : Form
    {
        public fm_Login()
        {
            InitializeComponent();
        }

        private void btn_Thoat_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Bạn muốn thoát không?", "Xác nhận", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }
        }


        private void txt_DangNhap_TextChanged(object sender, EventArgs e)
        {
            if ((txt_MatKhau.Text != "") && (txt_DangNhap.Text != ""))
            {
                btn_DangNhap.Enabled = true;
            }

        }

        private void txt_DangNhap_Leave(object sender, EventArgs e)
        {
            if (txt_DangNhap.Text == "")
                errorProvider1.SetError(txt_DangNhap, "Lỗi");
            else errorProvider1.Clear();
        }

        private void txt_MatKhau_Leave(object sender, EventArgs e)
        {
            if (txt_DangNhap.Text == "")
                errorProvider1.SetError(txt_MatKhau, "Lỗi");
            else errorProvider1.Clear();
        }

        private void btn_DangNhap_Click(object sender, EventArgs e)
        {
            if((txt_DangNhap.Text=="Admin") && (txt_MatKhau.Text == "Admin")){
                Form1 fr1 = new Form1();
                fr1.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Nhập lại!","Lỗi",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }
    }
}
