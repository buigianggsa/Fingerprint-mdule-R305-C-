using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Excel_12 = Microsoft.Office.Interop.Excel;

namespace MHCN_QLSV
{
    public partial class Form1 : Form
    {
        //Cái biến này để lưu lại đường dẫn ảnh của sinh viên
        private string link = "";
        // Khai báo string buff dùng cho hiển thị dữ liệu sau này.
        string InputData = String.Empty;
        private string _message;

        // Khai bao delegate SetTextCallBack voi tham so string
        delegate void SetTextCallback(string text); 

        public Form1()
        {
            InitializeComponent();
            ComPort.DataReceived += new SerialDataReceivedEventHandler(DataReceive);

        }

        //GỬi dữ liệu giữa 2 form không phải là dạng form MDI sẽ dùng đến constructor
        public Form1(string Message,int test): this()
        {
            _message = Message;
            if (test == 1)
            {
                cbbTenMon.Items.Add(_message);
            }
            if (test == 2)
            {
                cbbLop.Items.Add(_message);
            }
            if (test == 3)
            {
                cbbGiangVien.Items.Add(_message);
            }
        }

        private void load_data()
        {
            /// Cái này để binding dữ liệu vào page 2
            SqlConnection con = new SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True");
            SqlDataAdapter da = new SqlDataAdapter("Select MaVanTay as 'Mã vân tay', MaSV as 'Mã sinh viên',TenSV as 'Tên sinh viên',Lop as 'Lớp',Ngaysinh as 'Ngày sinh',GioiTinh as 'Giới tính',AnhThe as 'Ảnh thẻ' from SinhVien", con);
            System.Data.DataTable tb = new System.Data.DataTable();
            da.Fill(tb);
            dataGridView1.DataSource = tb;

            /// Cái này để binding dữ liệu vào page 1
            SqlDataAdapter da2 = new SqlDataAdapter("Select MaSV as 'Mã sinh viên',TenSV as 'Tên sinh viên',Lop as 'Lớp',Ngaysinh as 'Ngày sinh',GioiTinh as 'Giới tính' from SinhVien", con);
            System.Data.DataTable tb2 = new System.Data.DataTable();
            da2.Fill(tb2);
            dataGridView2.DataSource = tb2;

            //Cái này là cập nhật thời khóa biểu
            SqlDataAdapter da3 = new SqlDataAdapter("Select MaMon as 'Mã môn',TenMon as 'Tên môn',GioVao as 'Giờ vào',GioRa as 'Giờ ra',Lop as 'Lớp', GiangVien as 'Giảng viên' from ThoiKhoaBieu", con);
            System.Data.DataTable tb3 = new System.Data.DataTable();
            da3.Fill(tb3);
            dataGridView4.DataSource = tb3;

            //Lên ảnh của thằng đầu tiên
            if (dataGridView1.DataSource != null)
            {
                pictureBox1.Image = Image.FromFile(dataGridView1.Rows[0].Cells[6].Value.ToString());
            }

            //Lên ảnh của thằng đầu tiên
            if (dataGridView2.DataSource != null)
            {
                int Masv_temp = int.Parse(dataGridView2.Rows[0].Cells[0].Value.ToString());
                SqlDataAdapter da21 = new SqlDataAdapter("Select Anhthe from SinhVien where Masv=" + Masv_temp, con);
                System.Data.DataTable tb21 = new System.Data.DataTable();
                da21.Fill(tb21);
                dataGridView3.DataSource = tb21;
                pic_DD.Image = Image.FromFile(dataGridView3.Rows[0].Cells[0].Value.ToString());
            }
        }

        //Databinding cho cái tab2 - thông tin sinh viên
        private void DataBin1()
        {
            txt_Masv.DataBindings.Clear();
            txt_tensv.DataBindings.Clear();
            txt_Lop.DataBindings.Clear();
            txt_gt.DataBindings.Clear();
            dtpk_Ngaysinh.DataBindings.Clear();
            txt_Masv.DataBindings.Add("Text", dataGridView1.DataSource, "Mã sinh viên");
            txt_tensv.DataBindings.Add("Text", dataGridView1.DataSource, "Tên sinh viên");
            txt_Lop.DataBindings.Add("Text", dataGridView1.DataSource, "Lớp");
            txt_gt.DataBindings.Add("Text", dataGridView1.DataSource, "Giới tính");
            dtpk_Ngaysinh.DataBindings.Add("Text", dataGridView1.DataSource, "Ngày sinh");
        }

        ////Sự kiện load form chính
        private void Form1_Load(object sender, EventArgs e)
        {
            load_data();
            DataBin1();
            cnb_LoadCOm.DataSource = SerialPort.GetPortNames();
            cbb_Baurate.SelectedIndex = 4;
        }

        private void DataReceive(object obj, SerialDataReceivedEventArgs e)
        {
            InputData = ComPort.ReadExisting();
            if (InputData != String.Empty)
            {
                // textbox1 = InputData; // Ko dùng đc như thế này vì khác threads . 
                SetText(InputData); // Chính vì vậy phải sử dụng ủy quyền tại đây. Gọi delegate đã khai báo trước đó.
            }
        }
   
        // Hàm của em nó là ở đây. Đừng hỏi vì sao lại thế.
        private void SetText(string text)
        {
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText); // khởi tạo 1 delegate mới gọi đến SetText
                this.Invoke(d, new object[] { text });
            }
            else
            #region Nhận dữ liệu liên tục từ cổng COM
            {
                //DƯỚI DÒNG NÀY LÀ BẮT DỮ LIỆU XUẤT TỪ CỔNG COM
                this.textBox1.Text += text;
                int temp = int.Parse(text);
                if (temp == -2)
                    MessageBox.Show("Đã phát hiện cảm biến vân tay!");
                else
                {
                    if (((temp % 2 == 0) || (temp % 2 == 1)) && ((temp > -1) && (temp < 120)))
                    {
                        MessageBox.Show("Đã thêm một dấu vân tay mới");
                    }
                }


                //KẾT THÚC PHẦN NÀY   
            } 
            #endregion
        }

        //Chọn ảnh khi nhấn button
        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "image | *.jpg; *.png";
            DialogResult imagetest = openFileDialog1.ShowDialog();
            link = openFileDialog1.FileName.ToString();
            if(imagetest==DialogResult.OK)
                pictureBox1.Image = Image.FromFile(link);
        }

        //Đẩy ảnh lên khi cellclick
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int r = dataGridView1.CurrentCell.RowIndex;
            pictureBox1.Image = Image.FromFile(dataGridView1.Rows[r].Cells[6].Value.ToString());
            //
            DataBin1();

        }

        //Cái này phục vụ việc chọn 1 hàng dữ liệu cụa thể của sinh viên
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1_CellClick( sender, e);
        }

        //Kết nối COM và thông báo thành công
        private void btn_KetNoiCOM_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComPort.IsOpen)
                {
                    ComPort.PortName = cnb_LoadCOm.Text;
                    ComPort.Open();
                    lab_KetNoi.Text = ("Đã kết nối");
                    lab_KetNoi.ForeColor = Color.Green;
                    MessageBox.Show("Kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Cổng bạn chọn không khả dụng!","Lỗi cổng COM",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        //Ngắt kết nối COM
        private void button3_Click(object sender, EventArgs e)
        {
            ComPort.Close();
            MessageBox.Show("Ngắt kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            lab_KetNoi.Text = ("Chưa kết nối");
            lab_KetNoi.ForeColor = Color.Red;
        }

        //Tab1_Thay đổi lựa chọn hiển thị ra datagridview
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //SInh viên cả lớp
            if (cbb_ds.SelectedIndex == 1)
            {
                SqlConnection con = new SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True");
                SqlDataAdapter da = new SqlDataAdapter("Select MaSV as 'Mã sinh viên',TenSV as 'Tên sinh viên',Lop as 'Lớp',Ngaysinh as 'Ngày sinh',GioiTinh as 'Giới tính' from SinhVien", con);
                System.Data.DataTable tb = new System.Data.DataTable();
                da.Fill(tb);
                dataGridView2.DataSource = tb;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int r = dataGridView2.CurrentCell.RowIndex;
            int Masv_temp = int.Parse(dataGridView2.Rows[r].Cells[0].Value.ToString());
            SqlConnection con = new SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True");
            SqlDataAdapter da = new SqlDataAdapter("Select Anhthe from SinhVien where Masv="+ Masv_temp, con);
            System.Data.DataTable tb = new System.Data.DataTable();
            da.Fill(tb);
            dataGridView3.DataSource = tb;
            pic_DD.Image = Image.FromFile(dataGridView3.Rows[0].Cells[0].Value.ToString());
        }


        //Load môn, thời khóa biểu, Lớp, sv theo thời gian thực
        private void timer1_Tick(object sender, EventArgs e)
        {
            // Update time
            label14.Text= DateTime.Now.ToString("HH:mm:ss"); ;
            String time = DateTime.Now.ToString("HH:mm:ss");
            SqlConnection connection = new SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True; MultipleActiveResultSets=True");
            connection.Open();

            //Lấy thông tin lớp. môn theo TKB thời gian thực
            string sqlquery = ("SELECT TenMon,Lop,GioVao,GioRa FROM ThoiKhoaBieu WHERE GioVao = " + "'"+time+"'");
            SqlCommand command = new SqlCommand(sqlquery, connection);
            SqlDataReader sdr = command.ExecuteReader();
            while (sdr.Read())
            {
                txt_Mon.Text = sdr["TenMon"].ToString();
                txt_Lop_diemdanh.Text = sdr["lop"].ToString();
                txt_GioHoc.Text = sdr["GioVao"].ToString() +" - "+ sdr["GioRa"].ToString();
            }
            //Lấy số sinh viên trong lớp theo môn
            string sqlquery1 = ("SELECT Count (MaSV) as SLSV FROM SinhVien WHERE Lop = " + "'" + txt_Lop_diemdanh.Text + "'");
            SqlCommand command1 = new SqlCommand(sqlquery1, connection);
            SqlDataReader sdr1 = command1.ExecuteReader();
            while (sdr1.Read())
            {
                txt_slsv.Text = sdr1["SLSV"].ToString();
            }

        }

        //Thoát
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Bạn muốn thoát không?", "Xác nhận", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (dr == DialogResult.Yes)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        //Nút xóa sinh viên
        private void btn_Xoa_Click(object sender, EventArgs e)
        {
            DialogResult kq = MessageBox.Show("Bạn có muốn xóa sinh viên này", "Xác nhận", MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button1);
            if (kq == DialogResult.Yes)
            {
                SqlConnection con = new
                SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True");
                SqlCommand cmd = new SqlCommand("delete from Sinhvien where MaSV = '" + txt_Masv.Text + "'", con);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                DataBin1();
                load_data();
            }
        }

        //Thêm sinh viên vào CSDL
        private void btn_Them_Click(object sender, EventArgs e)
        {
            //Kiểm tra có ô nào trống không
            if ((txt_tensv.Text == "") || (txt_Masv.Text == "") || (txt_Lop.Text == "") ||
                (txt_gt.Text == "") || (dtpk_Ngaysinh.Text == "")||(link==""))
                MessageBox.Show("Vui lòng nhập đủ dữ liệu", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            else
            {
                SqlConnection con = new
                SqlConnection(@"Data Source = BDG-PC\SQLEXPRESS; Initial Catalog = QLSV_MHCN; Integrated Security = True");
                SqlCommand cm = new SqlCommand("insert into SinhVien values('" + txt_Masv.Text + "',N'" + txt_tensv.Text + "','" + txt_Lop.Text + "','" + dtpk_Ngaysinh.Text + "',N'" + txt_gt.Text + "',N'" + link + "')", con);
                con.Open();
                //TRy-catch để kiểm tra có trùng không
                try
                {
                    ComPort.Write("");
                    cm.ExecuteNonQuery();
                    con.Close();
                    DataBin1();
                    load_data();
                }
                catch
                {
                    MessageBox.Show("Dữ liệu đã tồn tại", "Lỗi", MessageBoxButtons.OK,MessageBoxIcon.Error,MessageBoxDefaultButton.Button1);
                }
            }
        }

        //SỬA
        private void btn_Sua_Click(object sender, EventArgs e)
        {

        }

        //Xuất file ra excel nhé
        private void btn_Xuat_Click(object sender, EventArgs e)
        {
            try
            {
                Excel_12.ApplicationClass excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                excel.Application.Workbooks.Add(true);

                Excel_12.Application oExcel_12 = null;                //Excel_12 Application
                Excel_12.Workbook oBook = null;                       // Excel_12 Workbook
                Excel_12.Sheets oSheetsColl = null;                   // Excel_12 Worksheets collection
                Excel_12.Worksheet oSheet = null;                     // Excel_12 Worksheet
                Excel_12.Range oRange = null;                         // Cell or Range in worksheet
                Object oMissing = System.Reflection.Missing.Value;
                oExcel_12 = new Excel_12.Application();
                oExcel_12.Visible = true;
                oExcel_12.UserControl = true;
                oBook = oExcel_12.Workbooks.Add(oMissing);
                oSheetsColl = oExcel_12.Worksheets;

                oSheet = (Excel_12.Worksheet)oSheetsColl.get_Item("Sheet1");
                //col-1 vì không xuất cot cuối: ImageLink
                //Xuất cái tiêu đề cột thôi
                oExcel_12.Columns.AutoFit();
                for (int j = 0; j < dataGridView1.Columns.Count-1; j++)
                {
                    oRange = (Excel_12.Range)oSheet.Cells[1, j + 1];
                    oRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                //Xuát dữ liệu cho cột
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count-1; j++)
                    {
                        oRange = (Excel_12.Range)oSheet.Cells[i + 2, j + 1];
                        oRange.Value2 = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // XỬ lý khi thêm thông tin về môn học, giảng viên, lớp học ở TAB ThoiKhoaBieu
        private void button4_Click(object sender, EventArgs e)
        {
            Them_Mon_Lop_GV Them = new Them_Mon_Lop_GV();
            Them.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Them_Mon_Lop_GV ThemTT = new Them_Mon_Lop_GV();
            ThemTT.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Them_Mon_Lop_GV Them = new Them_Mon_Lop_GV();
            Them.ShowDialog();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Them_Mon_Lop_GV Them = new Them_Mon_Lop_GV();
            Them.Show();
        }
    }
}
