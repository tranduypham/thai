using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Word = Microsoft.Office.Interop.Word;

namespace Test_Report
{
    public partial class Report : Form
    {

        private readonly string Temp = System.Windows.Forms.Application.StartupPath + "/Word/Reports1.docx";
        public Report()
        {
            InitializeComponent();
        }
        private void AllowNumberOnly(Object sender, KeyPressEventArgs e)
        {
            //Function ngăn ko cho nhập chữ, sẽ đc sử dụng trong sự kiện bên dưới
            //Kiểm tra xem ký tự nhập phải là số hoặc ký tự điểu khiển như esc hoặc Enter
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;// handled = ko cho nhập ký tự đó
            }
        }
        private void lay_du_lieu(DataGridView dt, int index)
        {
            //Cái này để lấy dữ liệu từ từng bẳng DataGridView
            //Dữ liệu sau đó sẽ được lưu ở object GiangDay
            //Giang dạy tao để ở dạng static
            //lưu vào Class static GiangDay dưới dạng string
            GiangDay.dulieu[index] = "";
            try
            {
                int i, j;
                int col = dt.Columns.Count - 1;
                int row = dt.Rows.Count - 1;

                for (i = 0; i < row; i++)
                {
                    for (j = 0; j <= col; j++)
                    {
                        if (dt[j, i].Value == null)
                        {
                            GiangDay.dulieu[index] += "0" + "$";
                        }
                        else
                        {
                            GiangDay.dulieu[index] += "0" + "$";
                        }
                        //Cái đoạn này t làm để tính toán giờ dạy ở cuối sau khi ấn nút send
                        if (j == 5)//Cột 5 là số tiết phải dạy
                        {
                            GiangDay.phai_giang_A += Convert.ToInt32(dt[j, i].Value == null ? 0 : dt[j, i].Value);
                        }
                        if (j == 6)//Cột 6 là số tiết thực dạy
                        {
                            GiangDay.thuc_giang_A += Convert.ToInt32(dt[j, i].Value == null ? 0 : dt[j, i].Value);
                        }
                    }
                    GiangDay.dulieu[index] += "\n";
                }
            }catch(Exception e)
            {
                //Cái này để phòng hờ có lỗi gì nó hiện lên
                MessageBox.Show(e.Message);
            }
            

        }
        private void lay_basic_infor()
        {
            //Cái này cũng là lẫy dữ liệu
            //Nhưng là lấy mấy cái họ tên các thứ
            //Cũng lưu vào Class static GiangDay dưới dạng string
            GiangDay.khoa = cbKhoa.Text==null?"0": cbKhoa.Text;
            GiangDay.boMon = cbBoMon_0.Text == null ? "0" : cbBoMon_0.Text;
            GiangDay.Day = numDay.Value.ToString();
            GiangDay.Month = numMonth.Value.ToString();
            GiangDay.Year = numYear.Value.ToString();
            GiangDay.HoTen = TbHoTen.Text == null ? "0" : TbHoTen.Text;
            GiangDay.namSinh = cbNamSinh.Text == null ? "0" : cbNamSinh.Text;
            GiangDay.chucVu = cbChucVu.Text == null ? "0" : cbChucVu.Text;
            GiangDay.luong = numLuong.Value.ToString() == null ? "0" : numLuong.Value.ToString();
            GiangDay.hocHam = TbHocHam.Text == null ? "0" : TbHocHam.Text;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Cái nay nhập ngày tháng năm hiện tại vào chỗ Hà nội ngày
            DateTime dt = DateTime.Now;
            numDay.Value = dt.Day;
            numMonth.Value = dt.Month;
            numYear.Value = dt.Year;

            //Chỉ lả gen ra các năm tuổi thôi 
            for(int i = 1800; i <= dt.Year; i++)
            {
                cbNamSinh.Items.Add(i);
            }
            
        }


        private void button2_Click(object sender, EventArgs e)
        {
            //Mấy dòng này nó đọc dữ liệu từ Grid table rồi lưu vào giảng dạy
            lay_basic_infor();
            lay_du_lieu(dataGridView1, 1);
            lay_du_lieu(dataGridView2, 2);
            lay_du_lieu(dataGridView4, 3);
            lay_du_lieu(dataGridView3, 4);
            
            this.Hide();

            this.Show();
        }


        private void numLuong_Leave(object sender, EventArgs e)
        {
            
        }

        private void cbBoMon_0_Leave(object sender, EventArgs e)
        {
            
        }

        //Control Showing : cái sự kiện này ko hiểu nó là gì nhưng trên mạng nó hướng dẫn thế
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Function ngăn cột thứ 0,5,6 nhập chữ
            if(dataGridView1.CurrentCell.ColumnIndex == 6|| dataGridView1.CurrentCell.ColumnIndex == 5 || dataGridView1.CurrentCell.ColumnIndex == 0)
            {
                e.Control.KeyPress += AllowNumberOnly;//cái này để ép ko cho nhập chữ vào

            }
            else
            {
                //Sau nhiều lần test nhận ra, ở trên có 3 ô cấm nhập chữ
                //Suy ra ở dưới có 3 lần trừ (đ hiểu tại sao đâu đừng hỏi)
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi 
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
            }
        }


        private void dataGridView2_EditingControlShowing_1(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Function ngăn cột thứ 0,5,6 nhập chữ
            if (dataGridView2.CurrentCell.ColumnIndex == 6|| dataGridView2.CurrentCell.ColumnIndex == 5|| dataGridView2.CurrentCell.ColumnIndex == 0)
            {
                e.Control.KeyPress += AllowNumberOnly;//cái này để ép ko cho nhập chữ vào

            }
            else
            {
                //Sau nhiều lần test nhận ra, ở trên có 3 ô cấm nhập chữ
                //Suy ra ở dưới có 3 lần trừ (đ hiểu tại sao đâu đừng hỏi)
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi 
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
            }
        }

        private void dataGridView3_EditingControlShowing_2(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Function ngăn cột thứ 0,5,6 nhập chữ
            if (dataGridView3.CurrentCell.ColumnIndex == 6|| dataGridView2.CurrentCell.ColumnIndex == 5|| dataGridView2.CurrentCell.ColumnIndex == 0)
            {
                e.Control.KeyPress += AllowNumberOnly;//cái này để ép ko cho nhập chữ vào

            }
            else
            {
                //Sau nhiều lần test nhận ra, ở trên có 3 ô cấm nhập chữ
                //Suy ra ở dưới có 3 lần trừ (đ hiểu tại sao đâu đừng hỏi)
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi 
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
            }

        }

        private void dataGridView4_EditingControlShowing_1(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Function ngăn cột thứ 0,5,6 nhập chữ
            if (dataGridView4.CurrentCell.ColumnIndex == 6|| dataGridView2.CurrentCell.ColumnIndex == 5|| dataGridView2.CurrentCell.ColumnIndex == 0)
            {
                e.Control.KeyPress += AllowNumberOnly;//cái này để ép ko cho nhập chữ vào

            }
            else
            {
                //Sau nhiều lần test nhận ra, ở trên có 3 ô cấm nhập chữ
                //Suy ra ở dưới có 3 lần trừ (đ hiểu tại sao đâu đừng hỏi)
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi 
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
                e.Control.KeyPress -= AllowNumberOnly;//cái này để bỏ sự kiện ko cho nhập kia đi
            }
        }

        private void label31_Click(object sender, EventArgs e)
        {

        }


        private void pictureBox5_MouseLeave(object sender, EventArgs e)
        {
            pictureBox5.ImageLocation = @"close.png";
        }

        private void pictureBox5_MouseEnter(object sender, EventArgs e)
        {
            pictureBox5.ImageLocation = @"close_1.png";
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.Close();
        }



        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void lay_du_lieu_2(DataGridView dt, int index)
        {
            GiangDay.dulieu[index] = "";

            int i, j;
            int col = dt.Columns.Count - 1;
            int row = dt.Rows.Count - 1;

            for (i = 0; i < row; i++)
            {
                for (j = 0; j <= col; j++)
                {
                    if (dt[j, i].Value == null)
                    {
                        GiangDay.dulieu[index] += "0" + "$";
                    }
                    else
                    {
                        GiangDay.dulieu[index] += "0" + "$";
                    }
                    //tính tổng số tiết chuyển đổi cảu bên hướng dẫn tốt nghiệp(DataGridView4)
                    if (index == 4 && j == 4)
                    {
                        GiangDay.thuc_giang_B += Convert.ToInt32(dt[j, i].Value == null ? 0 : dt[j, i].Value);
                    }

                }
                GiangDay.dulieu[index] += "\n";
            }

        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            lay_basic_infor();
            lay_du_lieu(dataGridView1, 1);
            lay_du_lieu(dataGridView2, 2);
            lay_du_lieu(dataGridView4, 3);
            lay_du_lieu(dataGridView3, 4);

            lay_du_lieu_2(dataGridView5, 5);
            lay_du_lieu_2(dataGridView6, 6);
            TbTongSoTiet.Text = (GiangDay.thuc_giang_A + GiangDay.thuc_giang_B).ToString();
            TbSoTietGiang.Text = GiangDay.phai_giang_A.ToString();
            TbSoGioChuaHT.Text = ((GiangDay.phai_giang_A - GiangDay.thuc_giang_A) < 0 ? 0 : (GiangDay.phai_giang_A - GiangDay.thuc_giang_A)).ToString();
            TbTongSoTietVuot.Text = ((GiangDay.phai_giang_A - GiangDay.thuc_giang_A) >= 0 ? 0 : (GiangDay.phai_giang_A - GiangDay.thuc_giang_A) * -1).ToString();
            MessageBox.Show("Thành công");
        }

        private void dataGridView5_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView5.CurrentCell.ColumnIndex == 4 || dataGridView5.CurrentCell.ColumnIndex == 3 || dataGridView5.CurrentCell.ColumnIndex == 0)
            {
                e.Control.KeyPress += AllowNumberOnly;//Nếu để Một mình cái ko có else này ngăn cả bảng nhập chữ luôn
            }
            else
            {
                e.Control.KeyPress -= AllowNumberOnly;
                e.Control.KeyPress -= AllowNumberOnly;
                e.Control.KeyPress -= AllowNumberOnly;
            }
        }

        private void dataGridView6_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Function ngăn cột thứ 6 nhập chữ
            if (dataGridView6.CurrentCell.ColumnIndex == 2)
            {
                //Phaỉ có cả hai cái dưới mới hoạt động
                e.Control.KeyPress += AllowNumberOnly;//Một mình cái này ngăn cả bảng nhập chữ luôn
            }
            else
            {
                e.Control.KeyPress -= AllowNumberOnly;
            }
        }

        //StackOverFlow
        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }


            ///////////////////////////////////////////////////////////
            ///Cod cuar Thái
            var hoTen = TbHoTen  == null ? "null": TbHoTen.Text;
            var khoa = cbKhoa.Text;
            var toBoMon = cbBoMon_0.Text;
            var ngay = numDay.Text;
            var thang = numMonth.Text;
            var nam = numYear.Text;
            var namHoc = TbNam.Text;
            var namSinh = cbNamSinh.Text;
            var chucVu = cbChucVu.Text;
            var luongThuc = numLuong.Text;
            var hocHam = TbHocHam.Text;
            var gridbang = dataGridView1.Rows[0].Cells[0].Value.ToString();

            var wordApp = new Microsoft.Office.Interop.Word.Application();
            var wordDocument = wordApp.Documents.Open(Temp);
            GiangDay connect = new GiangDay();
            connect.ReplaceWordStub("{Khoa}", khoa, wordDocument);
            connect.ReplaceWordStub("{Tobomon}", toBoMon, wordDocument);
            connect.ReplaceWordStub("{d}", ngay, wordDocument);
            connect.ReplaceWordStub("{m}", thang, wordDocument);
            connect.ReplaceWordStub("{y}", nam, wordDocument);
            connect.ReplaceWordStub("{Year}", namHoc, wordDocument);
            connect.ReplaceWordStub("{Hoten}", hoTen, wordDocument);
            connect.ReplaceWordStub("{Birthday}", namSinh, wordDocument);
            connect.ReplaceWordStub("{Chucvu}", chucVu, wordDocument);
            connect.ReplaceWordStub("{Luongthucnhan}", luongThuc, wordDocument);
            connect.ReplaceWordStub("{Hocham}", hocHam, wordDocument);
            connect.ReplaceWordStub("{bang}", gridbang, wordDocument);

            //Không thể lưu tên như ý muốn
            // TbHoTen.Text+@"_"+TbNam.Text.Trim()+
            string output = @"/Baocao.docx";
            wordDocument.SaveAs(Application.StartupPath + output);
            wordApp.Documents.Open(Application.StartupPath + output);
        }
    }
}
