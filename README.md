<p> 
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

namespace WindowsFormsOnThi
{
    public partial class Form1 : Form
    {
        DVChamSocEntities db = new DVChamSocEntities();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ResetListView(db.DichVus.ToList());
        }

        private void ResetListView(IEnumerable<DichVu> a)
        {
            lv_DSThuCung.Items.Clear();

            foreach (var dichVu in a)
            {
                ListViewItem item = new ListViewItem(dichVu.MaDon.Trim());

                if (int.TryParse(dichVu.CanNang.ToString(), out int canNang) && canNang > 40)
                    item.BackColor = Color.Yellow;

                item.SubItems.Add(dichVu.TenThuCung.Trim());
                item.SubItems.Add(dichVu.ChungLoai.Trim());
                item.SubItems.Add(dichVu.NgayNhan.ToString("dd/MM/yyyy"));

                lv_DSThuCung.Items.Add(item);
            }
        }

        private void Reset()
        {
            foreach (ListViewItem item in lv_DSThuCung.Items)
            {
                item.Selected = false;
            }
            txt_MaDon.Clear();
            txt_TenThuCung.Clear();
            txt_ChungLoai.Clear();
            txt_TinhTrang.Clear();
            txt_CanNang.Clear();
            txt_ChiPhiThuoc.Clear();
            txt_SoNgay.Clear();
            rad_ChuaBenh.Checked = true;
            dtp_NgayNhan.Value = DateTime.Now;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Co muon dong khong", "Thong Bao", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void rad_ChuaBenh_CheckedChanged(object sender, EventArgs e)
        {
            lbl_ChiPhiThuoc.Visible = true;
            txt_ChiPhiThuoc.Visible = true;
            lbl_SoNgay.Visible = false;
            txt_SoNgay.Visible = false;
            txt_SoNgay.Clear();
        }

        private void rad_ChamSoc_CheckedChanged(object sender, EventArgs e)
        {
            lbl_ChiPhiThuoc.Visible = false;
            txt_ChiPhiThuoc.Visible = false;
            txt_ChiPhiThuoc.Clear();
            lbl_SoNgay.Visible = true;
            txt_SoNgay.Visible = true;
        }

        private void btn_Them_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private bool KiemTra()
        {
            if (string.IsNullOrEmpty(txt_MaDon.Text) || string.IsNullOrEmpty(txt_TenThuCung.Text) || string.IsNullOrEmpty(txt_ChungLoai.Text) || string.IsNullOrEmpty(txt_CanNang.Text) || string.IsNullOrEmpty(txt_TinhTrang.Text) || (string.IsNullOrEmpty(txt_ChiPhiThuoc.Text) && rad_ChuaBenh.Checked) || (string.IsNullOrEmpty(txt_SoNgay.Text) && rad_ChamSoc.Checked))
            {
                MessageBox.Show("Không được để trống thông tin.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (dtp_NgayNhan.Value.Date > DateTime.Now.Date)
            {
                MessageBox.Show("Ngày ... không được ... hơn ngày hiện tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!int.TryParse(txt_CanNang.Text, out int canNang) || canNang <= 0)
            {
                MessageBox.Show("... phải là một số lớn hơn 0.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if ((!int.TryParse(txt_SoNgay.Text, out int soNgay) || soNgay <= 0) && rad_ChamSoc.Checked)
            {
                MessageBox.Show("... phải là một số lớn hơn 0.", "lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if ((!float.TryParse(txt_ChiPhiThuoc.Text, out float tienThuoc) || tienThuoc <= 0) && rad_ChuaBenh.Checked)
            {
                MessageBox.Show("... phải là một số lớn hơn 0.", "lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //if (txt_DienThoai.Text.Any(c => !char.IsDigit(c)) || txt_DienThoai.Text.Length != 10)
            //{
            //    MessageBox.Show("Số điện thoại phải là một dãy số có 10 chữ số.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}

            return true;
        }

        private void btn_Luu_Click(object sender, EventArgs e)
        {
            if (KiemTra())
            {
                double pt;
                bool c = db.DichVus.Any(d => d.MaDon == txt_MaDon.Text);
                if (c)
                    MessageBox.Show("Mã đơn đã tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    if (rad_ChuaBenh.Checked)
                        pt = double.Parse(txt_ChiPhiThuoc.Text) + 100000;
                    else
                        pt = double.Parse(txt_SoNgay.Text) * 200000;

                    db.DichVus.Add(new DichVu() { MaDon = txt_MaDon.Text, TenThuCung = txt_TenThuCung.Text, ChungLoai = txt_ChungLoai.Text, CanNang = int.Parse(txt_CanNang.Text), TinhTrang = txt_TinhTrang.Text, NgayNhan = dtp_NgayNhan.Value.Date, LoaiDV = rad_ChuaBenh.Checked ? true : false, Phi = pt });
                    db.SaveChanges();

                    ListViewItem i = new ListViewItem(txt_MaDon.Text);
                    i.SubItems.Add(txt_TenThuCung.Text);
                    //i.SubItems.Add(rad_TinhCam.Checked ? rad_TinhCam.Text : rad_HanhDong.Text);
                    i.SubItems.Add(txt_ChungLoai.Text);
                    i.SubItems.Add(dtp_NgayNhan.Value.ToString("dd/MM/yyyy"));

                    lv_DSThuCung.Items.Add(i);

                    i.Selected = true;

                    if (int.TryParse(txt_CanNang.Text, out int canNang) && canNang > 40)
                        i.BackColor = Color.Yellow;
                }
            }
        }

        private void lv_DSThuCung_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lv_DSThuCung.SelectedItems.Count > 0)
            {
                string maDon = lv_DSThuCung.SelectedItems[0].SubItems[0].Text;

                txt_MaDon.Enabled = false;

                var donDV = db.DichVus.SingleOrDefault(d => d.MaDon == maDon);

                if (donDV != null)
                {
                    txt_MaDon.Text = donDV.MaDon;
                    txt_TenThuCung.Text = donDV.TenThuCung;
                    txt_ChungLoai.Text = donDV.ChungLoai;
                    txt_TinhTrang.Text = donDV.TinhTrang;
                    txt_CanNang.Text = donDV.CanNang.ToString();
                    dtp_NgayNhan.Value = donDV.NgayNhan;

                    if (donDV.LoaiDV == true)
                    {
                        rad_ChuaBenh.Checked = true;
                        double t = (double)donDV.Phi - 100000;
                        txt_ChiPhiThuoc.Text = t.ToString();
                    }
                    else
                    {
                        rad_ChamSoc.Checked = true;
                        double t = (double)donDV.Phi / 200000;
                        txt_SoNgay.Text = t.ToString();
                    }
                }
            }
            else
            {
                txt_MaDon.Enabled = true;
                Reset();
            }
        }

        private void btn_Xoa_Click(object sender, EventArgs e)
        {
            if (lv_DSThuCung.SelectedItems.Count > 0)
            {
                if (MessageBox.Show("Bạn có chắc chắn muốn xóa ... đã chọn?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    int index = lv_DSThuCung.Items.IndexOf(lv_DSThuCung.SelectedItems[0]);

                    string maDon = lv_DSThuCung.SelectedItems[0].SubItems[0].Text;

                    DichVu dichVu = db.DichVus.Where(p => p.MaDon.Trim() == maDon).SingleOrDefault();
                    db.DichVus.Remove(dichVu);
                    db.SaveChanges();

                    lv_DSThuCung.Items.Remove(lv_DSThuCung.SelectedItems[0]);

                    if (lv_DSThuCung.Items.Count > 0)
                    {
                        if (index < lv_DSThuCung.Items.Count)
                            lv_DSThuCung.Items[index].Selected = true;
                        else
                            lv_DSThuCung.Items[index - 1].Selected = true;
                    }

                    if (lv_DSThuCung.Items.Count == 0)
                    {
                        Reset();
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một ... để xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_Sua_Click(object sender, EventArgs e)
        {
            if (lv_DSThuCung.SelectedItems.Count > 0)
            {
                int index = lv_DSThuCung.Items.IndexOf(lv_DSThuCung.SelectedItems[0]);

                string maDon = lv_DSThuCung.SelectedItems[0].SubItems[0].Text;
                DichVu dichVu = db.DichVus.Where(p => p.MaDon.Trim() == maDon).SingleOrDefault();

                double pt;
                if (rad_ChuaBenh.Checked)
                    pt = double.Parse(txt_ChiPhiThuoc.Text) + 100000;
                else
                    pt = double.Parse(txt_SoNgay.Text) * 200000;
                if (KiemTra())
                {
                    dichVu.TenThuCung = txt_TenThuCung.Text;
                    dichVu.ChungLoai = txt_ChungLoai.Text;
                    dichVu.TinhTrang = txt_ChungLoai.Text;
                    dichVu.NgayNhan = dtp_NgayNhan.Value;
                    dichVu.LoaiDV = rad_ChuaBenh.Checked ? true : false;
                    dichVu.CanNang = int.Parse(txt_CanNang.Text);
                    dichVu.Phi = pt;
                    db.SaveChanges();
                    txt_MaDon.Enabled = true;
                    ResetListView(db.DichVus.ToList());
                    Reset();
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một ... để sửa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_SapXep_Click(object sender, EventArgs e)
        {
            try
            {
                var s = db.DichVus.OrderByDescending(p => p.NgayNhan).ThenBy(p => p.CanNang).ToList();
                ResetListView(s);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_ThongKe_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    var list = from ldv in db.LoaiDichVus
            //               join dv in db.DichVus
            //               on ldv.LoaiDV equals dv.LoaiDV into gj
            //               from s in gj.DefaultIfEmpty()
            //               group s by new { ldv.LoaiDV } into g
            //               select new
            //               {
            //                   LoaiDV = g.Key.LoaiDV,
            //                   TongSoLuong = g.Count(p => p != null),
            //                   TongDoanhThu = g.Sum(p => p != null ? p.Phi : 0)
            //               };


            //    string message = "Thống kê theo loại ... :\n\n";

            //    foreach (var s in list)
            //    {
            //        if (s.LoaiDV)
            //            message += $"---Chữa bệnh---\n";
            //        else
            //            message += $"---Chăm sóc hộ---\n";

            //        message += $"Số lượng: {s.TongSoLuong}\n";
            //        if (s.TongDoanhThu == 0)
            //            message += $"Doanh thu: {s.TongDoanhThu}\n";
            //        else
            //            message += $"Doanh thu: {s.TongDoanhThu:#,#}\n\n";
            //    }

            //    MessageBox.Show(message, "Thống kê", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet excelWs = excelWb.Worksheets[1];

            Excel.Range excelRange = excelWs.Cells[1, 1];
            excelRange.Font.Size = 16;
            excelRange.Font.Bold = true;
            excelRange.Font.Color = Color.Blue;
            excelRange.Value = "...";

            var catalogs = db.DichVus.Select(c => new { maDon = c.MaDon, tenThuCung = c.TenThuCung }).ToList();
            int row = 2;
            foreach (var c in catalogs)
            {
                excelWs.Range["A" + row].Font.Bold = true; excelWs.Range["A" + row].Value = c.maDon;
                row++;
                // Lấy sp theo danh mục
                var products = from p in db.DichVus where p.MaDon == c.maDon select p;
                foreach (var p in products)
                {
                    excelWs.Range["A" + row].Value = p.MaDon;
                    excelWs.Range["B" + row].ColumnWidth = 50;
                    excelWs.Range["B" + row].Value = p.TenThuCung;
                    excelWs.Range["C" + row].Value = p.Phi;
                    row++;
                }
            }

            excelWs.Name = "DanhMuc"; excelWb.Activate();
            // Luu file
            SaveFileDialog saveFileDialog = new SaveFileDialog(); if (saveFileDialog.ShowDialog() == DialogResult.OK) excelWb.SaveAs(saveFileDialog.FileName);
            excelApp.Quit();
        }
    }
}
</p>
