using DataMacroWi.Controller;
using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataMacroWi
{
    public partial class Form1 : Form
    {
        public static Form1 _Form1;

        public Form1()
        {
            InitializeComponent();
            _Form1 = this;

        }



        private void btnGDPTheoGiaHienHanh_Click(object sender, EventArgs e)
        {
            FillDataController gDPNamController = new FillDataController();


            //    gDPNamController.FillData(7, "Year", "Value", "Tỷ", "Du lieu vi mo/San Luong/GDP Nam.txt");
            //    gDPNamController.FillData(5, "Quarter", "Value", "Tỷ", "Du lieu vi mo/San Luong/GDP Hien Hanh Quy.txt");
            //    gDPNamController.FillData(6, "Quarter", "Value", "Tỷ", "Du lieu vi mo/San Luong/GDP So Sanh Quy.txt");
            //    gDPNamController.FillData(6, "Quarter", "YoY", "%", "Du lieu vi mo/San Luong/GDP So Sanh Quy YoY.txt");




            //gDPNamController.FillData(4, "Year", "Value", "Điểm", "Du lieu vi mo/San Luong/PMI.txt");

            //gDPNamController.FillData(3, "Month", "Value", "Điểm", "Du lieu vi mo/Tieu Dung/CPI/CPI_Thang.txt");
            //gDPNamController.FillData(3, "Month", "MoM", "%", "Du lieu vi mo/Tieu Dung/CPI/CPI_Thang_MoM.txt");
            //gDPNamController.FillData(3, "Month", "YoY", "%", "Du lieu vi mo/Tieu Dung/CPI/CPI_Thang_YoY.txt");
            //gDPNamController.FillData(3, "Month", "YoY Ave", "%", "Du lieu vi mo/Tieu Dung/CPI/CPI_Thang_YoYAve.txt");

            //gDPNamController.FillData(1, "Month", "Value", "Tỷ", "Du lieu vi mo/Tieu Dung/Ban Le HH DV/Value.txt");
            //gDPNamController.FillData(1, "Month", "MoM", "%", "Du lieu vi mo/Tieu Dung/Ban Le HH DV/MoM.txt");
            //gDPNamController.FillData(1, "Month", "YoY", "%", "Du lieu vi mo/Tieu Dung/Ban Le HH DV/YoY.txt");

            //gDPNamController.FillData(8, "Month", "Value", "", "Du lieu vi mo/Dau Tu/Dang ky kinh doanh/Value.txt");
            //gDPNamController.FillData(8, "Month", "MoM", "%", "Du lieu vi mo/Dau Tu/Dang ky kinh doanh/MoM.txt");
            //gDPNamController.FillData(8, "Month", "YoY", "%", "Du lieu vi mo/Dau Tu/Dang ky kinh doanh/YoY.txt");

            //gDPNamController.FillData(9, "Month", "Value", "Nghìn tỷ", "Du lieu vi mo/Dau Tu/Von dau tu phat trien xa hoi/Value.txt");
            //gDPNamController.FillData(9, "Month", "QoQ", "%", "Du lieu vi mo/Dau Tu/Von dau tu phat trien xa hoi/QoQ.txt");
            //gDPNamController.FillData(9, "Month", "YoY", "%", "Du lieu vi mo/Dau Tu/Von dau tu phat trien xa hoi/YoY.txt");

            //gDPNamController.FillData(10, "Month", "Value", "Tỷ", "Du lieu vi mo/Dau Tu/Von dau tu tu nsnn/Value.txt");
            //gDPNamController.FillData(10, "Month", "MoM", "%", "Du lieu vi mo/Dau Tu/Von dau tu tu nsnn/MoM.txt");
            //gDPNamController.FillData(10, "Month", "YoY", "%", "Du lieu vi mo/Dau Tu/Von dau tu tu nsnn/YoY.txt");

            //gDPNamController.FillData(11, "Month", "Value", "Điểm", "Du lieu vi mo/San Xuat/IIP/Value.txt");
            //gDPNamController.FillData(11, "Month", "YoY", "%", "Du lieu vi mo/San Xuat/IIP/YoY.txt");

            //gDPNamController.FillData(12, "Month", "Value", "", "Du lieu vi mo/San Xuat/San pham cong nghiep/Value.txt");
            //gDPNamController.FillData(12, "Month", "MoM", "%", "Du lieu vi mo/San Xuat/San pham cong nghiep/MoM.txt");
            //gDPNamController.FillData(12, "Month", "YoY", "%", "Du lieu vi mo/San Xuat/San pham cong nghiep/YoY.txt");


            //gDPNamController.FillData(13, "Quarter", "QoQ", "%", "Du lieu vi mo/San Xuat/Chi so gia van tai, kho bai/QoQ.txt");
            //gDPNamController.FillData(13, "Quarter", "YoY", "%", "Du lieu vi mo/San Xuat/Chi so gia van tai, kho bai/YoY.txt");

            //gDPNamController.FillData(14, "Quarter", "QoQ", "%", "Du lieu vi mo/San Xuat/Chi so gia dau vao san xuat/QoQ.txt");
            //gDPNamController.FillData(14, "Quarter", "YoY", "%", "Du lieu vi mo/San Xuat/Chi so gia dau vao san xuat/YoY.txt");

            //gDPNamController.FillData(15, "Quarter", "QoQ", "%", "Du lieu vi mo/San Xuat/PPI/QoQ.txt");
            //gDPNamController.FillData(15, "Quarter", "YoY", "%", "Du lieu vi mo/San Xuat/PPI/YoY.txt");

            //gDPNamController.FillData(16, "Month", "Value", "", "Du lieu vi mo/FDI/Dau tu truc tiep tu nuoc ngoai/Value.txt");
            //gDPNamController.FillData(16, "Month", "MoM", "%", "Du lieu vi mo/FDI/Dau tu truc tiep tu nuoc ngoai/MoM.txt");
            //gDPNamController.FillData(16, "Month", "YoY", "%", "Du lieu vi mo/FDI/Dau tu truc tiep tu nuoc ngoai/YoY.txt");

            //gDPNamController.FillData(17, "Month", "Value", "Triệu USD", "Du lieu vi mo/FDI/FDI Dang ky theo linh vuc/Value.txt");
            //gDPNamController.FillData(17, "Month", "MoM", "%", "Du lieu vi mo/FDI/FDI Dang ky theo linh vuc/MoM.txt");
            //gDPNamController.FillData(17, "Month", "YoY", "%", "Du lieu vi mo/FDI/FDI Dang ky theo linh vuc/YoY.txt");

            //gDPNamController.FillData(18, "Month", "Value", "Triệu USD", "Du lieu vi mo/FDI/FDI Dang ky theo quoc gia/Value.txt");
            //gDPNamController.FillData(18, "Month", "MoM", "%", "Du lieu vi mo/FDI/FDI Dang ky theo quoc gia/MoM.txt");
            //gDPNamController.FillData(18, "Month", "YoY", "%", "Du lieu vi mo/FDI/FDI Dang ky theo quoc gia/YoY.txt");

            //gDPNamController.FillData(19, "Month", "Value", "Triệu USD", "Du lieu vi mo/FDI/FDI Dang Ky Theo Tinh Thanh/Value.txt");
            //gDPNamController.FillData(19, "Month", "MoM", "%", "Du lieu vi mo/FDI/FDI Dang Ky Theo Tinh Thanh/MoM.txt");
            //gDPNamController.FillData(19, "Month", "YoY", "%", "Du lieu vi mo/FDI/FDI Dang Ky Theo Tinh Thanh/YoY.txt");

            //gDPNamController.FillData(20, "Month", "Value", "Triệu USD", "Du lieu vi mo/FDI/FDI Luy Ke Chua Thuc hien/Value.txt");
            //gDPNamController.FillData(20, "Month", "MoM", "%", "Du lieu vi mo/FDI/FDI Luy Ke Chua Thuc hien/MoM.txt");
            //gDPNamController.FillData(20, "Month", "YoY", "%", "Du lieu vi mo/FDI/FDI Luy Ke Chua Thuc hien/YoY.txt");



            //gDPNamController.FillData(21, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Xuat nhap khau/Value.txt");
            //gDPNamController.FillData(21, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Xuat nhap khau/MoM.txt");
            //gDPNamController.FillData(21, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Xuat nhap khau/YoY.txt");

            //gDPNamController.FillData(22, "Month", "Value", "USD/KG", "Du lieu vi mo/Xuat nhap khau/Gia Xuat Nhap Khau/Value.txt");


            //gDPNamController.FillData(23, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo mat hang/Value.txt");
            //gDPNamController.FillData(23, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo mat hang/MoM.txt");
            //gDPNamController.FillData(23, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo mat hang/YoY.txt");

            //gDPNamController.FillData(24, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Nhap Khau Theo mat Hang/Value.txt");
            //gDPNamController.FillData(24, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Nhap Khau Theo mat Hang/MoM.txt");
            //gDPNamController.FillData(24, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Nhap Khau Theo mat Hang/YoY.txt");

            //gDPNamController.FillData(25, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Xuat khau mat hang- quoc gia/Value.txt");
            //gDPNamController.FillData(25, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau mat hang- quoc gia/MoM.txt");
            //gDPNamController.FillData(25, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau mat hang- quoc gia/YoY.txt");

            //gDPNamController.FillData(26, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Nhap khau mat hang - quoc gia/Value.txt");
            //gDPNamController.FillData(26, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Nhap khau mat hang - quoc gia/MoM.txt");
            //gDPNamController.FillData(26, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Nhap khau mat hang - quoc gia/YoY.txt");

            //gDPNamController.FillData(27, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo quoc gia - mat hang/Value.txt");
            //gDPNamController.FillData(27, "Month", "MoM", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo quoc gia - mat hang/MoM.txt");
            //gDPNamController.FillData(27, "Month", "YoY", "%", "Du lieu vi mo/Xuat nhap khau/Xuat khau theo quoc gia - mat hang/YoY.txt");





            //gDPNamController.FillData(28, "Month", "Value", "Triệu USD", "Du lieu vi mo/Xuat nhap khau/Nhap khau quoc gia - mat hang/Value.txt");
            //gDPNamController.FillData(28, "Month", "MoM", "", "Du lieu vi mo/Xuat nhap khau/Nhap khau quoc gia - mat hang/MoM.txt");
            //gDPNamController.FillData(28, "Month", "YoY", "", "Du lieu vi mo/Xuat nhap khau/Nhap khau quoc gia - mat hang/YoY.txt");

            //gDPNamController.FillData(29, "Month", "Value", "Nghìn người", "Du lieu vi mo/Van Tai/Khach quoc te theo loai hinh/Value.txt");
            //gDPNamController.FillData(29, "Month", "MoM", "", "Du lieu vi mo/Van Tai/Khach quoc te theo loai hinh/MoM.txt");
            //gDPNamController.FillData(29, "Month", "YoY", "", "Du lieu vi mo/Van Tai/Khach quoc te theo loai hinh/YoY.txt");

            //gDPNamController.FillData(30, "Month", "Value", "Nghìn người", "Du lieu vi mo/Van Tai/Khach quoc te theo quoc gia/Value.txt");
            //gDPNamController.FillData(30, "Month", "MoM", "", "Du lieu vi mo/Van Tai/Khach quoc te theo quoc gia/MoM.txt");
            //gDPNamController.FillData(30, "Month", "YoY", "", "Du lieu vi mo/Van Tai/Khach quoc te theo quoc gia/YoY.txt");

            //gDPNamController.FillData(31, "Month", "Value", "Nghìn người", "Du lieu vi mo/Van Tai/Van chuyen hanh khach/Value.txt");
            //gDPNamController.FillData(31, "Month", "MoM", "", "Du lieu vi mo/Van Tai/Van chuyen hanh khach/MoM.txt");
            //gDPNamController.FillData(31, "Month", "YoY", "", "Du lieu vi mo/Van Tai/Van chuyen hanh khach/YoY.txt");



            //gDPNamController.FillData(32, "Month", "Value", "Triệu HK/Km", "Du lieu vi mo/Van Tai/Luan chuyen hanh khach/Value.txt");
            //gDPNamController.FillData(32, "Month", "MoM", "%", "Du lieu vi mo/Van Tai/Luan chuyen hanh khach/MoM.txt");
            //gDPNamController.FillData(32, "Month", "YoY", "%", "Du lieu vi mo/Van Tai/Luan chuyen hanh khach/YoY.txt");


            //gDPNamController.FillData(33, "Month", "Value", "Nghìn tấn", "Du lieu vi mo/Van Tai/Van chuyen hang hoa/Value.txt");
            //gDPNamController.FillData(33, "Month", "MoM", "%", "Du lieu vi mo/Van Tai/Van chuyen hang hoa/MoM.txt");
            //gDPNamController.FillData(33, "Month", "YoY", "%", "Du lieu vi mo/Van Tai/Van chuyen hang hoa/YoY.txt");

            gDPNamController.FillData(34, "Month", "Value", "Triệu tấn/Km", "Du lieu vi mo/Van Tai/luan chuyen hang hoa/Value.txt");
            gDPNamController.FillData(34, "Month", "MoM", "%", "Du lieu vi mo/Van Tai/luan chuyen hang hoa/MoM.txt");
            gDPNamController.FillData(34, "Month", "YoY", "%", "Du lieu vi mo/Van Tai/luan chuyen hang hoa/YoY.txt");

            //gDPNamController.FillData(35, "Month", "Value", "USD mm", "Du lieu vi mo/Ty gia/du tru ngoai hoi/Value.txt");
            //gDPNamController.FillData(35, "Month", "MoM", "%", "Du lieu vi mo/Ty gia/du tru ngoai hoi/MoM.txt");

            //gDPNamController.FillData(36, "Quarter", "Q", "Triệu USD", "Du lieu vi mo/Ty gia/Can can thanh toan/Q.txt");
            //gDPNamController.FillData(36, "Quarter", "YTD", "Triệu USD", "Du lieu vi mo/Ty gia/Can can thanh toan/YTD.txt");
            //gDPNamController.FillData(36, "Quarter", "TTM", "Triệu USD", "Du lieu vi mo/Ty gia/Can can thanh toan/TTM.txt");

            //gDPNamController.FillData(37, "Quarter", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/Thu ngan sach/Value.txt");
            //gDPNamController.FillData(37, "Quarter", "YoY", "%", "Du lieu vi mo/Tai khoa/Thu ngan sach/YoY.txt");

            //gDPNamController.FillData(38, "Quarter", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/Chi ngan sach/Value.txt");
            //gDPNamController.FillData(38, "Quarter", "YoY", "%", "Du lieu vi mo/Tai khoa/Chi ngan sach/YoY.txt");

            //gDPNamController.FillData(39, "Year", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/vay no chinh phu/Value.txt");
            //gDPNamController.FillData(39, "Year", "YoY", "%", "Du lieu vi mo/Tai khoa/vay no chinh phu/YoY.txt");

            //gDPNamController.FillData(40, "Year", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/vay no chinh phu bao lanh/Value.txt");
            //gDPNamController.FillData(40, "Year", "YoY", "%", "Du lieu vi mo/Tai khoa/vay no chinh phu bao lanh/YoY.txt");

            //gDPNamController.FillData(41, "Year", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/vay no chinh quyen dia phuong/Value.txt");
            //gDPNamController.FillData(41, "Year", "YoY", "%", "Du lieu vi mo/Tai khoa/vay no chinh quyen dia phuong/YoY.txt");

            //gDPNamController.FillData(42, "Year", "Value", "Tỷ", "Du lieu vi mo/Tai khoa/Vay no ngoai quoc gia/Value.txt");
            //gDPNamController.FillData(42, "Year", "YoY", "%", "Du lieu vi mo/Tai khoa/Vay no ngoai quoc gia/YoY.txt");



            MessageBox.Show("Done");
        }

        private async void btnCheckBug_Click(object sender, EventArgs e)
        {
            //MacroService macroService = new MacroService();
            //macroService.GetAllMacro();

            // MacroTypeService macroTypeService = new MacroTypeService();
            //macroTypeService.Get_MacroType_By_KeyIDMacro("tieu-dung");
            //MacroDataService macroDataService = new MacroDataService();
            //await macroDataService.FillMacroData();
            TrashController trashController = new TrashController();
            //trashController.Update_KeyID_Row_Xuat_Khau_Quoc_Gia_Mat_Hang();
            XuatNhapKhauController xuatNhapKhauController = new XuatNhapKhauController();
            //xuatNhapKhauController.Load_MatHang_QuocGia();
            trashController.Update_KeyID_Row_Nhap_Khau_Quoc_Gia_Mat_Hang();
            MessageBox.Show("Done");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            MacroService macroService = new MacroService();
            var list = new BindingList<Macro>(macroService.GetAllMacro());
            dgvMacro.DataSource = list;
            ResizeView();

        }
        public void updateTxtBug(string message)
        {

            txtBug.Text = getTxtBug() + message + Environment.NewLine;
            txtBug.SelectionStart = txtBug.Text.Length;
            txtBug.ScrollToCaret();
        }
        public string getTxtBug()
        {
            return txtBug.Text;
        }
        private void ResizeView()
        {
            for (int i = 0; i < dgvMacro.Columns.Count - 1; i++)
            {
                dgvMacro.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dgvMacroType.Columns.Count - 1; i++)
            {
                dgvMacroType.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            for (int i = 0; i < dgvTable.Columns.Count - 1; i++)
            {
                dgvTable.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }


        private void dgvMacro_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgvMacro.Rows[e.RowIndex];
                string keyID = row.Cells["KeyID"].Value.ToString();
                MacroTypeService macroTypeService = new MacroTypeService();
                var list = new BindingList<MacroType>(macroTypeService.Get_MacroType_By_KeyIDMacro(keyID));
                dgvMacroType.DataSource = list;
                ResizeView();

                txt_Macro_ID.Text = row.Cells["Id"].Value.ToString();
                txt_Macro_KeyID.Text = row.Cells["KeyID"].Value.ToString();
                try
                {
                    if (row.Cells["Name"].Value!=null)
                    {
                        txt_Macro_Name_Vi.Text = row.Cells["Name"].Value.ToString();

                    }
                }
                catch { }
                //txt_MacroType_KeyIDMacro.Text = row.Cells["KeyIDMacro"].Value.ToString();

            }
        }

        private void dgvMacroType_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgvMacroType.Rows[e.RowIndex];
                int id = Int32.Parse(row.Cells["id"].Value.ToString());
                TableService tableService = new TableService();
                var list = new BindingList<Table>(tableService.Get_Table_By_IDMacroType(id));
                dgvTable.DataSource = list;
                ResizeView();

                txt_MacroType__ID.Text = row.Cells["id"].Value.ToString();
                txt_MacroType_ID_Detail.Text= row.Cells["IdDetail"].Value.ToString();
                txt_MacroType_Stt.Text = row.Cells["Stt"].Value.ToString();

                txt_Table_IDMacroType.Text = row.Cells["id"].Value.ToString();

                if (row.Cells["Name"].Value!=null)
                {
                    txt_MacroType_Name.Text = row.Cells["Name"].Value.ToString();
                }
                else
                {
                    txt_MacroType_Name.Text = "";
                }

                txt_MacroType_KeyID.Text= row.Cells["KeyID"].Value.ToString();
                txt_MacroType_KeyIDMacro.Text = row.Cells["KeyIDMacro"].Value.ToString();

            }
        }

        private void dgvRowDataLevel1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgvRowDataLevel1.Rows[e.RowIndex];
                int id = Int32.Parse(row.Cells["id"].Value.ToString());
                RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                var list = new BindingList<Row_Data_Level2>(rowDataLevel2Service.Get_RowDataLevel2_By_IdRowLevel1(id));
                dgvRowDataLevel2.DataSource = list;
                txt_Row_Data_Level2_IDRowDataLevel1.Text = id.ToString() ;
                ResizeView();
            }
        }

        private void dgvTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgvTable.Rows[e.RowIndex];
                int id = Int32.Parse(row.Cells["id"].Value.ToString());
                RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                var list = new BindingList<Row_Data_Level1>(rowDataLevel1Service.Get_RowDataLevel1_By_IdTable(id));
                dgvRowDataLevel1.DataSource = list;

                RowService rowService = new RowService();
                var listRow = new BindingList<Row>(rowService.Get_Rows_By_IdTable(id));
                dgv_Row.DataSource = listRow;

                ResizeView();

                txt_Table_ID.Text = row.Cells["id"].Value.ToString();
                txt_Table_KeyID.Text = row.Cells["KeyID"].Value.ToString();
                txt_Table_ValueType.Text = row.Cells["ValueType"].Value.ToString();
                txt_Table_DateType.Text = row.Cells["DateType"].Value.ToString();
                txt_Table_Stt.Text = row.Cells["Stt"].Value.ToString();
                txt_Table_IDMacroType.Text = row.Cells["IDMacroType"].Value.ToString();
                txt_Table_Name.Text = row.Cells["Name"].Value.ToString();
                txt_Table_Unit.Text = row.Cells["Unit"].Value.ToString();

                try
                {
                    var tableType = row.Cells["TableType"].Value;
                    if (tableType != null)
                    {
                        txt_Table_TableType.Text = tableType.ToString();
                    }
                }
                catch
                {
                    txt_Table_TableType.Text = "";
                }


            }
        }

        private void dgvRowDataLevel3_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvRowDataLevel2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgvRowDataLevel2.Rows[e.RowIndex];
                int id = Int32.Parse(row.Cells["id"].Value.ToString());
                RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();
                var list = new BindingList<Row_Data_Level3>(rowDataLevel3Service.Get_RowDataLevel3_By_IdRowLevel2(id));
                dgvRowDataLevel3.DataSource = list;
                ResizeView();
            }
        }

        private void btn_Edit_MacroType_Click(object sender, EventArgs e)
        {
            AllKeyService allKeyService = new AllKeyService();
            allKeyService.Update(txt_MacroType_KeyID.Text, txt_MacroType_Name.Text);
            MacroTypeService macroTypeService = new MacroTypeService();
            MacroType macroType = new MacroType();

            macroType.Id = Int32.Parse(txt_MacroType__ID.Text);
            macroType.IdDetail = Int32.Parse(txt_MacroType_ID_Detail.Text);
            macroType.KeyID = txt_MacroType_KeyID.Text;
            macroType.KeyIDMacro = txt_MacroType_KeyIDMacro.Text;
            macroType.Name = txt_MacroType_Name.Text;
            macroType.Stt = Int32.Parse(txt_MacroType_Stt.Text);

            macroTypeService.Update(macroType);
            var list = new BindingList<MacroType>(macroTypeService.Get_MacroType_By_KeyIDMacro(txt_MacroType_KeyIDMacro.Text));
            dgvMacroType.DataSource = list;
            ResizeView();
        }

        private void btn_Add_MacroType_Click(object sender, EventArgs e)
        {
            MacroTypeService macroTypeService = new MacroTypeService();
            MacroType macroType = new MacroType();
            try
            {
                macroType.IdDetail = Int32.Parse(txt_MacroType_ID_Detail.Text);
            }
            catch { }
            macroType.KeyID = txt_MacroType_KeyID.Text;
            macroType.KeyIDMacro = txt_MacroType_KeyIDMacro.Text;
            macroType.Name = txt_MacroType_Name.Text;
            macroType.Stt = Int32.Parse(txt_MacroType_Stt.Text);

            AllKeyService allKeyService = new AllKeyService();
            allKeyService.Update(macroType.KeyID, macroType.Name);

            string mes = macroTypeService.Insert(macroType);
            var list = new BindingList<MacroType>(macroTypeService.Get_MacroType_By_KeyIDMacro(txt_MacroType_KeyIDMacro.Text));
            dgvMacroType.DataSource = list;
            ResizeView();
            MessageBox.Show(mes);

        }

        private void btn_Table_Add_Click(object sender, EventArgs e)
        {
            AllKeyService allKeyService = new AllKeyService();
            allKeyService.Update(txt_Table_KeyID.Text, txt_Table_Name.Text);
            TableService tableService = new TableService();
            Table table = new Table();

            table.DateType = txt_Table_DateType.Text;
            table.IdMacroType = Int32.Parse(txt_Table_IDMacroType.Text);
            table.KeyID = txt_Table_KeyID.Text;
            table.Name = txt_Table_Name.Text;
            table.ValueType = txt_Table_ValueType.Text;
            table.TableType = txt_Table_TableType.Text;
            table.Unit = txt_Table_Unit.Text;
            table.Stt = Int32.Parse(txt_Table_Stt.Text);

            tableService.InsertPG(table);
            var list = new BindingList<Table>(tableService.Get_Table_By_IDMacroType(table.IdMacroType));
            dgvTable.DataSource = list;
            ResizeView();
        }

        private void btn_Table_Edit_Click(object sender, EventArgs e)
        {
            AllKeyService allKeyService = new AllKeyService();
            allKeyService.Update(txt_Table_KeyID.Text, txt_Table_Name.Text);
            TableService tableService = new TableService();
            Table table = new Table();

            table.Id = Int32.Parse(txt_Table_ID.Text);
            table.DateType = txt_Table_DateType.Text;
            table.IdMacroType = Int32.Parse(txt_Table_IDMacroType.Text);
            table.KeyID = txt_Table_KeyID.Text;
            table.Name = txt_Table_Name.Text;
            table.ValueType = txt_Table_ValueType.Text;
            table.Stt = Int32.Parse(txt_Table_Stt.Text);
            table.TableType = txt_Table_TableType.Text;
            table.Unit = txt_Table_Unit.Text;
            tableService.Update(table);
            var list = new BindingList<Table>(tableService.Get_Table_By_IDMacroType(table.IdMacroType));
            dgvTable.DataSource = list;
            ResizeView();
        }

        private void btnAddMacro_Click(object sender, EventArgs e)
        {

        }

        private void btnLoadDataFromExcel_Click(object sender, EventArgs e)
        {
            //LoadDataExcelController loadDataExcelController = new LoadDataExcelController();
            //loadDataExcelController.Load_TieuDung_BanLeHangHoaVaDichVu();

            //loadDataExcelController.Load_DauTu_VonDauTuTuNSNN();
            //loadDataExcelController.Load_DauTu_DangKyKinhDoanh();

            //loadDataExcelController.Load_SanXuat_SanPhamCongNghiep();
            //loadDataExcelController.Load_SanXuat_IIP();

            FDIController fDIController = new FDIController();
            //fDIController.Load_FDI_DauTuTrucTiepTuNuocNgoai();
            //fDIController.Load_FDI_Tinh_Thanh();

            //XuatNhapKhauController xuatNhapKhauController = new XuatNhapKhauController();
            //xuatNhapKhauController.Load_XuatNhapKhau_QuocGia_MatHang();
            //xuatNhapKhauController.Load_NhapKhau_QuocGia_MatHang();
            //xuatNhapKhauController.Load_NhapKhau_MatHang_QuocGia();

            //xuatNhapKhauController.Load_MatHang_QuocGia();

            VanTaiController vanTaiController = new VanTaiController();
            //vanTaiController.Load_VanTai_KhachQuocTeTheoLoaiHinh();
            //vanTaiController.Load_VanTai_KhachQuocTeTheoQuocGia();
            //vanTaiController.Load_VanTai_VanTaiHanhKhach();
            //vanTaiController.Load_VanTai_VanTaiHanhKhach();
            //vanTaiController.Load_VanTai_LuanChuyenHanhKhach();
            //vanTaiController.Load_VanTai_VanChuyenHangHoa();
            //vanTaiController.Load_VanTai_LuanChuyenHangHoa();

            TieuDungController tieuDungController = new TieuDungController();
            //tieuDungController.Load_TieuDung_CPI();
            //tieuDungController.Load_TieuDung_BanLeHangHoaVaDichVu();

            DauTuController dauTuController = new DauTuController();
            //dauTuController.Load_DauTu_DangKyKinhDoanh();
            //dauTuController.Load_DauTu_VonDauTuTuNSNN();

            SanXuatController sanXuatController = new SanXuatController();
            sanXuatController.Load_SanXuat_IIP();
            MessageBox.Show("Done");
        }

        private void btnEditMacro_Click(object sender, EventArgs e)
        {
            Macro macro = new Macro();
            macro.Id = int.Parse(txt_Macro_ID.Text);
            macro.KeyID = txt_Macro_KeyID.Text;
            macro.Name = txt_Macro_Name_Vi.Text;
            MacroService macroService = new MacroService();
            macroService.Update(macro);
            var list = new BindingList<Macro>(macroService.GetAllMacro());
            dgvRowDataLevel2.DataSource = list;
        }

        private void dgvTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_Row_Data_Level2_Them_Click(object sender, EventArgs e)
        {
            Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();

            row_Data_Level2.IdRowDataLevel1 = int.Parse(txt_Row_Data_Level2_IDRowDataLevel1.Text);
            row_Data_Level2.KeyID = txt_Row_Data_Level2_KeyID.Text;
            row_Data_Level2.Name = txt_Row_Data_Level2_Name.Text;
            row_Data_Level2.Stt = int.Parse(txt_Row_Data_Level2_Stt.Text);
            row_Data_Level2.Unit = txt_Row_Data_Level2_Unit.Text;
            RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
            rowDataLevel2Service.InsertPG(row_Data_Level2);
            var list = new BindingList<Row_Data_Level2>(rowDataLevel2Service.Get_RowDataLevel2_By_IdRowLevel1(row_Data_Level2.IdRowDataLevel1));
            dgvRowDataLevel2.DataSource = list;
        }

        private void dgv_Row_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = this.dgv_Row.Rows[e.RowIndex];

            txt_Row_ID.Text = row.Cells["id"].Value.ToString();
            txt_Row_ID_Table.Text = row.Cells["id_table"].Value.ToString();
            txt_Row_KeyID.Text = row.Cells["key_id"].Value.ToString();
            txt_Row_Name.Text = row.Cells["name"].Value.ToString();
            txt_Row_Stt.Text = row.Cells["stt"].Value.ToString();
            txt_Row_Unit.Text = row.Cells["unit"].Value.ToString();
            txt_Row_Level.Text = row.Cells["level"].Value.ToString();
            
        }

        private void btn_Row_Edit_Click(object sender, EventArgs e)
        {
            Row row = new Row();
            row.ID = int.Parse(txt_Row_ID.Text);
            row.Key_ID = txt_Row_KeyID.Text;
            row.ID_Table = int.Parse(txt_Row_ID_Table.Text);
            row.Level = int.Parse(txt_Row_Level.Text);
            row.Name = txt_Row_Name.Text;
            row.Stt = int.Parse(txt_Row_Stt.Text);
            row.Unit = txt_Row_Unit.Text;
            RowService rowService = new RowService();
            rowService.Update(row);
            var list = new BindingList<Row>(rowService.Get_Rows_By_IdTable(row.ID_Table));
            dgv_Row.DataSource = list;
        }

        private void btn_Row_Add_Click(object sender, EventArgs e)
        {
            Row row = new Row();
            YAxisService yAxisService = new YAxisService();
            row.Key_ID = txt_Row_KeyID.Text;
            row.Level = int.Parse(txt_Row_Level.Text);
            row.Name = txt_Row_Name.Text;
            row.Stt = int.Parse(txt_Row_Stt.Text);
            row.Unit = txt_Row_Unit.Text;
            row.YAxis = yAxisService.GetYAxis(row.Unit);
            row.ID_Table = int.Parse(txt_Table_ID.Text);
            RowService rowService = new RowService();
            rowService.Insert_And_Update_STT(row);
            var list = new BindingList<Row>(rowService.Get_Rows_By_IdTable(row.ID_Table));
            dgv_Row.DataSource = list;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Row row = new Row();
            row.ID = int.Parse(txt_Row_ID.Text);
            row.Key_ID = txt_Row_KeyID.Text;
            row.ID_Table = int.Parse(txt_Row_ID_Table.Text);
            row.Level = int.Parse(txt_Row_Level.Text);
            row.Name = txt_Row_Name.Text;
            row.Stt = int.Parse(txt_Row_Stt.Text);
            row.Unit = txt_Row_Unit.Text;
            RowService rowService = new RowService();
            rowService.Delete(row);
            var list = new BindingList<Row>(rowService.Get_Rows_By_IdTable(row.ID_Table));
            dgv_Row.DataSource = list;
        }

        private void btn_Table_ClearAllTable_Click(object sender, EventArgs e)
        {
            MacroTypeService macroTypeService = new MacroTypeService();
            MacroType macroType = new MacroType();
            TableService tableService = new TableService();
            macroType.Id = Int32.Parse(txt_MacroType__ID.Text);
            macroType.IdDetail = Int32.Parse(txt_MacroType_ID_Detail.Text);
            macroType.KeyID = txt_MacroType_KeyID.Text;
            macroType.KeyIDMacro = txt_MacroType_KeyIDMacro.Text;
            macroType.Name = txt_MacroType_Name.Text;
            macroType.Stt = Int32.Parse(txt_MacroType_Stt.Text);

            macroTypeService.ClearAllTable(macroType);
            var list = new BindingList<Table>(tableService.Get_Table_By_IDMacroType(macroType.Id));
            dgvTable.DataSource = list;
        }

        private void btn_Table_Delete_Click(object sender, EventArgs e)
        {
            TableService tableService = new TableService();
            Table table = new Table();
            table.Id = int.Parse(txt_Table_ID.Text);
            tableService.Delete(table);
            var list = new BindingList<Table>(tableService.Get_Table_By_IDMacroType(int.Parse(txt_MacroType__ID.Text)));
            dgvTable.DataSource = list;
        }
    }
}
