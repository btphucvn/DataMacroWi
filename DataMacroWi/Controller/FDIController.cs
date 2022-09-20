
using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DataMacroWi.Controller
{
    class FDIController
    {
        public void Load_FDI_DauTuTrucTiepTuNuocNgoai()
        {


            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--window-size=1300,1000");


            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            TableService tableService = new TableService();
            Table table = new Table();
            table = tableService.Get_Table_By_KeyID_TableType_ValueType("fdi", "", "Value");

            //Ẩn command dos
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            ChromeDriver driver = new ChromeDriver(driverService, options);
            DateTime now = DateTime.Now;
            driver.Navigate().GoToUrl("https://www.mpi.gov.vn/congkhaithongtin/Pages/solieudautunuocngoai.aspx");

            //string json = "";
            try
            {


                bool flag = true;

                while (flag)
                {
                    try
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        string checkLoad = js.ExecuteScript("return document.readyState").ToString();
                        if (checkLoad == "complete")
                        {
                            flag = false;
                        }
                        while (checkLoad != "complete")
                        {
                            checkLoad = js.ExecuteScript("return document.readyState").ToString();
                            Thread.Sleep(1000);
                        }

                    }
                    catch
                    {
                        driver.Quit();
                        Thread.Sleep(5000);
                        driver = new ChromeDriver(driverService, options);
                    }
                }
                Thread.Sleep(5000);
                //*[contains(@class, 'content_2col_block_content')] / ul / li//cufon[1]/cufontext
                for (int tr = 4; tr >= 1; tr--)
                {
                    for (int td = 3; td >= 1; td--)
                    {
                        string vonThucHien = "";
                        try
                        {
                            vonThucHien = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[1]/span[2]")).Text;
                        }
                        catch
                        {
                            continue;
                        }
                        string thang = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//h3//a")).Text;
                        string dangKyCapMoi = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[2]/p[1]/span[2]")).Text;
                        string dangKyThem = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[2]/p[2]/span[2]")).Text;
                        string gopMuaCoPhan = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[2]/p[3]/span[2]")).Text;
                        string capMoi = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[3]/p[1]/span[2]")).Text;
                        string tangVon = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[3]/p[2]/span[2]")).Text;
                        string gopVonMuaCoPhan = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[3]/p[3]/span[2]")).Text;
                        //string xuatKhauKeCaDauTho = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[4]/p[1]/span[2]")).Text;
                        //string xuatKhauKhongKeCaDauTho = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[4]/p[2]/span[2]")).Text;
                        //string nhapKhau = driver.FindElement(By.XPath("//*[@id='ctl00_m_g_37199170_22bb_4ce6_a280_94eaf23b03ad']//table[2]//tbody/tr[" + tr + "]//td[" + td + "]//ul/li[5]/span[2]")).Text;

                        thang = CleanNumber(thang);
                        double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp("01" + "-" + thang + "-" + now.Year);
                        double num_VonThucHien = double.Parse(CleanNumber(vonThucHien.Replace(".", "")));
                        double num_DangKyCapMoi = double.Parse(CleanNumber(dangKyCapMoi.Replace(",", ".")));
                        double num_DangKyThem = double.Parse(CleanNumber(dangKyThem.Replace(",", ".")));
                        double num_GopMuaCoPhan = double.Parse(CleanNumber(gopMuaCoPhan.Replace(",", ".")));
                        double num_CapMoi = double.Parse(CleanNumber(capMoi.Replace(",", ".")));
                        double num_TangVon = double.Parse(CleanNumber(tangVon.Replace(",", ".")));
                        double num_SoDuAn_GopVonMuaCoPhan = double.Parse(CleanNumber(gopVonMuaCoPhan.Replace(",", ".")));
                        double vonDangKy = num_DangKyCapMoi + num_DangKyThem + num_GopMuaCoPhan;
                        Row row_VonThucHien = new Row();
                        row_VonThucHien = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-thuc-hien", "Triệu USD", table.Id);
                        Row_Value row_Value = new Row_Value();
                        row_Value.ID_Row = row_VonThucHien.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_VonThucHien;
                        rowValueService.Insert_Update(row_Value);

                        Row row_VonDangKy = new Row();
                        row_VonDangKy = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky", "Triệu USD", table.Id);
                        row_Value.ID_Row = row_VonDangKy.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = vonDangKy;
                        rowValueService.Insert_Update(row_Value);

                        Row row_DangKyCapMoi = new Row();
                        row_DangKyCapMoi = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_dang-ky-cap-moi", "Triệu USD", table.Id);
                        row_Value.ID_Row = row_DangKyCapMoi.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_DangKyCapMoi;
                        rowValueService.Insert_Update(row_Value);


                        Row row_DangKyTangThem = new Row();
                        row_DangKyTangThem = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_dang-ky-tang-them", "Triệu USD", table.Id);
                        row_Value.ID_Row = row_DangKyTangThem.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_DangKyThem;
                        rowValueService.Insert_Update(row_Value);

                        Row row_GopVonMuaCoPhan = new Row();
                        row_GopVonMuaCoPhan = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_gop-von-mua-co-phan", "Triệu USD", table.Id);
                        row_Value.ID_Row = row_GopVonMuaCoPhan.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_GopMuaCoPhan;
                        rowValueService.Insert_Update(row_Value);



                        Row row_SoDuAnCapMoi = new Row();
                        row_SoDuAnCapMoi = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_so-du-an-cap-moi", "Dự án", table.Id);
                        row_Value.ID_Row = row_SoDuAnCapMoi.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_SoDuAn_GopVonMuaCoPhan;
                        rowValueService.Insert_Update(row_Value);


                        Row row_SoDuAnTangVon = new Row();
                        row_SoDuAnTangVon = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_so-du-an-tang-von", "Dự án", table.Id);
                        row_Value.ID_Row = row_SoDuAnTangVon.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_TangVon;
                        rowValueService.Insert_Update(row_Value);

                        Row row_SoDuAn_GonVonMuaCoPhan = new Row();
                        row_SoDuAn_GonVonMuaCoPhan = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi_Value__von-dang-ky_gop-von-mua-co-phan", "Dự án", table.Id);
                        row_Value.ID_Row = row_SoDuAn_GonVonMuaCoPhan.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = num_SoDuAn_GopVonMuaCoPhan;
                        rowValueService.Insert_Update(row_Value);

                    }

                }





            }
            catch (Exception e)
            {
            }

            driver.Close();
            driver.Quit();


        }

        public void Load_FDI_TheoLinhVuc()
        {
            string linkFolder = "DataMacro\\FDI\\FDI Theo Linh Vuc";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            Table table;
            string keyIDMacroType = "fdi-dang-ky-theo-linh-vuc";

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 4; col <= 4; col++)
                {


                    string keyIDTable = (excelSheet.Cells[4, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[5, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[6, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }


                    int level = -1;
                    int stt = 0;
                    string unit = table.Unit;
                    for (int rowIndex = 8; rowIndex <= 150; rowIndex++)
                    {
                        try
                        {
                            level = int.Parse((excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                        }
                        catch
                        {
                            break;
                        }
                        string keyID = (excelSheet.Cells[rowIndex, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                        keyID = keyID.Replace("ValueType", valueType);
                        keyID = keyID.Replace("TableType", tableType);
                        try
                        {
                            double value = double.NaN;
                            try
                            {
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.Stt = stt;
                            rowService.Update(row);
                            stt = stt + 1;
                            if (row.Key_ID == null)
                            {
                                unit = ToolData.getUnitFromTitle(rowName);
                                rowName = ToolData.removeUnitFromTitle(rowName);
                                row.ID_String = keyID;
                                row.ID_Table = table.Id;
                                row.Key_ID = keyID;
                                row.Level = level;
                                row.Name = rowName;
                                row.Stt = rowIndex;
                                if (unit == "")
                                {
                                    unit = table.Unit;
                                }
                                row.Unit = unit;
                                row.YAxis = yAxisService.GetYAxis(unit);
                                row.ID = rowService.Insert(row);
                            }
                            Row_Value row_Value = new Row_Value();
                            row_Value.ID_Row = row.ID;
                            row_Value.TimeStamp = timeStamp;



                            row_Value.Value = value;

                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "Value", keyIDMacroType);
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "MoM", keyIDMacroType);
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "YoY", keyIDMacroType);
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "YoY", keyIDMacroType);
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "Value", keyIDMacroType);
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");
        }


        public void Load_FDI_Quoc_Gia()
        {
            string linkFolder = "DataMacro\\FDI\\FDI Theo Quoc Gia";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            CountryService countryService = new CountryService();
            Table table;
            string keyIDMacroType = "fdi-dang-ky-theo-quoc-gia";
            double lastestTimeStamp = 0;
            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                lastestTimeStamp = timeStamp;
                string keyIDTable = "fdi-dang-ky";
                string valueType = "Value";
                string tableType = "YTD";
                table = new Table();
                table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
                if (table.KeyID == null)
                {
                    continue;
                }


                int level = 3;
                int stt = 0;
                string unit = table.Unit;
                for (int rowIndex = 1; rowIndex <= 150; rowIndex++)
                {
                    double value = double.NaN;
                    try
                    {
                        value = double.Parse((excelSheet.Cells[rowIndex, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                    }
                    catch { break; }
                    string nuoc = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    string chau = countryService.GetContinent(nuoc.Trim());
                    string keyID = keyIDTable+"_"+valueType+"_"+tableType+"_"+Tool.titleToKeyID(chau)+"_"+ Tool.titleToKeyID(nuoc);
                    if (chau == "")
                    {
                        Form1._Form1.updateTxtBug("Không tìm thấy nước: " + nuoc);
                        continue;
                    }
                    try
                    {

                        string rowName = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                        Row row = rowService.Get_Row_By_KeyID(keyID);
                        if(row==null)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyID);
                            continue;
                        }
                        Row_Value row_Value = new Row_Value();

                        row_Value.ID_Row = row.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;

                        rowValueService.Insert_Update(row_Value);
                    }
                    catch { }


                }
                wb.Close();
            }



            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "Value", keyIDMacroType);

            Table tableYTDValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "Value", keyIDMacroType);

            Row row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-a", "Triệu USD", tableYTDValue.Id);
            Row_Value row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-a", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-au", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-au", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-my", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-my", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-phi", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-phi", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-uc", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-uc", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_chau-dai-duong", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("chau-dai-duong", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);


            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableYTDValue.Id);
            rowService.Clear(tableValue);
            toolController.Insert_Row_Value_Calculate(listRow, tableValue, "", "Value");

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "MoM", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            Table tableMoMYTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "MoM", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoMYTD);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoMYTD, "YTD", "MoM");

            Table tableYoYYTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "YoY", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableYoYYTD);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoYYTD, "YTD", "YoY");

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "YoY", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");

            Form1._Form1.updateTxtBug("Hoàn tất FDI theo quốc gia");

        }

        public void Load_FDI_Tinh_Thanh()
        {
            string linkFolder = "DataMacro\\FDI\\FDI Dang ky theo tinh thanh";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            ProvinceService provinceService = new ProvinceService();
            Table table;
            string keyIDMacroType = "fdi-dang-ky-theo-tinh-thanh";
            double lastestTimeStamp = 0;
            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                lastestTimeStamp = timeStamp;
                string keyIDTable = "fdi-dang-ky";
                string valueType = "Value";
                string tableType = "YTD";
                table = new Table();
                table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
                if (table.KeyID == null)
                {
                    continue;
                }


                int level = 3;
                int stt = 0;
                string unit = table.Unit;
                for (int rowIndex = 1; rowIndex <= 150; rowIndex++)
                {
                    double value = double.NaN;
                    try
                    {
                        value = double.Parse((excelSheet.Cells[rowIndex, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                    }
                    catch { break; }
                    string nuoc = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    string chau = provinceService.GetRegion(nuoc.Trim());
                    string keyID = keyIDTable + "_" + valueType + "_" + tableType + "_" + Tool.titleToKeyID(chau) + "_" + Tool.titleToKeyID(nuoc);
                    if (chau == "")
                    {
                        Form1._Form1.updateTxtBug("Không tìm thấy nước: " + nuoc);
                        continue;
                    }
                    try
                    {

                        string rowName = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                        Row row = rowService.Get_Row_By_KeyID(keyID);
                        if (row == null)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyID);
                            continue;
                        }
                        Row_Value row_Value = new Row_Value();

                        row_Value.ID_Row = row.ID;
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;

                        rowValueService.Insert_Update(row_Value);
                    }
                    catch { }


                }
                wb.Close();
            }



            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "Value", keyIDMacroType);

            Table tableYTDValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "Value", keyIDMacroType);

            Row row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_mien-bac", "Triệu USD", tableYTDValue.Id);
            Row_Value row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("mien-bac", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_mien-trung", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("mien-trung", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);

            row_ChauLuc = rowService.Get_Row_By_KeyID_Unit_IDTable("fdi-dang-ky_Value_YTD_mien-nam", "Triệu USD", tableYTDValue.Id);
            row_ChauLuc_Value = rowValueService.Get_Row_Value_By_IDRow_TimeStamp(row_ChauLuc.ID, lastestTimeStamp);
            row_ChauLuc_Value.Value = rowService.Sum_Contain_KeyID("mien-nam", tableYTDValue.Id, lastestTimeStamp);
            row_ChauLuc_Value.TimeStamp = lastestTimeStamp;
            row_ChauLuc_Value.ID_Row = row_ChauLuc.ID;
            rowValueService.Insert_Update(row_ChauLuc_Value);




            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableYTDValue.Id);
            rowService.Clear(tableValue);
            toolController.Insert_Row_Value_Calculate(listRow, tableValue, "", "Value");

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "MoM", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            Table tableMoMYTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "MoM", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoMYTD);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoMYTD, "YTD", "MoM");

            Table tableYoYYTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "YTD", "YoY", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableYoYYTD);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoYYTD, "YTD", "YoY");

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("fdi-dang-ky", "", "YoY", keyIDMacroType);
            listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");

            Form1._Form1.updateTxtBug("Hoàn tất FDI theo Tỉnh thành");

        }

        string CleanNumber(string input)
        {
            string result = string.Empty;
            foreach (var c in input)
            {
                int ascii = (int)c;
                if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
                    result += c;
            }
            return result;
        }
    }
}
