using DataMacroWi.DB;
using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using Microsoft.Office.Interop.Excel;
using Npgsql;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Controller
{
    class XuatNhapKhauController
    {
        public void Load_MatHang_QuocGia()
        {
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            TableService tableService = new TableService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "Value", "xuat-khau-quoc-gia-mat-hang");
            Table tableMatHangQuocGia = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "Value", "xuat-khau-mat-hang-theo-quoc-gia");

            rowService.Clear(tableMatHangQuocGia);

            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            List<Row> listRowCountry = new List<Row>();
            List<Row> listMatHang = new List<Row>();
            //foreach (Row row in listRow)
            //{
            //    if (row.Level == 2)
            //    {
            //        row.Rows = new List<Row>();
            //        listRowCountry.Add(row);
            //    }
            //}
            //foreach (Row row in listRow)
            //{
            //    if (row.Level == 3)
            //    {
            //        string[] tmpKeyID = row.Key_ID.Split('_');
            //        for (int i = 0; i < listRowCountry.Count; i++)
            //        {
            //            if (listRowCountry[i].Key_ID.Contains("_" + tmpKeyID[4]))
            //            {
            //                listRowCountry[i].Rows.Add(row);
            //            }
            //        }
            //    }
            //}
            foreach (Row row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                row.Row_Values = listRowValue;
            }
            //lấy danh sách mặt hàng
            foreach (Row row in listRow)
            {
                string[] tmpKeyID = row.Key_ID.Split('_');
                if (tmpKeyID.Length >= 6)
                {
                    Row rowMatHang = listMatHang.Find(x => x.Key_ID == tmpKeyID[5]);
                    if (rowMatHang == null)
                    {
                        rowMatHang = new Row();
                        rowMatHang.Key_ID = tmpKeyID[5];
                        rowMatHang.Name = row.Name;
                        rowMatHang.Rows = new List<Row>();
                        listMatHang.Add(rowMatHang);
                    }

                }
            }
            //lấy danh sách mặt hàng - quốc gia
            foreach (Row row in listRow)
            {

                for (int i = 0; i < listMatHang.Count; i++)
                {

                    if (row.Key_ID.Contains(listMatHang[i].Key_ID))
                    {
                        listMatHang[i].Rows.Add(row);
                    }
                }
            }
            //Cập nhật lại tên quốc gia trong childRow
            foreach (Row row in listMatHang)
            {

                foreach (Row row_child in row.Rows)
                {
                    string[] tmp = row_child.Key_ID.Split('_');
                    string keySearch = tmp[0] + "_" + tmp[1] + "_" + tmp[2] + "_" + tmp[3] + "_" + tmp[4];
                    Row rowSearch = listRow.Find(x => x.Key_ID == keySearch);
                    row_child.Name = rowSearch.Name;
                }
            }
            int stt = 0;
            foreach (Row row in listMatHang)
            {
                row.Row_Values = rowService.Sum_Contain_KeyID_By_TimeStamp(table.Id, row.Key_ID);
                row.ID_Table = tableMatHangQuocGia.Id;
                row.Level = 2;
                row.Unit = tableMatHangQuocGia.Unit;
                row.Stt = stt;
                stt = stt + 1;
                //row.Key_ID = tableMatHangQuocGia.KeyID+"_"+
                row.ID = rowService.Insert(row);
                foreach (Row_Value row_Value in row.Row_Values)
                {
                    row_Value.ID_Row = row.ID;
                    rowValueService.Insert_Update(row_Value);
                }
                foreach (Row rowChild in row.Rows)
                {
                    rowChild.Row_Values = rowService.Sum_Contain_KeyID_By_TimeStamp(table.Id, rowChild.Key_ID);
                    rowChild.ID_Table = tableMatHangQuocGia.Id;
                    rowChild.Level = 3;
                    rowChild.Unit = tableMatHangQuocGia.Unit;
                    rowChild.Stt = stt;
                    stt = stt + 1;
                    rowChild.ID = rowService.Insert(rowChild);

                    foreach (Row_Value row_Value in rowChild.Row_Values)
                    {
                        row_Value.ID_Row = rowChild.ID;
                        rowValueService.Insert_Update(row_Value);
                    }
                }
            }
            //chỉnh sửa lại keyid
            foreach (Row row in listMatHang)
            {
                foreach (Row rowChild in row.Rows)
                {

                    string[] tmp = rowChild.Key_ID.Split('_');
                    if (tmp.Length > 2)
                    {
                        rowChild.Key_ID = tmp[5] + "_" + tmp[4];
                    }

                }
            }
            //chỉnh sửa lại keyid theo table
            for (int i = 0; i < listMatHang.Count; i++)
            {
                listMatHang[i].Key_ID = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + listMatHang[i].Key_ID;
                try
                {
                    listMatHang[i].Rows = listMatHang[i].Rows.OrderByDescending(item => item.Row_Values[0].Value).ToList();
                }
                catch { }
                foreach (Row rowChild in listMatHang[i].Rows)
                {
                    string[] tmp = rowChild.Key_ID.Split('_');
                    rowChild.Key_ID = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + tmp[1] + "_" + tmp[0];
                    rowService.Update(rowChild);
                }
            }
            //Lấy tổng bên table xuất khẩu mặt hàng lắp sang
            Table tableXuatKhauMatHang = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "Value", "xuat-khau-theo-mat-hang");
            List<Row> listRowXuatKhauMatHang = rowService.Get_Rows_By_IdTable(tableXuatKhauMatHang.Id);
            foreach (Row row in listMatHang)
            {
                string[] tmpRow = row.Key_ID.Split('_');
                foreach (Row rowXuatKhauMatHang in listRowXuatKhauMatHang)
                {
                    string[] tmpRowXuatKhauMatHang = rowXuatKhauMatHang.Key_ID.Split('_');
                    if (tmpRow[tmpRow.Length - 1] == tmpRowXuatKhauMatHang[tmpRowXuatKhauMatHang.Length - 1])
                    {
                        List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(rowXuatKhauMatHang.ID);
                        row.Row_Values = listRowValue.OrderByDescending(x => x.TimeStamp).ToList();
                        foreach (Row_Value row_Value in row.Row_Values)
                        {
                            row_Value.ID_Row = row.ID;
                        }
                    }
                }
            }
            //update lại stt
            stt = 0;
            foreach (Row row in listMatHang)
            {

                row.Stt = stt;
                stt = stt + 1;
                rowService.Update(row);
                foreach (Row_Value row_value in row.Row_Values)
                {
                    rowValueService.Insert_Update(row_value);
                }
                foreach (Row rowChild in row.Rows)
                {
                    rowChild.Stt = stt;
                    stt = stt + 1;
                    rowService.Update(rowChild);

                }
            }



        }
        public void Load_XuatNhapKhau_QuocGia_MatHang()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Xuat Khau Quoc Gia - Mat hang";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            CountryService countryService = new CountryService();
            Table table;
            string keyIDMacroType = "xuat-khau-quoc-gia-mat-hang";
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
                string keyIDTable = "xuat-khau";
                string valueType = "Value";
                string tableType = "";
                table = new Table();
                table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
                if (table.KeyID == null)
                {
                    continue;
                }



                string unit = table.Unit;
                string keyIDRow = keyIDTable + "_" + valueType + "_" + tableType;

                for (int rowIndex = 1; rowIndex <= 1500; rowIndex++)
                {
                    string rowName = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    if (rowName == null || rowName == "")
                    {
                        break;
                    }
                    double value = 0;
                    try
                    {
                        string stringValue = (excelSheet.Cells[rowIndex, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                        value = Math.Round(double.Parse(stringValue.Replace(".", "")) / 1000000, 2);
                    }
                    catch
                    {

                    }
                    if (!rowName.Any(char.IsLower))
                    {

                        CountryModel countryModel = new CountryModel();
                        countryModel = countryService.Get_By_Country_Name_Vi(rowName);
                        if (countryModel.Continent == null)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + countryModel.Continent + " " + countryModel.Country + " " + rowName);
                            break;
                        }
                        try
                        {
                            keyIDRow = keyIDTable + "_" + valueType + "_" + tableType + "_" + Tool.titleToKeyID(countryModel.Continent) + "_" + countryModel.Key_ID;
                        }
                        catch
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + countryModel.Continent + " " + countryModel.Country);
                            break;
                        }
                        Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyIDRow, unit, table.Id);
                        if (row == null || row.ID == 0)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRow);
                            break;
                        }
                        Row_Value row_Value = new Row_Value();
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;
                        row_Value.ID_Row = row.ID;
                        rowValueService.Insert_Update(row_Value);
                    }
                    else
                    {
                        string keyIDRowLevel3 = "";
                        try
                        {
                            keyIDRowLevel3 = keyIDRow + "_" + Tool.titleToKeyID(rowName);
                        }
                        catch
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRow + " " + rowName);
                            break;
                        }
                        Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyIDRowLevel3, unit, table.Id);
                        if (row == null || row.ID == 0)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRowLevel3);
                            break;
                        }
                        Row_Value row_Value = new Row_Value();
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;
                        row_Value.ID_Row = row.ID;
                        rowValueService.Insert_Update(row_Value);

                    }


                }
                wb.Close();
            }



            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "Value", keyIDMacroType);
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-a", "xuat-khau_Value__chau-a");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-au", "xuat-khau_Value__chau-au");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-my", "xuat-khau_Value__chau-my");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-uc", "xuat-khau_Value__chau-uc");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-dai-duong", "xuat-khau_Value__chau-dai-duong");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-phi", "xuat-khau_Value__chau-phi");










            Form1._Form1.updateTxtBug("Hoàn tất Xuat Nhap Khau");

        }
        public void Insert_Sum_Value_By_Continient(int idTable, string keyIDContinient, string keyIDRow)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();

            Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyIDRow, "Triệu USD", idTable);
            string query = "WITH rows_table AS " +
                "(" +
                "SELECT * FROM rows WHERE id_table=" + idTable + " AND key_id like '%\\_" + keyIDContinient + "\\_%' AND level=2 " +
                "), row_value_table AS " +
                "( " +
                "SELECT * FROM rows_table INNER JOIN row_value ON rows_table.id=row_value.id_row " +
                ") " +
                "SELECT timestamp, SUM(row_value_table.value) as value " +
                "FROM row_value_table " +
                "GROUP BY row_value_table.timestamp " +
                "ORDER BY timestamp DESC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Value> list = new List<Row_Value>();
                while (reader.Read())
                {
                    Row_Value row_value = new Row_Value();

                    row_value.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    row_value.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    row_value.ID_Row = row.ID;
                    rowValueService.Insert_Update(row_value);
                }
                conn.Close();
            }
            catch (Exception e)
            {
                Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
        }

        public void Load_NhapKhau_QuocGia_MatHang()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Nhap khau quoc gia - mat hang";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            CountryService countryService = new CountryService();
            Table table;
            string keyIDMacroType = "nhap-khau-quoc-gia-mat-hang";
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
                string keyIDTable = "nhap-khau";
                string valueType = "Value";
                string tableType = "";
                table = new Table();
                table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
                if (table.KeyID == null)
                {
                    continue;
                }



                string unit = table.Unit;
                string keyIDRow = keyIDTable + "_" + valueType + "_" + tableType;

                for (int rowIndex = 1; rowIndex <= 1500; rowIndex++)
                {
                    string rowName = (excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    if (rowName == null || rowName == "")
                    {
                        break;
                    }
                    double value = 0;
                    try
                    {
                        string stringValue = (excelSheet.Cells[rowIndex, 4] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                        value = double.Parse(stringValue.Replace(".", "")) / 1000000;
                    }
                    catch
                    {

                    }
                    if (!rowName.Any(char.IsLower))
                    {

                        CountryModel countryModel = new CountryModel();
                        countryModel = countryService.Get_By_Country_Name_Vi(rowName);
                        if (countryModel.Continent == null)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + countryModel.Continent + " " + countryModel.Country + " " + rowName);
                            break;
                        }
                        try
                        {
                            keyIDRow = keyIDTable + "_" + valueType + "_" + tableType + "_" + Tool.titleToKeyID(countryModel.Continent) + "_" + countryModel.Key_ID;
                        }
                        catch
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + countryModel.Continent + " " + countryModel.Country);
                            break;
                        }
                        Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyIDRow, unit, table.Id);
                        if (row == null || row.ID == 0)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRow);
                            break;
                        }
                        Row_Value row_Value = new Row_Value();
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;
                        row_Value.ID_Row = row.ID;
                        rowValueService.Insert_Update(row_Value);
                    }
                    else
                    {
                        string keyIDRowLevel3 = "";
                        try
                        {
                            keyIDRowLevel3 = keyIDRow + "_" + Tool.titleToKeyID(rowName);
                        }
                        catch
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRow + " " + rowName);
                            break;
                        }
                        Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyIDRowLevel3, unit, table.Id);
                        if (row == null || row.ID == 0)
                        {
                            Form1._Form1.updateTxtBug("Không tìm thấy: " + keyIDRowLevel3);
                            break;
                        }
                        Row_Value row_Value = new Row_Value();
                        row_Value.TimeStamp = timeStamp;
                        row_Value.Value = value;
                        row_Value.ID_Row = row.ID;
                        rowValueService.Insert_Update(row_Value);

                    }


                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("nhap-khau", "", "Value", keyIDMacroType);
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-a", "nhap-khau_Value__chau-a");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-au", "nhap-khau_Value__chau-au");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-my", "nhap-khau_Value__chau-my");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-uc", "nhap-khau_Value__chau-uc");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-dai-duong", "nhap-khau_Value__chau-dai-duong");
            Insert_Sum_Value_By_Continient(tableValue.Id, "chau-phi", "nhap-khau_Value__chau-phi");



            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            List<Row> listRowRoot = new List<Row>();

            foreach (Row row in listRow)
            {
                if (row.Level == 1)
                {
                    listRowRoot.Add(row);
                }
            }
            foreach (Row row in listRowRoot)
            {
                List<Row> listRowLevel2 = new List<Row>();
                foreach (Row row_level2 in listRow)
                {
                    if (row_level2.Key_ID.Contains(row.Key_ID) && row_level2.Level == 2)
                    {
                        listRowLevel2.Add(row_level2);
                    }
                }
                row.Rows = listRowLevel2;
            }
            foreach (Row row in listRowRoot)
            {
                foreach (Row row_level2 in row.Rows)
                {
                    List<Row> listRowLevel3 = new List<Row>();
                    foreach(Row row_level3 in listRow)
                    {
                        if (row_level3.Key_ID.Contains(row_level2.Key_ID) && row_level3.Level == 3)
                        {
                            listRowLevel3.Add(row_level3);
                        }
                    }
                    row_level2.Rows = listRowLevel3;
                }
            }
            int stt = 0;
            foreach (Row row in listRowRoot)
            {
                row.Stt = stt;
                rowService.Update(row);
                stt = stt + 1;
                foreach (Row row_level2 in row.Rows)
                {
                    row_level2.Stt = stt;
                    rowService.Update(row_level2);
                    stt = stt + 1;
                    foreach (Row row_level3 in row_level2.Rows)
                    {
                        row_level3.Stt = stt;
                        rowService.Update(row_level3);
                        stt = stt + 1;
                    }
                }
            }
            //Row rowChauA = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-a", "Triệu USD", tableValue.Id);
                             //Row rowChauAu = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-au", "Triệu USD", tableValue.Id);
                             //Row rowChauMy = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-my", "Triệu USD", tableValue.Id);
                             //Row rowChauPhi = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-phi", "Triệu USD", tableValue.Id);
                             //Row rowChauUc = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-uc", "Triệu USD", tableValue.Id);
                             //Row rowChauDaiDuong = rowService.Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__chau-duong", "Triệu USD", tableValue.Id);
                             //rowChauA.Rows = new List<Row>();
                             //rowChauAu.Rows = new List<Row>();
                             //rowChauMy.Rows = new List<Row>();
                             //rowChauPhi.Rows = new List<Row>();
                             //rowChauUc.Rows = new List<Row>();
                             //rowChauDaiDuong.Rows = new List<Row>();
                             //foreach (Row row in listRow)
                             //{
                             //    if (row.Key_ID.Contains("_chau-a_"))
                             //    {
                             //        rowChauA.Rows.Add(row);
                             //    }
                             //    if (row.Key_ID.Contains("_chau-au_"))
                             //    {
                             //        rowChauAu.Rows.Add(row);
                             //    }
                             //    if (row.Key_ID.Contains("_chau-my_"))
                             //    {
                             //        rowChauMy.Rows.Add(row);
                             //    }
                             //    if (row.Key_ID.Contains("_chau-uc_"))
                             //    {
                             //        rowChauUc.Rows.Add(row);
                             //    }
                             //    if (row.Key_ID.Contains("_chau-dai-duong_"))
                             //    {
                             //        rowChauDaiDuong.Rows.Add(row);
                             //    }
                             //}
                             //listRow = new List<Row>();
                             //listRow.Add(rowChauA);
                             //listRow.Add(rowChauUc);
                             //listRow.Add(rowChauMy);
                             //listRow.Add(rowChauUc);
                             //listRow.Add(rowChauPhi);



            Form1._Form1.updateTxtBug("Hoàn tất Xuat Nhap Khau");

        }

        public void Load_NhapKhau_MatHang_QuocGia()
        {
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            TableService tableService = new TableService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("nhap-khau", "", "Value", "nhap-khau-quoc-gia-mat-hang");
            Table tableMatHangQuocGia = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("nhap-khau", "", "Value", "nhap-khau-mat-hang-theo-quoc-gia");

            rowService.Clear(tableMatHangQuocGia);

            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            List<Row> listRowCountry = new List<Row>();
            List<Row> listMatHang = new List<Row>();

            foreach (Row row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                row.Row_Values = listRowValue;
            }
            //lấy danh sách mặt hàng
            foreach (Row row in listRow)
            {
                string[] tmpKeyID = row.Key_ID.Split('_');
                if (tmpKeyID.Length >= 6)
                {
                    Row rowMatHang = listMatHang.Find(x => x.Key_ID == tmpKeyID[5]);
                    if (rowMatHang == null)
                    {
                        rowMatHang = new Row();
                        rowMatHang.Key_ID = tmpKeyID[5];
                        rowMatHang.Name = row.Name;
                        rowMatHang.Rows = new List<Row>();
                        listMatHang.Add(rowMatHang);
                    }

                }
            }
            //lấy danh sách mặt hàng - quốc gia
            foreach (Row row in listRow)
            {

                for (int i = 0; i < listMatHang.Count; i++)
                {

                    if (row.Key_ID.Contains(listMatHang[i].Key_ID))
                    {
                        listMatHang[i].Rows.Add(row);
                    }
                }
            }
            //Cập nhật lại tên quốc gia trong childRow
            foreach (Row row in listMatHang)
            {

                foreach (Row row_child in row.Rows)
                {
                    string[] tmp = row_child.Key_ID.Split('_');
                    string keySearch = tmp[0] + "_" + tmp[1] + "_" + tmp[2] + "_" + tmp[3] + "_" + tmp[4];
                    Row rowSearch = listRow.Find(x => x.Key_ID == keySearch);
                    row_child.Name = rowSearch.Name;
                }
            }
            int stt = 0;
            foreach (Row row in listMatHang)
            {
                row.Row_Values = rowService.Sum_Contain_KeyID_By_TimeStamp(table.Id, row.Key_ID);
                row.ID_Table = tableMatHangQuocGia.Id;
                row.Level = 2;
                row.Unit = tableMatHangQuocGia.Unit;
                row.Stt = stt;
                stt = stt + 1;
                //row.Key_ID = tableMatHangQuocGia.KeyID+"_"+
                row.ID = rowService.Insert(row);
                foreach (Row_Value row_Value in row.Row_Values)
                {
                    row_Value.ID_Row = row.ID;
                    rowValueService.Insert_Update(row_Value);
                }
                foreach (Row rowChild in row.Rows)
                {
                    rowChild.Row_Values = rowService.Sum_Contain_KeyID_By_TimeStamp(table.Id, rowChild.Key_ID);
                    rowChild.ID_Table = tableMatHangQuocGia.Id;
                    rowChild.Level = 3;
                    rowChild.Unit = tableMatHangQuocGia.Unit;
                    rowChild.Stt = stt;
                    stt = stt + 1;
                    rowChild.ID = rowService.Insert(rowChild);

                    foreach (Row_Value row_Value in rowChild.Row_Values)
                    {
                        row_Value.ID_Row = rowChild.ID;
                        rowValueService.Insert_Update(row_Value);
                    }
                }
            }
            //chỉnh sửa lại keyid
            foreach (Row row in listMatHang)
            {
                foreach (Row rowChild in row.Rows)
                {

                    string[] tmp = rowChild.Key_ID.Split('_');
                    if (tmp.Length > 2)
                    {
                        rowChild.Key_ID = tmp[5] + "_" + tmp[4];
                    }

                }
            }
            //chỉnh sửa lại keyid theo table
            for (int i = 0; i < listMatHang.Count; i++)
            {
                listMatHang[i].Key_ID = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + listMatHang[i].Key_ID;
                try
                {
                    listMatHang[i].Rows = listMatHang[i].Rows.OrderByDescending(item => item.Row_Values[0].Value).ToList();
                }
                catch { }
                foreach (Row rowChild in listMatHang[i].Rows)
                {
                    string[] tmp = rowChild.Key_ID.Split('_');
                    rowChild.Key_ID = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + tmp[1] + "_" + tmp[0];
                    rowService.Update(rowChild);
                }
            }
            //Lấy tổng bên table xuất khẩu mặt hàng lắp sang
            Table tableXuatKhauMatHang = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("nhap-khau", "", "Value", "nhap-khau-theo-mat-hang");
            List<Row> listRowXuatKhauMatHang = rowService.Get_Rows_By_IdTable(tableXuatKhauMatHang.Id);
            foreach (Row row in listMatHang)
            {
                string[] tmpRow = row.Key_ID.Split('_');
                foreach (Row rowXuatKhauMatHang in listRowXuatKhauMatHang)
                {
                    string[] tmpRowXuatKhauMatHang = rowXuatKhauMatHang.Key_ID.Split('_');
                    if (tmpRow[tmpRow.Length - 1] == tmpRowXuatKhauMatHang[tmpRowXuatKhauMatHang.Length - 1])
                    {
                        List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(rowXuatKhauMatHang.ID);
                        row.Row_Values = listRowValue.OrderByDescending(x => x.TimeStamp).ToList();
                        foreach (Row_Value row_Value in row.Row_Values)
                        {
                            row_Value.ID_Row = row.ID;
                        }
                    }
                }
            }
            //update lại stt
            stt = 0;
            foreach (Row row in listMatHang)
            {

                row.Stt = stt;
                stt = stt + 1;
                rowService.Update(row);
                foreach (Row_Value row_value in row.Row_Values)
                {
                    rowValueService.Insert_Update(row_value);
                }
                foreach (Row rowChild in row.Rows)
                {
                    rowChild.Stt = stt;
                    stt = stt + 1;
                    rowService.Update(rowChild);

                }
            }



        }

        public bool Check_Country_Name_Vi(string linkFolder)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + linkFolder;
            bool result = true;
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(@path);
            Worksheet excelSheet = wb.ActiveSheet;


            for (int row = 1; row < 2000; row++)
            {

                string nuoc = (excelSheet.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                if (nuoc.Any(char.IsLower))
                {
                    continue;
                }
                if (nuoc == "")
                {
                    break;
                }

                CountryService countryService = new CountryService();

                if (!countryService.Check_Exist_Country_Name_Vi(nuoc))
                {
                    Form1._Form1.updateTxtBug("Không tồn tại: " + nuoc);
                    result = false;
                }
                //countryService.Insert(chau, nuoc);

            }

            return result;


        }


        public void Load_XuatKhau_MatHang()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Xuat khau hang hoa";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;
            table = new Table();
            table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "Value", "xuat-khau-theo-mat-hang");

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);



                if (table.KeyID == null)
                {
                    continue;
                }


                int level = -1;
                int stt = 0;
                for (int rowIndex = 1; rowIndex <= 150; rowIndex++)
                {
                    string unit = table.Unit;
                    try
                    {
                        level = int.Parse((excelSheet.Cells[rowIndex, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                    }
                    catch
                    {
                        break;
                    }
                    string keyID = (excelSheet.Cells[rowIndex, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    try
                    {
                        double value = double.NaN;
                        try
                        {
                            string valuetmp = (excelSheet.Cells[rowIndex, 6] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                            valuetmp = valuetmp.Replace(".", "");
                            value = Math.Round(double.Parse(valuetmp) / 1000000, 2);
                        }
                        catch { }
                        string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                        Row row = rowService.Get_Row_By_KeyID_Unit_IDTable(keyID, "Triệu USD", table.Id);
                        row.Stt = rowIndex;
                        rowService.Update(row);
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

                wb.Close();
            }

            ToolController toolController = new ToolController();
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "MoM", "xuat-khau-theo-mat-hang");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            ////----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "", "YoY", "xuat-khau-theo-mat-hang");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            ////-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "YTD", "YoY", "xuat-khau-theo-mat-hang");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("xuat-khau", "YTD", "Value", "xuat-khau-theo-mat-hang");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }


    }
}
