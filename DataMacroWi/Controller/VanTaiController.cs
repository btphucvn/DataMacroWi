using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Controller
{
    class VanTaiController
    {
        public void Load_VanTai_KhachQuocTeTheoLoaiHinh()
        {
            string linkFolder = "DataMacro\\Van tai\\Khach quoc te theo loai hinh";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "Value", "khach-quoc-te-theo-loai-hinh"); 
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 5; col++)
                {


                    int level = -1;
                    int stt = 0;
                    for (int rowIndex = 2; rowIndex <= 150; rowIndex++)
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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString())/1000;
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.Stt = stt;
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                            if (col == 4)
                            {
                                row_Value.TimeStamp = Tool.Get_Previous_Month_Date_By_TimeStamp(timeStamp);
                            }


                            row_Value.Value = value;

                            rowValueService.Insert_Update(row_Value);
                        }
                        catch(Exception e) {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "MoM", "khach-quoc-te-theo-loai-hinh");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "YoY", "khach-quoc-te-theo-loai-hinh");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "YTD", "YoY", "khach-quoc-te-theo-loai-hinh");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "YTD", "Value", "khach-quoc-te-theo-loai-hinh");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }

        public void Load_VanTai_KhachQuocTeTheoQuocGia()
        {
            string linkFolder = "DataMacro\\Van tai\\Khach quoc te theo quoc gia";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table  = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "Value", "khach-quoc-te-theo-quoc-gia"); 
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 5; col++)
                {


                    int level = -1;
                    int stt = 0;
                    for (int rowIndex = 2; rowIndex <= 150; rowIndex++)
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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString()) / 1000;
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                            if (col == 4)
                            {
                                row_Value.TimeStamp = Tool.Get_Previous_Month_Date_By_TimeStamp(timeStamp);
                            }


                            row_Value.Value = value;

                            rowValueService.Insert_Update(row_Value);
                        }
                        catch (Exception e)
                        {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "MoM", "khach-quoc-te-theo-quoc-gia");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "", "YoY", "khach-quoc-te-theo-quoc-gia");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "YTD", "YoY", "khach-quoc-te-theo-quoc-gia");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("khach-quoc-te", "YTD", "Value", "khach-quoc-te-theo-quoc-gia");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }

        public void Load_VanTai_VanTaiHanhKhach()
        {
            string linkFolder = "DataMacro\\Van tai\\Van chuyen hanh khach";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hk", "", "Value", "van-chuyen-hanh-khach");
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 4; col++)
                {


                    int level = -1;
                    int stt = 0;
                    for (int rowIndex = 2; rowIndex <= 150; rowIndex++)
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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                        catch (Exception e)
                        {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hk", "", "MoM", "van-chuyen-hanh-khach");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hk", "", "YoY", "van-chuyen-hanh-khach");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hk", "YTD", "YoY", "van-chuyen-hanh-khach");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hk", "YTD", "Value", "van-chuyen-hanh-khach");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }

        public void Load_VanTai_LuanChuyenHanhKhach()
        {
            string linkFolder = "DataMacro\\Van tai\\Luan chuyen hanh khach";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hk", "", "Value", "luan-chuyen-hanh-khach");
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 4; col++)
                {


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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                        catch (Exception e)
                        {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hk", "", "MoM", "luan-chuyen-hanh-khach");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hk", "", "YoY", "luan-chuyen-hanh-khach");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hk", "YTD", "YoY", "luan-chuyen-hanh-khach");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hk", "YTD", "Value", "luan-chuyen-hanh-khach");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }

        public void Load_VanTai_VanChuyenHangHoa()
        {
            string linkFolder = "DataMacro\\Van tai\\Van chuyen hang hoa";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hh", "", "Value", "van-chuyen-hang-hoa");
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 4; col++)
                {


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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                        catch (Exception e)
                        {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hh", "", "MoM", "van-chuyen-hang-hoa");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hh", "", "YoY", "van-chuyen-hang-hoa");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hh", "YTD", "YoY", "van-chuyen-hang-hoa");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("van-chuyen-hh", "YTD", "Value", "van-chuyen-hang-hoa");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }

        public void Load_VanTai_LuanChuyenHangHoa()
        {
            string linkFolder = "DataMacro\\Van tai\\Luan chuyen hang hoa";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hh", "", "Value", "luan-chuyen-hang-hoa");
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

                for (int col = 4; col <= 4; col++)
                {


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
                                value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            }
                            catch { }
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            row.YAxis = yAxisService.GetYAxis(unit);
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
                        catch (Exception e)
                        {
                            Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                        }
                    }
                }

                wb.Close();
            }

            ToolController toolController = new ToolController();

            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hh", "", "MoM", "luan-chuyen-hang-hoa");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hh", "", "YoY", "luan-chuyen-hang-hoa");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hh", "YTD", "YoY", "luan-chuyen-hang-hoa");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("luan-chuyen-hh", "YTD", "Value", "luan-chuyen-hang-hoa");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }


    }
}
