using DataMacroWi.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using DataMacroWi.Service;
using DataMacroWi.Model;
using Newtonsoft.Json;

namespace DataMacroWi.Controller
{
    class LoadDataExcelController
    {
        ToolController toolController = new ToolController();
        public void Load_Data_CPI()
        {
            string linkFolder = "DataMacro\\CPI";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('_');
                string[] tmpArrFileName = tmpArr[1].Split('.');

                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                string keyIDTable = tmpArrFileName[0];
                for (int z = 3; z <= 7; z++)
                {
                    //Value
                    string valueType = (excelSheet.Cells[1, z] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    TableService tableService = new TableService();
                    Table table = tableService.Get_Table_By_KeyID_ValueType(keyIDTable, valueType);
                    RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                    RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                    RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();

                    RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                    RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                    RowDataLevel3ValueService rowDataLevel3ValueService = new RowDataLevel3ValueService();
                    string unit = "";
                    if (valueType == "Value")
                    {
                        unit = "Điểm";
                    }
                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int idLevel1 = -1;
                    int idLevel2 = -1;
                    for (int i = 2; i <= 19; i++)
                    {
                        int level = int.Parse((excelSheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                        string name = (excelSheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                        double value = double.NaN;
                        try
                        {
                            value = double.Parse((excelSheet.Cells[i, z] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            if (valueType == "YoY" || valueType == "YoY Ave" || valueType == "MoM")
                            {
                                value = value - 100;
                            }
                        }
                        catch { };
                        if (name.Contains("Lạm phát cơ bản"))
                        {
                            int test = 0;
                        }
                        if (level == 1 && value != double.NaN)
                        {
                            //Row_Data_Level1 rowDataLevel1 = new Row_Data_Level1();
                            //rowDataLevel1.IdTable = table.Id;
                            //rowDataLevel1.KeyID = Tool.titleToKeyID(name);
                            //rowDataLevel1.Name = name;
                            //rowDataLevel1.Stt = row;
                            //rowDataLevel1.Unit = unit;
                            string keyID = Tool.titleToKeyID(name);
                            Row_Data_Level1 rowDataLevel1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(table.Id, keyID);
                            idLevel1 = rowDataLevel1.Id;

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idLevel1;
                            row_Data_Level1_Value.TimeStamp = timeStamp;
                            row_Data_Level1_Value.Value = value;
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);
                        }

                        if (level == 2 && value != double.NaN)
                        {
                            //Row_Data_Level2 rowDataLevel2 = new Row_Data_Level2();
                            //rowDataLevel2.IdRowDataLevel1 = idLevel1;
                            //rowDataLevel2.KeyID = Tool.titleToKeyID(name);
                            //rowDataLevel2.Name = name;
                            //rowDataLevel2.Stt = row;
                            //rowDataLevel2.Unit = unit;
                            //idLevel2 = rowDataLevel2Service.InsertPG(rowDataLevel2);
                            Row_Data_Level2 rowDataLevel2 = rowDataLevel2Service.Get_RowDataLevel2_By_IdRowLevel1_KeyID(idLevel1, Tool.titleToKeyID(name));
                            idLevel2 = rowDataLevel2.Id;

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idLevel2;
                            row_Data_Level2_Value.TimeStamp = timeStamp;
                            row_Data_Level2_Value.Value = value;
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);
                        }
                        if (level == 3 && value != double.NaN)
                        {
                            //Row_Data_Level3 rowDataLevel3 = new Row_Data_Level3();
                            //rowDataLevel3.IdRowDataLevel2 = idLevel2;
                            //rowDataLevel3.KeyID = Tool.titleToKeyID(name);
                            //rowDataLevel3.Name = name;
                            //rowDataLevel3.Stt = row;
                            //rowDataLevel3.Unit = unit;
                            //int idLevel3 = rowDataLevel3Service.InsertPG(rowDataLevel3);
                            string keyID = Tool.titleToKeyID(name);
                            Row_Data_Level3 rowDataLevel3 = rowDataLevel3Service.Get_RowDataLevel3_By_IdRowLevel2_KeyID(idLevel2, keyID);
                            int idLevel3 = rowDataLevel3.Id;

                            Row_Data_Level3_Value row_Data_Level3_Value = new Row_Data_Level3_Value();
                            row_Data_Level3_Value.IdRowDataLevel3 = idLevel3;
                            row_Data_Level3_Value.TimeStamp = timeStamp;
                            row_Data_Level3_Value.Value = value;
                            rowDataLevel3ValueService.InsertPG(row_Data_Level3_Value);

                        }
                    }
                }

                wb.Close();
            }
        }



        


        public void Load_TieuDung_CPI()
        {
            string linkFolder = "DataMacro\\Tieu Dung\\CPI";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');

                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 4; col <= 8; col++)
                {


                    string keyIDTable = (excelSheet.Cells[1, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[2, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[3, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    TableService tableService = new TableService();

                    Table table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }
                    RowService rowService = new RowService();
                    RowValueService rowValueService = new RowValueService();

                    string unit = "";

                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int level = -1;

                    for (int rowIndex = 5; rowIndex <= 30; rowIndex++)
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
                            double value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            if (row.Key_ID == null)
                            {
                                row.ID_String = keyID;
                                row.ID_Table = table.Id;
                                row.Key_ID = keyID;
                                row.Level = level;
                                row.Name = rowName;
                                row.Stt = rowIndex;
                                row.Unit = unit;
                                YAxisService yAxisService = new YAxisService();
                                row.YAxis = yAxisService.GetYAxis(unit);
                                row.ID = rowService.Insert(row);
                            }
                            Row_Value row_Value = new Row_Value();
                            row_Value.ID_Row = row.ID;
                            row_Value.TimeStamp = timeStamp;
                            if(unit == "%")
                            {
                                row_Value.Value = value-100;

                            }
                            else
                            {
                                row_Value.Value = value;
                            }
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


        }

        public void Load_TieuDung_BanLeHangHoaVaDichVu()
        {
            string linkFolder = "DataMacro\\Tieu Dung\\Ban Le Hang Hoa Va Dich Vu";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 4; col <= 5; col++)
                {


                    string keyIDTable = (excelSheet.Cells[2, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[3, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[4, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }

                    string unit = "";

                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int level = -1;

                    for (int rowIndex = 5; rowIndex <= 30; rowIndex++)
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
                            double value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
                            if (row.Key_ID == null)
                            {
                                row.ID_String = keyID;
                                row.ID_Table = table.Id;
                                row.Key_ID = keyID;
                                row.Level = level;
                                row.Name = rowName;
                                row.Stt = rowIndex;
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
                            if (unit == "%")
                            {
                                row_Value.Value = value - 100;

                            }
                            else
                            {
                                row_Value.Value = value;
                            }
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "", "Value");
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "", "MoM");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRowValue_MoM = toolController.Calculate_MoM(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableMoM.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableMoM.KeyID + "_" + tableMoM.ValueType + "_" + tableMoM.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_Value_MoM in listRowValue_MoM)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_Value_MoM.TimeStamp;
                    row_Value.Value = row_Value_MoM.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }
            //--------------------------
            //Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "YTD", "Value");
            //rowService.Clear(table_Value_YTD);
            //foreach (var row in listRow)
            //{
            //    List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
            //    List<Row_Value> listRow_Value_YoY = Calculate_YTD_Value(listRowValue);

            //    Row row_MoM = new Row();
            //    row_MoM.ID_Table = table_Value_YTD.Id;
            //    string[] arrIDStringRow = row.ID_String.Split('_');
            //    string id_string = "";
            //    for (int i = 3; i < arrIDStringRow.Length; i++)
            //    {
            //        id_string = id_string + "_" + arrIDStringRow[i];
            //    }
            //    row_MoM.ID_String = table_Value_YTD.KeyID + "_" + table_Value_YTD.ValueType + "_" + table_Value_YTD.TableType + id_string;
            //    row_MoM.Key_ID = row_MoM.ID_String;
            //    row_MoM.Level = row.Level;
            //    row_MoM.Name = row.Name;
            //    row_MoM.Stt = row.Stt;
            //    row_MoM.Unit = "%";
            //    row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
            //    row_MoM.ID = rowService.Insert(row_MoM);
            //    foreach (var row_YTD_YoY in listRow_Value_YoY)
            //    {
            //        Row_Value row_Value = new Row_Value();
            //        row_Value.ID_Row = row_MoM.ID;
            //        row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
            //        row_Value.Value = row_YTD_YoY.Value;
            //        rowValueService.Insert_Update(row_Value);
            //    }

            //}
            //----------------------------
            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "", "YoY");
            rowService.Clear(tableYoY);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YoY = toolController.Calculate_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableYoY.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableYoY.KeyID + "_" + tableYoY.ValueType + "_" + tableYoY.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }
            //----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YTD_YoY = toolController.Calculate_YTD_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = table_YoY_YTD.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = table_YoY_YTD.KeyID + "_" + table_YoY_YTD.ValueType + "_" + table_YoY_YTD.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YTD_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }

        }

        public void Load_DauTu_DangKyKinhDoanh()
        {
            string linkFolder = "DataMacro\\Dau tu\\Dang ky kinh doanh";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 6; col <= 7; col++)
                {


                    string keyIDTable = (excelSheet.Cells[7, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[8, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[9, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }

                    string unit = "";

                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int level = -1;

                    for (int rowIndex = 10; rowIndex <= 30; rowIndex++)
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
                            double value = double.Parse((excelSheet.Cells[rowIndex, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                            string rowName = (excelSheet.Cells[rowIndex, 3] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                            Row row = rowService.Get_Row_By_KeyID(keyID);
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
                                row.Unit = unit;
                                row.YAxis = yAxisService.GetYAxis(unit);
                                row.ID = rowService.Insert(row);
                            }
                            Row_Value row_Value = new Row_Value();
                            row_Value.ID_Row = row.ID;
                            row_Value.TimeStamp = timeStamp;
                            if (col == 6)
                            {
                                row_Value.TimeStamp = Tool.Get_Previous_Month_Date_By_TimeStamp(timeStamp);
                            }
                            if (unit == "%")
                            {
                                row_Value.Value = value - 100;

                            }
                            else
                            {
                                row_Value.Value = value;
                            }
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType("dang-ky-kinh-doanh", "", "Value");
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType("dang-ky-kinh-doanh", "", "MoM");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRowValue_MoM = toolController.Calculate_MoM(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableMoM.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableMoM.KeyID + "_" + tableMoM.ValueType + "_" + tableMoM.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_Value_MoM in listRowValue_MoM)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_Value_MoM.TimeStamp;
                    row_Value.Value = row_Value_MoM.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }
            //--------------------------
            //Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("ban-le-hhdv", "YTD", "Value");
            //rowService.Clear(table_Value_YTD);
            //foreach (var row in listRow)
            //{
            //    List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
            //    List<Row_Value> listRow_Value_YoY = Calculate_YTD_Value(listRowValue);

            //    Row row_MoM = new Row();
            //    row_MoM.ID_Table = table_Value_YTD.Id;
            //    string[] arrIDStringRow = row.ID_String.Split('_');
            //    string id_string = "";
            //    for (int i = 3; i < arrIDStringRow.Length; i++)
            //    {
            //        id_string = id_string + "_" + arrIDStringRow[i];
            //    }
            //    row_MoM.ID_String = table_Value_YTD.KeyID + "_" + table_Value_YTD.ValueType + "_" + table_Value_YTD.TableType + id_string;
            //    row_MoM.Key_ID = row_MoM.ID_String;
            //    row_MoM.Level = row.Level;
            //    row_MoM.Name = row.Name;
            //    row_MoM.Stt = row.Stt;
            //    row_MoM.Unit = "%";
            //    row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
            //    row_MoM.ID = rowService.Insert(row_MoM);
            //    foreach (var row_YTD_YoY in listRow_Value_YoY)
            //    {
            //        Row_Value row_Value = new Row_Value();
            //        row_Value.ID_Row = row_MoM.ID;
            //        row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
            //        row_Value.Value = row_YTD_YoY.Value;
            //        rowValueService.Insert_Update(row_Value);
            //    }

            //}

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("dang-ky-kinh-doanh", "", "YoY");
            rowService.Clear(tableYoY);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YoY = toolController.Calculate_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableYoY.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableYoY.KeyID + "_" + tableYoY.ValueType + "_" + tableYoY.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }


            //----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("dang-ky-kinh-doanh", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YTD_YoY = toolController.Calculate_YTD_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = table_YoY_YTD.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = table_YoY_YTD.KeyID + "_" + table_YoY_YTD.ValueType + "_" + table_YoY_YTD.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YTD_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }

        }

        public void Load_DauTu_VonDauTuTuNSNN()
        {
            string linkFolder = "DataMacro\\Dau tu\\Von dau tu tu nsnn";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 4; col <= 5; col++)
                {


                    string keyIDTable = (excelSheet.Cells[5, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[6, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[7, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }

                    string unit = "";

                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int level = -1;

                    for (int rowIndex = 9; rowIndex <= 50; rowIndex++)
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
                                if(unit == "")
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
                            if (unit == "%")
                            {
                                row_Value.Value = value - 100;

                            }
                            else
                            {
                                row_Value.Value = value;
                            }
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType("von-dau-tu-tu-nsnn", "", "Value");
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType("von-dau-tu-tu-nsnn", "", "MoM");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRowValue_MoM = toolController.Calculate_MoM(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableMoM.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableMoM.KeyID + "_" + tableMoM.ValueType + "_" + tableMoM.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_Value_MoM in listRowValue_MoM)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_Value_MoM.TimeStamp;
                    row_Value.Value = row_Value_MoM.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }


            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("von-dau-tu-tu-nsnn", "", "YoY");
            rowService.Clear(tableYoY);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YoY = toolController.Calculate_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = tableYoY.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = tableYoY.KeyID + "_" + tableYoY.ValueType + "_" + tableYoY.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }


            //----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("von-dau-tu-tu-nsnn", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRow_YTD_YoY = toolController.Calculate_YTD_YoY(listRowValue);

                Row row_MoM = new Row();
                row_MoM.ID_Table = table_YoY_YTD.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                row_MoM.ID_String = table_YoY_YTD.KeyID + "_" + table_YoY_YTD.ValueType + "_" + table_YoY_YTD.TableType + id_string;
                row_MoM.Key_ID = row_MoM.ID_String;
                row_MoM.Level = row.Level;
                row_MoM.Name = row.Name;
                row_MoM.Stt = row.Stt;
                row_MoM.Unit = "%";
                row_MoM.YAxis = yAxisService.GetYAxis(row_MoM.Unit);
                row_MoM.ID = rowService.Insert(row_MoM);
                foreach (var row_YTD_YoY in listRow_YTD_YoY)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = row_MoM.ID;
                    row_Value.TimeStamp = row_YTD_YoY.TimeStamp;
                    row_Value.Value = row_YTD_YoY.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }

        }

        public void Load_SanXuat_IIP()
        {
            string linkFolder = "DataMacro\\San Xuat\\IIP";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 4; col <= 8; col++)
                {


                    string keyIDTable = (excelSheet.Cells[4, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[5, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[6, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }

                    string unit = "";

                    if (valueType == "YoY" || valueType == "MoM" || valueType == "YTD" || valueType == "YoY Ave")
                    {
                        unit = "%";
                    }
                    int level = -1;
                    int stt = 0;
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
                            if (col == 4)
                            {
                                row_Value.TimeStamp = Tool.Get_Previous_Month_Date_By_TimeStamp(timeStamp);
                            }
                            if (unit == "%")
                            {
                                row_Value.Value = value - 100;

                            }
                            else
                            {
                                row_Value.Value = value;
                            }
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "", "Value");
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "", "MoM");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "","MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "", "YoY");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");




        }

        public void Load_SanXuat_SanPhamCongNghiep()
        {
            string linkFolder = "DataMacro\\San Xuat\\San Pham Cong Nghiep";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();

            Table table;

            foreach (var item in listFile)
            {
                string path = Directory.GetCurrentDirectory() + "\\" + linkFolder + "\\" + item;
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(@path);
                Worksheet excelSheet = wb.ActiveSheet;
                //string test = excelSheet.Cells[1, 4].Value2.ToString();
                string[] tmpArr = item.Split('.');
                double timeStamp = Tool.Convert_DDMMYYYY_To_Timestamp(tmpArr[0]);
                for (int col = 5; col <= 6; col++)
                {


                    string keyIDTable = (excelSheet.Cells[8, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string valueType = (excelSheet.Cells[9, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    string tableType = (excelSheet.Cells[10, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();

                    table = new Table();
                    table = tableService.Get_Table_By_KeyID_TableType_ValueType(keyIDTable, tableType, valueType);
                    if (table.KeyID == null)
                    {
                        continue;
                    }


                    int level = -1;
                    int stt = 0;
                    for (int rowIndex = 12; rowIndex <= 150; rowIndex++)
                    {
                        string unit = (excelSheet.Cells[rowIndex, 4] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
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
                            if (col == 5)
                            {
                                row_Value.TimeStamp = Tool.Get_Previous_Month_Date_By_TimeStamp(timeStamp);
                            }


                            row_Value.Value = value;
                            
                            rowValueService.Insert_Update(row_Value);
                        }
                        catch { }

                    }
                }
                wb.Close();
            }


            Table tableValue = tableService.Get_Table_By_KeyID_TableType_ValueType("san-pham-cong-nghiep", "", "Value");
            Table tableMoM = tableService.Get_Table_By_KeyID_TableType_ValueType("san-pham-cong-nghiep", "", "MoM");
            List<Row> listRow = rowService.Get_Rows_By_IdTable(tableValue.Id);
            rowService.Clear(tableMoM);
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("san-pham-cong-nghiep", "", "YoY");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("san-pham-cong-nghiep", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("san-pham-cong-nghiep", "YTD", "Value");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");

        }



    }
}
