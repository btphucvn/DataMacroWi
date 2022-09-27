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
    class DauTuController
    {
        public void Load_DauTu_DangKyKinhDoanh()
        {
            string linkFolder = "DataMacro\\Dau tu\\Dang ky kinh doanh";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            Table table;
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
                            if(keyID.Contains("lao-dong-dang-ky")|| keyID.Contains("von-dang-ky-binh-quan-1-doanh-nghiep"))
                            {
                                value = value / 1000;
                            }
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

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("dang-ky-kinh-doanh", "YTD", "Value", "dang-ky-kinh-doanh");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");
        }

        public void Load_DauTu_VonDauTuTuNSNN()
        {
            string linkFolder = "DataMacro\\Dau tu\\Von dau tu tu nsnn";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            listFile = Tool.Sort_File_Name_By_Date_DESC(listFile);
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

            Table table_Value_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType("von-dau-tu-tu-nsnn", "YTD", "Value", "von-dau-tu-tu-nsnn");
            rowService.Clear(table_Value_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_Value_YTD, "YTD", "Value");
        }

    }
}
