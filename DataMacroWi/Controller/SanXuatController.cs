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
    class SanXuatController
    {
        public void Load_SanXuat_IIP()
        {
            string linkFolder = "DataMacro\\San Xuat\\IIP";
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
                        if (keyID.Contains("cong-nghiep-che-bien-che-tao_san-xuat-than-coc-san-pham-dau-mo-tinh-che_san-xuat-than-coc"))
                        {
                            int it = 0;
                        }
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
                            row.Level = level;
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
                                row.Stt = stt;
                                stt = stt + 1;
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
            toolController.Insert_Row_Value_Calculate(listRow, tableMoM, "", "MoM");

            //----------------------------

            Table tableYoY = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "", "YoY");
            rowService.Clear(tableYoY);
            toolController.Insert_Row_Value_Calculate(listRow, tableYoY, "", "YoY");


            //-----------------------------
            Table table_YoY_YTD = tableService.Get_Table_By_KeyID_TableType_ValueType("iip", "YTD", "YoY");
            rowService.Clear(table_YoY_YTD);
            toolController.Insert_Row_Value_Calculate(listRow, table_YoY_YTD, "YTD", "YoY");




        }

    }
}
