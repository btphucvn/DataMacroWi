using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Controller
{
    class ToolController
    {
        public void Insert_Row_Value_Calculate(List<Row> listRow, Table table, string tableType, string valueType)
        {
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            RowService rowService = new RowService();

            foreach (var row in listRow)
            {
                List<Row_Value> listRowValue_Calculate = rowValueService.Get_Row_Value_By_IDRow(row.ID);
                List<Row_Value> listRowValue = new List<Row_Value>();
                if (valueType == "MoM" && tableType == "")
                {
                    listRowValue = Calculate_MoM(listRowValue_Calculate);
                }
                if (valueType == "YoY" && tableType == "")
                {
                    listRowValue = Calculate_YoY(listRowValue_Calculate);

                }
                if (valueType == "YoY" && tableType == "YTD")
                {
                    listRowValue = Calculate_YTD_YoY(listRowValue_Calculate);

                }
                if (valueType == "Value" && tableType == "YTD")
                {
                    listRowValue = Calculate_YTD_Value(listRowValue_Calculate);

                }
                if (valueType == "Value" && tableType == "")
                {
                    listRowValue = Calculate_YTD_Value_To_Value(listRowValue_Calculate);
                }

                Row rowTable = new Row();
                rowTable.ID_Table = table.Id;
                string[] arrIDStringRow = row.ID_String.Split('_');
                string id_string = "";
                for (int i = 3; i < arrIDStringRow.Length; i++)
                {
                    id_string = id_string + "_" + arrIDStringRow[i];
                }
                rowTable.ID_String = table.KeyID + "_" + table.ValueType + "_" + table.TableType + id_string;
                rowTable.Key_ID = rowTable.ID_String;
                rowTable.Level = row.Level;
                rowTable.Name = row.Name;
                rowTable.Stt = row.Stt;
                rowTable.Unit = "%";
                rowTable.YAxis = yAxisService.GetYAxis(rowTable.Unit);
                rowTable.ID = rowService.Insert(rowTable);
                if (listRowValue == null)
                {
                    continue;
                }
                foreach (var row_Value_Calculate in listRowValue)
                {
                    Row_Value row_Value = new Row_Value();
                    row_Value.ID_Row = rowTable.ID;
                    row_Value.TimeStamp = row_Value_Calculate.TimeStamp;
                    row_Value.Value = row_Value_Calculate.Value;
                    rowValueService.Insert_Update(row_Value);
                }

            }

        }
        private List<Row_Value> Calculate_YTD_Value_To_Value(List<Row_Value> list)
        {
            if (list.Count == 0)
            {
                return null;
            }
            List<Row_Value> listResult = new List<Row_Value>();
            list = list.OrderByDescending(q => q.TimeStamp).ToList();
            int flagMonth = Tool.Get_Month_From_TimeStamp(list[list.Count - 1].TimeStamp);
            for (int i = 0; i<list.Count-2; i++)
            {
                double value = double.NaN;
                if (flagMonth == 1)
                {
                    value = list[i].Value;
                }
                else
                {
                    value = list[i].Value - list[i + 1].Value;
                }
                Row_Value result = new Row_Value();
                result.Value = value;
                result.TimeStamp = list[i].TimeStamp;
                listResult.Add(result);
            }
            listResult = listResult.OrderByDescending(q => q.TimeStamp).ToList();
            return listResult;
        }

        public List<Row_Value> Calculate_YTD_YoY(List<Row_Value> list)
        {
            if(list.Count==0)
            {
                return null;
            }
            List<Row_Value> listRowValue = Calculate_YTD_Value(list);
            List<Row_Value> listResult = new List<Row_Value>();
            foreach (Row_Value row_Value in listRowValue)
            {

                Row_Value row_Value_Previous = listRowValue.FirstOrDefault(x
                    => x.TimeStamp == Tool.Get_Previous_Year_TimeStamp(row_Value.TimeStamp));
                if (row_Value_Previous != null)
                {
                    Row_Value row_Value_Result = new Row_Value();
                    row_Value_Result.TimeStamp = row_Value.TimeStamp;
                    row_Value_Result.Value = Math.Round(((row_Value.Value - row_Value_Previous.Value)/ row_Value_Previous.Value)*100, 2);
                    if (row_Value_Previous.Value==0)
                    {
                        row_Value_Result.Value = 0;
                    }

                    listResult.Add(row_Value_Result);
                }

            }
            return listResult;
        }

        public List<Row_Value> Calculate_YTD_Value(List<Row_Value> list)
        {
            List<Row_Value> listResult = new List<Row_Value>();
            list = list.OrderByDescending(q => q.TimeStamp).ToList();
            double ytd = 0;
            if (list.Count == 0)
            {
                return null;
            }
            int flagYear = Tool.Get_Year_From_TimeStamp(list[list.Count - 1].TimeStamp);

            for (int i = list.Count - 1; i >= 0; i--)
            {
                if (flagYear != Tool.Get_Year_From_TimeStamp(list[i].TimeStamp))
                {
                    flagYear = Tool.Get_Year_From_TimeStamp(list[i].TimeStamp);
                    ytd = list[i].Value;
                }
                else
                {
                    ytd = ytd + list[i].Value;
                }
                Row_Value result = new Row_Value();
                result.Value = ytd;
                result.TimeStamp = list[i].TimeStamp;
                listResult.Add(result);
            }
            listResult = listResult.OrderByDescending(q => q.TimeStamp).ToList();
            return listResult;
        }

        public List<dynamic> Calculate_TTM(List<dynamic> list)
        {
            List<dynamic> listResult = new List<dynamic>();
            double ttm = 0;
            for (int i = list.Count - 12; i >= 0; i--)
            {
                ttm = list[i].Value + list[i + 1].Value + list[i + 2].Value + list[i + 3].Value + list[i + 4].Value
                     + list[i + 5].Value + list[i + 6].Value + list[i + 7].Value + list[i + 8].Value
                    + list[i + 9].Value + list[i + 10].Value + list[i + 11].Value;
                dynamic result = new System.Dynamic.ExpandoObject();
                result.Value = ttm;
                result.TimeStamp = list[i].TimeStamp;
                listResult.Add(result);
            }
            return listResult;
        }

        public  List<Row_Value> Calculate_MoM(List<Row_Value> list)
        {
            List<Row_Value> listResult = new List<Row_Value>();
            list = list.OrderByDescending(q => q.TimeStamp).ToList();
            double mom = 0;
            for (int i = list.Count - 2; i >= 0; i--)
            {

                mom = Math.Round((list[i].Value / list[i + 1].Value) * 100 - 100, 2);
                if (list[i].Value == 0 || list[i + 1].Value==0)
                {
                    mom = 0;
                }
                Row_Value result = new Row_Value();
                result.Value = mom;
                result.TimeStamp = list[i].TimeStamp;
                if (list[i].Value == 0 && list[i + 1].Value == 0)
                {
                    result.Value = double.NaN;
                }
                listResult.Add(result);
            }
            return listResult;
        }

        public  List<Row_Value> Calculate_YoY(List<Row_Value> list)
        {
            List<Row_Value> listResult = new List<Row_Value>();
            list = list.OrderByDescending(q => q.TimeStamp).ToList();

            for (int i = list.Count - 1; i >= 0; i--)
            {
                double timeStamp_YoY = Tool.Get_TimeStamp_YoY_Month_From_TimeStamp(list[i].TimeStamp);
                Row_Value row_Value_Previous = list.FirstOrDefault(x
                        => x.TimeStamp == timeStamp_YoY);

                if (row_Value_Previous != null)
                {

                    Row_Value result = new Row_Value();
                    if(list[i].Value== 9.539)
                    {
                        int test = 0;
                    }
                    result.Value = Math.Round(((list[i].Value - row_Value_Previous.Value) / row_Value_Previous.Value) * 100, 2);

                    result.TimeStamp = list[i].TimeStamp;
                    if (row_Value_Previous.Value == 0)
                    {
                        result.Value = 0;
                    }
                    listResult.Add(result);
                }
            }
            return listResult;
        }
    }
}
