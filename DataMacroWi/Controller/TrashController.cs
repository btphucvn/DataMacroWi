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
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DataMacroWi.Controller
{
    class TrashController
    {
        public void Load_Country()
        {
            string linkFolder = "Chau.xlsx";
            string path = Directory.GetCurrentDirectory() + "\\" + linkFolder;

            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(@path);
            Worksheet excelSheet = wb.ActiveSheet;

            for (int col = 1; col <= 5; col++)
            {
                string chau = (excelSheet.Cells[1, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                for (int row = 2; row < 100; row++)
                {

                    string nuoc = (excelSheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    if (nuoc == "")
                    {
                        break;
                    }
                    CountryService countryService = new CountryService();
                    countryService.Insert(chau, nuoc);

                }
            }



        }

        public void Load_Provinces()
        {
            string linkFolder = "TinhThanh.xlsx";
            string path = Directory.GetCurrentDirectory() + "\\" + linkFolder;

            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(@path);
            Worksheet excelSheet = wb.ActiveSheet;

            for (int col = 1; col <= 3; col++)
            {
                string region = (excelSheet.Cells[1, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                for (int row = 2; row < 100; row++)
                {

                    string province = (excelSheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Text.ToString();
                    if (province == "")
                    {
                        break;
                    }
                    ProvinceService countryService = new ProvinceService();
                    countryService.Insert(province, region);

                }
            }



        }

        public void Check_Country_Name_Vi()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Xuat Khau Quoc Gia - Mat hang\\01-08-2022.xlsx";
            string path = Directory.GetCurrentDirectory() + "\\" + linkFolder;

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
                }
                //countryService.Insert(chau, nuoc);

            }




        }

        public void Update_KeyID()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Xuat Khau Quoc Gia - Mat hang";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            ProvinceService provinceService = new ProvinceService();
            Table table;
            string keyIDMacroType = "xuat-khau-quoc-gia-mat-hang";
            string keyIDTable = "xuat-khau";
            string valueType = "Value";
            string tableType = "";
            table = new Table();
            table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            for(int i = 0; i < listRow.Count; i++)
            {
                //string[] tmpArr = listRow[i].Key_ID.Split('_');

                Row row = new Row();
                row = listRow[i];
                //row.Key_ID = Tool.titleToKeyID(row)
         
            }
            CountryService countryService = new CountryService();
            List<CountryModel> listCountry = countryService.Get_All();
            for(int i = 0; i < listCountry.Count; i++)
            {
                CountryModel countryModel = listCountry[i];
                if(countryModel.Country== "Bermuda")
                {
                    int test = 0;
                }
                countryModel.Key_ID = Tool.titleToKeyID(countryModel.Country);
                countryService.Update_KeyID(countryModel);
            }


        }

        public void Update_KeyID_Row_Xuat_Khau_Quoc_Gia_Mat_Hang()
        {
            string linkFolder = "DataMacro\\Xuat Nhap Khau\\Xuat Khau Quoc Gia - Mat hang";
            List<string> listFile = Tool.GetAllFolderName(linkFolder);
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            YAxisService yAxisService = new YAxisService();
            ToolController toolController = new ToolController();
            ProvinceService provinceService = new ProvinceService();
            CountryService countryService = new CountryService();
            Table table;
            string keyIDMacroType = "xuat-khau-quoc-gia-mat-hang";
            string keyIDTable = "xuat-khau";
            string valueType = "Value";
            string tableType = "";
            table = new Table();
            table = tableService.Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(keyIDTable, tableType, valueType, keyIDMacroType);
            List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
            for (int i = 0; i < listRow.Count; i++)
            {
                string[] tmpArr = listRow[i].Key_ID.Split('_');

                Row row = new Row();
                row = listRow[i];
                string key_id_country = "";
                CountryModel country = countryService.Get_By_Country_KeyID(Tool.titleToKeyID(tmpArr[4]));
                if (country == null)
                {
                    continue;
                }
                if (country.Continent == null)
                {
                    continue;
                }
                tmpArr[3] = Tool.titleToKeyID(country.Continent);
                key_id_country = string.Join("_", tmpArr);
                row.Key_ID = key_id_country;
                rowService.Update(row);
                //row.Key_ID = Tool.titleToKeyID(row)

            }


        }

    }
}
