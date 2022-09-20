using DataMacroWi.Extension;
using DataMacroWi.Model;
using DataMacroWi.Service;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Controller
{
    class SanLuongController
    {
        public void FillDataGDPNam()
        {
            string data = File.ReadAllText("Du lieu vi mo/San Luong/GDP Nam.txt");
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            for (int i =0;i< count; i++)
            {
                Table table = new Table();
                string title = JsonConvert.SerializeObject(result["content"]["parent"][i]["title"]);
                table.KeyID = Tool.titleToKeyID(title);
                table.Stt = i;
                table.IdMacroType = 6;
                table.ValueType = "Value";
                table.DateType = "Year";
                TableService tableService = new TableService();
                int tableID = tableService.InsertPG(table);
                int idRowDataLevel1 = -1;
                
                int idRowDataLevel2 = -1;
                int idRowDataLevel3 = -1;

                for (int k=0;k< result["content"]["parent"][i]["child"].Count;k++)
                {
                    if (result["content"]["parent"][i]["child"][k]["level"] == 1)
                    {
                        Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                        row_Data_Level1.IdTable = tableID;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        name = ToolData.removeUnitFromTitle(name);

                        row_Data_Level1.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level1.Unit = "Tỷ";
                        RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                        idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level1_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 2)
                    {
                        Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();
                        row_Data_Level2.IdRowDataLevel1 = idRowDataLevel1;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        name = ToolData.removeUnitFromTitle(name);

                        row_Data_Level2.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level2.Unit = "Tỷ";
                        RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                        idRowDataLevel2 = rowDataLevel2Service.InsertPG(row_Data_Level2);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level2_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);

                        }
                    }
                }

            }
        }
        public void FillDataGDPHienHanh()
        {
            string data = File.ReadAllText("Du lieu vi mo/San Luong/GDP Hien Hanh Quy.txt");
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            for (int i = 0; i < count; i++)
            {
                Table table = new Table();
                string title = JsonConvert.SerializeObject(result["content"]["parent"][i]["title"]);
                table.KeyID = Tool.titleToKeyID(title);
                table.Stt = i;
                table.IdMacroType = 4;
                table.ValueType = "Value";
                table.DateType = "Year";

                TableService tableService = new TableService();
                int tableID = tableService.InsertPG(table);
                int idRowDataLevel1 = -1;
                int idRowDataLevel2 = -1;
                int idRowDataLevel3 = -1;

                for (int k = 0; k < result["content"]["parent"][i]["child"].Count; k++)
                {
                    if (result["content"]["parent"][i]["child"][k]["level"] == 1)
                    {
                        Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                        row_Data_Level1.IdTable = tableID;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level1.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level1.Unit = "Tỷ";
                        RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                        idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level1_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 2)
                    {
                        Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();
                        row_Data_Level2.IdRowDataLevel1 = idRowDataLevel1;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level2.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level2.Unit = "Tỷ";
                        RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                        idRowDataLevel2 = rowDataLevel2Service.InsertPG(row_Data_Level2);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level2_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 3)
                    {
                        Row_Data_Level3 row_Data_Level3 = new Row_Data_Level3();
                        row_Data_Level3.IdRowDataLevel2 = idRowDataLevel2;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level3.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level3.Unit = "Tỷ";
                        if(row_Data_Level3.KeyID=="khai-khoang")
                        {
                            int test = 0;
                        }
                        RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();
                        idRowDataLevel3 = rowDataLevel3Service.InsertPG(row_Data_Level3);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level3_Value row_Data_Level3_Value = new Row_Data_Level3_Value();
                            row_Data_Level3_Value.IdRowDataLevel3 = idRowDataLevel3;
                            row_Data_Level3_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level3_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel3ValueService rowDataLevel3ValueService = new RowDataLevel3ValueService();
                            rowDataLevel3ValueService.InsertPG(row_Data_Level3_Value);

                        }
                    }

                }

            }
        }
        public void FillDataGDPSoSanh()
        {
            string data = File.ReadAllText("Du lieu vi mo/San Luong/GDP So Sanh Quy.txt");
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            for (int i = 0; i < count; i++)
            {
                Table table = new Table();
                string title = JsonConvert.SerializeObject(result["content"]["parent"][i]["title"]);
                table.KeyID = Tool.titleToKeyID(title);
                table.Stt = i;
                table.IdMacroType = 5;
                table.ValueType = "Value";
                table.DateType = "Year";
                table.TableType = ToolData.getUnitFromTitle(title);
                TableService tableService = new TableService();
                int tableID = tableService.InsertPG(table);
                int idRowDataLevel1 = -1;
                int idRowDataLevel2 = -1;
                int idRowDataLevel3 = -1;

                for (int k = 0; k < result["content"]["parent"][i]["child"].Count; k++)
                {
                    if (result["content"]["parent"][i]["child"][k]["level"] == 1)
                    {
                        Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                        row_Data_Level1.IdTable = tableID;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level1.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level1.Unit = "Tỷ";
                        RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                        idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level1_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 2)
                    {
                        Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();
                        row_Data_Level2.IdRowDataLevel1 = idRowDataLevel1;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level2.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level2.Unit = "Tỷ";
                        RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                        idRowDataLevel2 = rowDataLevel2Service.InsertPG(row_Data_Level2);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level2_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 3)
                    {
                        Row_Data_Level3 row_Data_Level3 = new Row_Data_Level3();
                        row_Data_Level3.IdRowDataLevel2 = idRowDataLevel2;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level3.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level3.Unit = "Tỷ";
                        RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();
                        idRowDataLevel3 = rowDataLevel3Service.InsertPG(row_Data_Level3);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level3_Value row_Data_Level3_Value = new Row_Data_Level3_Value();
                            row_Data_Level3_Value.IdRowDataLevel3 = idRowDataLevel3;
                            row_Data_Level3_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level3_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel3ValueService rowDataLevel3ValueService = new RowDataLevel3ValueService();
                            rowDataLevel3ValueService.InsertPG(row_Data_Level3_Value);

                        }
                    }

                }

            }
        }
        public void FillDataPMI()
        {
            string data = File.ReadAllText("Du lieu vi mo/San Luong/PMI.txt");
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            for (int i = 0; i < count; i++)
            {
                Table table = new Table();
                string title = JsonConvert.SerializeObject(result["content"]["parent"][i]["title"]);
                table.KeyID = Tool.titleToKeyID(title);
                table.Stt = i;
                table.IdMacroType = 3;
                table.ValueType = "Value";
                table.DateType = "Year";

                TableService tableService = new TableService();
                int tableID = tableService.InsertPG(table);
                int idRowDataLevel1 = -1;
                int idRowDataLevel2 = -1;
                int idRowDataLevel3 = -1;

                for (int k = 0; k < result["content"]["parent"][i]["child"].Count; k++)
                {
                    if (result["content"]["parent"][i]["child"][k]["level"] == 2)
                    {
                        Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                        row_Data_Level1.IdTable = tableID;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level1.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level1.Unit = "Tỷ";
                        RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                        idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level1_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 3)
                    {
                        Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();
                        row_Data_Level2.IdRowDataLevel1 = idRowDataLevel1;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level2.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level2.Unit = "Tỷ";
                        RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                        idRowDataLevel2 = rowDataLevel2Service.InsertPG(row_Data_Level2);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level2_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 4)
                    {
                        Row_Data_Level3 row_Data_Level3 = new Row_Data_Level3();
                        row_Data_Level3.IdRowDataLevel2 = idRowDataLevel2;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        row_Data_Level3.KeyID = Tool.titleToKeyID(name);
                        row_Data_Level3.Unit = "Tỷ";
                        RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();
                        idRowDataLevel3 = rowDataLevel3Service.InsertPG(row_Data_Level3);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level3_Value row_Data_Level3_Value = new Row_Data_Level3_Value();
                            row_Data_Level3_Value.IdRowDataLevel3 = idRowDataLevel3;
                            row_Data_Level3_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level3_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel3ValueService rowDataLevel3ValueService = new RowDataLevel3ValueService();
                            rowDataLevel3ValueService.InsertPG(row_Data_Level3_Value);

                        }
                    }

                }

            }
        }

        public string getUnitFromTitle(string title)
        {
            if (title.Contains("(Dự án)"))
            {
                return "Dự án";
            }
            if (title.Contains("(Triệu USD)"))
            {
                return "Triệu USD";
            }
            if (title.Contains("(Nghìn tấn)"))
            {
                return "Nghìn tấn";
            }
            if (title.Contains("(Triệu m3)"))
            {
                return "Triệu m3";
            }
            if (title.Contains("(Triệu lít)"))
            {
                return "Triệu lít";
            }
            if (title.Contains("(Triệu bao)"))
            {
                return "Triệu bao";
            }
            if (title.Contains("(Triệu m2)"))
            {
                return "Triệu m2";
            }
            if (title.Contains("(Triệu viên)"))
            {
                return "Triệu viên";
            }
            if (title.Contains("(Nghìn cái)"))
            {
                return "Nghìn cái";
            }
            if (title.Contains("(Nghìn chiếc)"))
            {
                return "Nghìn chiếc";
            }
            if (title.Contains("(Tỷ kwh)"))
            {
                return "Tỷ kwh";
            }
            if (title.Contains("(Triệu cái)"))
            {
                return "Triệu cái";
            }
            if (title.Contains("(Triệu đôi)"))
            {
                return "Triệu đôi";
            }
            if (title.Contains("(Nghìn tỷ)"))
            {
                return "Triệu cái";
            }
            if (title.Contains("(Chiếc)"))
            {
                return "Chiếc";
            }
            if (title.Contains("(DN)"))
            {
                return "DN";
            }
            if (title.Contains("(Triệu tấn)"))
            {
                return "Triệu tấn";
            }
            if (title.Contains("(USD mn)"))
            {
                return "USD mn";
            }
            return "";
        }
        public void FillData(int idMacroType,string dateType,string valueType,string unit,string linkText)
        {
            string data = File.ReadAllText(linkText);
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            AllKeyService allKeyService = new AllKeyService();
            for (int i = 0; i < count; i++)
            {
                Table table = new Table();
                string title = JsonConvert.SerializeObject(result["content"]["parent"][i]["title"]);
                table.TableType = ToolData.GetTableTypeFromTitle(title);
                title = ToolData.removeUnitFromTitle(title);
                title = ToolData.RemoveTableTypeFromTitle(title);
                table.KeyID = Tool.titleToKeyID(title);
                table.Stt = i;
                table.IdMacroType = idMacroType;
                table.ValueType = valueType;
                table.DateType = dateType;
                table.Unit = unit;
                
                if (!allKeyService.CheckExitsAllKeyByKeyID(table.KeyID))
                {
                    allKeyService.InsertPG(table.KeyID, title);
                }
                TableService tableService = new TableService();
                int tableID = tableService.InsertPG(table);
                int idRowDataLevel1 = -1;
                int idRowDataLevel2 = -1;
                int idRowDataLevel3 = -1;
                bool dontHaveLevel = false ;
                int level = 2;
                try { 
                    level = result["content"]["parent"][i]["child"][0]["level"];
                }
                catch { 
                    dontHaveLevel = true; 
                }
                if (level != 1)
                {
                    Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                    row_Data_Level1.IdTable = tableID;
                    string name = "All";
                    row_Data_Level1.KeyID = "all";
                    row_Data_Level1.Unit = unit;
                    if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level1.KeyID))
                    {
                        allKeyService.InsertPG(row_Data_Level1.KeyID, name);
                    }

                    RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                    idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                }
                for (int k = 0; k < result["content"]["parent"][i]["child"].Count; k++)
                {

                    if (result["content"]["parent"][i]["child"][k]["level"] == 1)
                    {
                        Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
                        row_Data_Level1.IdTable = tableID;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        string tmpUnit = ToolData.getUnitFromTitle(name);

                        name = ToolData.removeUnitFromTitle(name);
                        row_Data_Level1.KeyID = Tool.titleToKeyID(name);

                        if (tmpUnit != "")
                        {
                            row_Data_Level1.Unit = tmpUnit;
                        }
                        else
                        {
                            row_Data_Level1.Unit = unit;
                        }

                        if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level1.KeyID))
                        {
                            allKeyService.InsertPG(row_Data_Level1.KeyID, name);
                        }

                        RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
                        idRowDataLevel1 = rowDataLevel1Service.InsertPG(row_Data_Level1);
                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = idRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level1_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel1ValueService rowDataLevel1ValueService = new RowDataLevel1ValueService();
                            rowDataLevel1ValueService.InsertPG(row_Data_Level1_Value);

                        }
                    }
                    
                    if (result["content"]["parent"][i]["child"][k]["level"] == 2 || dontHaveLevel)
                    {
                        Row_Data_Level2 row_Data_Level2 = new Row_Data_Level2();
                        row_Data_Level2.IdRowDataLevel1 = idRowDataLevel1;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        string tmpUnit = ToolData.getUnitFromTitle(name);

                        name = ToolData.removeUnitFromTitle(name);
                        row_Data_Level2.KeyID = Tool.titleToKeyID(name);
                        if (tmpUnit != "")
                        {
                            row_Data_Level2.Unit = tmpUnit;
                        }
                        else
                        {
                            row_Data_Level2.Unit = unit;
                        }
                        if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level2.KeyID))
                        {
                            allKeyService.InsertPG(row_Data_Level2.KeyID, name);
                        }

                        RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
                        idRowDataLevel2 = rowDataLevel2Service.InsertPG(row_Data_Level2);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = idRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level2_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel2ValueService rowDataLevel2ValueService = new RowDataLevel2ValueService();
                            rowDataLevel2ValueService.InsertPG(row_Data_Level2_Value);

                        }
                    }
                    if (result["content"]["parent"][i]["child"][k]["level"] == 3)
                    {
                        Row_Data_Level3 row_Data_Level3 = new Row_Data_Level3();
                        row_Data_Level3.IdRowDataLevel2 = idRowDataLevel2;
                        string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                        string tmpUnit = ToolData.getUnitFromTitle(name);
                        name = ToolData.removeUnitFromTitle(name);
                        row_Data_Level3.KeyID = Tool.titleToKeyID(name);
                        if (tmpUnit != "")
                        {
                            row_Data_Level3.Unit = tmpUnit;
                        }
                        else
                        {
                            row_Data_Level3.Unit = unit;
                        }
                        if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level3.KeyID))
                        {
                            allKeyService.InsertPG(row_Data_Level3.KeyID, name);
                        }
                        RowDataLevel3Service rowDataLevel3Service = new RowDataLevel3Service();
                        idRowDataLevel3 = rowDataLevel3Service.InsertPG(row_Data_Level3);

                        for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                        {

                            Row_Data_Level3_Value row_Data_Level3_Value = new Row_Data_Level3_Value();
                            row_Data_Level3_Value.IdRowDataLevel3 = idRowDataLevel3;
                            row_Data_Level3_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Data_Level3_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];
                            }
                            catch (Exception e)
                            {

                            }
                            RowDataLevel3ValueService rowDataLevel3ValueService = new RowDataLevel3ValueService();
                            rowDataLevel3ValueService.InsertPG(row_Data_Level3_Value);

                        }
                    }

                }

            }
        }

    }
}
