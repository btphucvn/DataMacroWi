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
    class FillDataController
    {
        public void FillData(int idMacroType, string dateType, string valueType, string unit, string linkText)
        {
            string data = File.ReadAllText(linkText);
            dynamic result = JsonConvert.DeserializeObject<dynamic>(data);
            var count = result["content"]["parent"].Count;
            AllKeyService allKeyService = new AllKeyService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
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
                title = title.Replace("\"", "");
                table.Name = title;

                TableService tableService = new TableService();
                table.Id = tableService.InsertPG(table);

                bool dontHaveLevel = false;
                int level = 2;
                try
                {
                    level = result["content"]["parent"][i]["child"][0]["level"];
                }
                catch
                {
                    dontHaveLevel = true;
                }
                string id_String = "";
                string root_IDLevel1 = "";
                string root_IDLevel2 = "";
                for (int k = 0; k < result["content"]["parent"][i]["child"].Count; k++)
                {
                    string name = JsonConvert.SerializeObject(result["content"]["parent"][i]["child"][k]["name"]);
                    string nameInsert = ToolData.removeUnitFromTitle_NotLower(name);
                    string tmpUnit = ToolData.getUnitFromTitle(name);
                    string rowUnit = "";
                    name = ToolData.removeUnitFromTitle(name);
                    name = name.Replace("\"", "");
   
                    string keyID = Tool.titleToKeyID(name);

                    try
                    {
                        level = result["content"]["parent"][i]["child"][k]["level"];
                    }
                    catch { }
                    if (!dontHaveLevel)
                    {
                        if (level == 1)
                        {
                            root_IDLevel1 = table.KeyID+"_"+table.ValueType+"_"+table.TableType + "_" + keyID;
                            id_String = root_IDLevel1;
                        }
                        if (level == 2)
                        {
                            if (root_IDLevel1 == "")
                            {
                                root_IDLevel2 = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + keyID;
                            }
                            else
                            {
                                root_IDLevel2 = root_IDLevel1 + "_" + keyID;
                            }
                            id_String = root_IDLevel2;
                        }
                        if (level == 3)
                        {
                            if (root_IDLevel1 == "")
                            {
                                id_String = root_IDLevel2 + "_" + keyID;

                            }
                            else
                            {
                                id_String =  root_IDLevel2 + "_" + keyID;
                            }
                        }
                    }
                    else
                    {
                        id_String = table.KeyID + "_" + table.ValueType + "_" + table.TableType + "_" + keyID;
                    }
                    if (tmpUnit != "")
                    {
                        rowUnit = tmpUnit;
                    }
                    else
                    {
                        rowUnit = table.Unit;
                    }


                    Row row = new Row();
                    row.ID_Table = table.Id;
                    row.Key_ID = keyID;
                    if (!dontHaveLevel)
                    {
                        row.Level = level;

                    }
                    YAxisService yAxisService = new YAxisService();
                    id_String = id_String.Replace(" ", "");
                    row.Name = nameInsert.Replace("\"","");
                    row.Stt = k;
                    row.Unit = rowUnit;
                    row.ID_String = id_String;
                    row.Key_ID = id_String;
                    row.YAxis = yAxisService.GetYAxis(rowUnit);

                    row.ID = rowService.Insert(row);


                    for (int h = 0; h < result["content"]["parent"][i]["child"][k]["data"].Count; h++)
                    {

                            Row_Value row_Value = new Row_Value();
                            row_Value.ID_Row = row.ID;
                            row_Value.TimeStamp = result["content"]["parent"][i]["child"][k]["data"][h][0];
                            try
                            {
                                row_Value.Value = result["content"]["parent"][i]["child"][k]["data"][h][1];

                            }
                            catch (Exception e)
                            {
                            }
                            rowValueService.Insert(row_Value);

                    }
                    

                }

            }
        }

    }
}
