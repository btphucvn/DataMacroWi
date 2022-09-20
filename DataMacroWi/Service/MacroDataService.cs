using DataMacroWi.DB;
using DataMacroWi.Model;
using Newtonsoft.Json;
using Npgsql;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class MacroDataService
    {
        public async Task<string> Get_Data_Macro(string keyIDMacroType,string valueType)
        {
            string url = "http://localhost:8080/api/get-table-by-keyidmacrotype?key_id_macro_type=" 
                + keyIDMacroType
                + "&value_type="+ valueType;
            Uri myUri = new Uri(url);

            var options = new RestClientOptions()
            {
                BaseUrl = myUri,
                Timeout = 1000000  // 1 second. or whatever time you want.
            };
            var client = new RestClient(options);
            var request = new RestRequest("", Method.Get);
            var response = await client.ExecuteAsync<string>(request);

            var result = JsonConvert.DeserializeObject<dynamic>(response.Content);
            string stringResult = JsonConvert.SerializeObject(result);


            return response.Content;
        }
        public async Task<List<string>> Get_All_Value_Type_By_KeyIDMacro(string keyIDMacro) {
            var client = new RestClient("http://localhost:8080/api/get-value-type-by-keyidmacrotype?key_id_macro=" + keyIDMacro);
            var request = new RestRequest("", Method.Get);
            var response = await client.ExecuteAsync<string>(request);

            var result = JsonConvert.DeserializeObject<dynamic>(response.Content);
            List<string> listValue = new List<string>();

            for(int i =0;i<result["data"].Count;i++)
            {
                string value = JsonConvert.SerializeObject(result["data"][i]);
                value = value.Replace("\"", "");
                listValue.Add(value);
            }

            return listValue;
        }
        public async Task FillMacroData()
        {
            MacroTypeService macroTypeService = new MacroTypeService();
            List<MacroType> listMacroType = macroTypeService.Get_All_MacroType();

            for(int i =0;i< listMacroType.Count; i++)
            {
                List<string> listValue = await Get_All_Value_Type_By_KeyIDMacro(listMacroType[i].KeyID);

                for(int k = 0; k < listValue.Count; k++)
                {
                    try
                    {
                        MacroData data = new MacroData();
                        data.KeyID = listMacroType[i].KeyID;
                        data.ValueType = listValue[k];
                        data.Data = await Get_Data_Macro(listMacroType[i].KeyID, listValue[k]);
                        data.Data = data.Data.Replace("\\", "");
                        Insert(data);
                    }
                    catch(Exception e)
                    {
                        string test = e.Message;
                    }

                }
            }
        }
        public int Insert(MacroData macroData)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into macro_datas(key_id_macro_type,value_type,data) values('"
                + macroData.KeyID + "','"
                + macroData.ValueType + "','"
                + macroData.Data + "') RETURNING id;";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();
                int id = (int)cmd.ExecuteScalar();

                conn.Close();
                return id;
            }
            catch (Exception a)
            {
                string err = a.Message;
            }
            finally
            {
                conn.Close();
            }
            return -1;

        }

    }
}
