using DataMacroWi.Model;
using DataMacroWi.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Extension
{
    public static class ToolData
    {
        private static string[] arrUnit = { "(Tỷ)","(Tỷ đồng)", "(Dự án)", "(Triệu USD)", "(Nghìn Người)"
                , "(Nghìn tấn)", "(Triệu m3)", "(Triệu lít)"
        ,"(Triệu bao)","(Triệu m2)","(Triệu viên)","(Nghìn cái)","(Nghìn chiếc)"
        ,"(Tỷ kwh)","(Triệu cái)","(Triệu đôi)","(Nghìn tỷ)","(Chiếc)","(DN)","(Triệu tấn)","(USD mn)","(Triệu tấn/Km)","(Triệu tấn.Km)","(triệu hk.km)"};

        private static string[] arrTableType = { "(YTD)","(TTM)" };
        public static string RemoveTableTypeFromTitle(string title)
        {
            string lowerTitle = title.ToLower();
            for (int i = 0; i < arrTableType.Length; i++)
            {
                if (title.Contains(arrTableType[i]) || lowerTitle.Contains(arrTableType[i].ToLower()))
                {

                    title = title.ToLower().Replace(arrTableType[i].ToLower(), "").Replace("\"","").Trim();
                    return FirstLetterToUpper(title);
                }
            }
            return FirstLetterToUpper(title.ToLower());
        }



        public static string GetTableTypeFromTitle(string title)
        {
            if (title.Contains("YTD")|| title.Contains("ytd"))
            {
                return "YTD";
            }
            if (title.Contains("TTM") || title.Contains("ttm"))
            {
                return "TTM";
            }
            return "";
        }
        

        public static string getUnitFromTitle(string title)
        {
            for(int i = 0; i < arrUnit.Length; i++)
            {
                if (title.Contains(arrUnit[i]))
                {
                    string result = arrUnit[i].Replace("(","");
                    result = result.Replace(")", "");
                    return result;
                }
            }
            return "";
        }
        public static string removeUnitFromTitle(string title)
        {
            string lowerTitle = title.ToLower();
            for (int i = 0; i < arrUnit.Length; i++)
            {
                if (title.Contains(arrUnit[i])|| lowerTitle.Contains(arrUnit[i].ToLower()))
                {
                    title = title.ToLower().Replace(arrUnit[i].ToLower(), "").Replace("\"", "").Trim();
                    return FirstLetterToUpper(title);
                }
            }
            return FirstLetterToUpper(title.ToLower());
        }
        public static string removeUnitFromTitle_NotLower(string title)
        {
            //string lowerTitle = title.ToLower();
            for (int i = 0; i < arrUnit.Length; i++)
            {
                if (title.Contains(arrUnit[i]))
                {
                    title = title.Replace(arrUnit[i], "").Replace("\"", "").Trim();
                    return title;
                }
            }
            return title;
        }
        private static string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }

       
    }
}
