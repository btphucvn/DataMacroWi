using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataMacroWi.Extension
{
    public static class Tool
    {
        public static List<string> Sort_File_Name_By_Date_DESC(List<string> listFile)
        {

            List<double> listTimeStamp = new List<double>();
            foreach (string fileName in listFile)
            {
                listTimeStamp.Add(Convert_DDMMYYYY_To_Timestamp(fileName.Replace(".xlsx", "")));
            }
            listTimeStamp = listTimeStamp.OrderByDescending(x=>x).ToList();
            List<string> listResult = new List<string>();
            foreach(double item in listTimeStamp)
            {
                listResult.Add(Convert_TimeStamp_To_DateString(item)+".xlsx");
            }
            return listResult;
        }
        public static string RemoveAccents(this string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            text = text.Normalize(NormalizationForm.FormD);
            char[] chars = text
                .Where(c => CharUnicodeInfo.GetUnicodeCategory(c)
                != UnicodeCategory.NonSpacingMark).ToArray();

            return new string(chars).Normalize(NormalizationForm.FormC);
        }
        public static string titleToKeyID(this string phrase)
        {
            phrase = phrase.Replace("Đ", "d");
            phrase = phrase.Replace("đ", "d");

            // Remove all accents and make the string lower case.  
            string output = phrase.RemoveAccents().ToLower();

            // Remove all special characters from the string.  
            output = Regex.Replace(output, @"[^A-Za-z0-9\s-]", "");

            // Remove all additional spaces in favour of just one.  
            output = Regex.Replace(output, @"\s+", " ").Trim();

            // Replace all spaces with the hyphen.  
            output = Regex.Replace(output, @"\s", "-");

            // Return the slug.  
            return output;
        }
        
        public static List<string> GetAllFolderName(string link)
        {
            List<string> listFileName = new List<string>();
            DirectoryInfo d = new DirectoryInfo(link); //Assuming Test is your Folder

            FileInfo[] Files = d.GetFiles("*.*"); //Getting Text files
            var filtered = Files.Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden));


            foreach (FileInfo file in filtered)
            {
                listFileName.Add(file.Name);
            }
            return listFileName;
        }
        public static double Convert_DDMMYYYY_To_Timestamp(string date)
        {
            DateTime myDate = DateTime.ParseExact(date+" 00:00:00", "dd-MM-yyyy HH:mm:ss",
                                       System.Globalization.CultureInfo.InvariantCulture);
            double timeStamp = ((DateTimeOffset)myDate).ToUnixTimeSeconds();
            return timeStamp*1000;

        }

        public static double Get_Previous_Month_Date_By_TimeStamp(double timeStamp)
        {
            string date = Convert_TimeStamp_To_DateString(timeStamp);
            string[] arrString = date.Split('-');
            int month = int.Parse(arrString[1]);
            int year  = int.Parse(arrString[2]);
            month = month - 1;
            if (month == 0)
            {
                month = 12;
                year = year - 1;
            }
            string monthResult = month.ToString();
            if (month < 10)
            {
                monthResult = "0" + month;
            }
            date =  "01" + "-" + monthResult + "-" + year.ToString();
            return Convert_DDMMYYYY_To_Timestamp(date);
        }

        public static double Get_Previous_Year_TimeStamp(double timestamp)
        {
            string date = Convert_TimeStamp_To_DateString(timestamp);
            string[] arrString = date.Split('-');
            string month = arrString[1];
            int year = int.Parse(arrString[2]);
            year = year - 1;
            date = "01" + "-" + month +"-"+ year.ToString();
            
            return Convert_DDMMYYYY_To_Timestamp(date);
        }


        public static string Get_Previous_Month_Date(string date)
        {
            string[] arrString = date.Split('-');
            int month = int.Parse(arrString[1]);
            int year = int.Parse(arrString[2]);
            month = month - 1;
            if (month == 0)
            {
                month = 12;
                year = year - 1;
            }
            string monthResult = month.ToString();
            if (month < 10)
            {
                monthResult = "0" + month;
            }
            return "01" + "-" + monthResult + "-" + year.ToString();
        }
        public static string Convert_TimeStamp_To_DateString(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp/1000).ToLocalTime();
            string month = "";
            string day = "";
            if (dateTime.Month < 10)
            {
                month = "0" + dateTime.Month.ToString();
            }
            else
            {
                month = dateTime.Month.ToString();
            }
            if (dateTime.Day < 10)
            {
                day = "0" + dateTime.Day.ToString();
            }
            return day+"-"+month+"-"+dateTime.Year;
        }
        public static int Get_Year_From_TimeStamp(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp / 1000).ToLocalTime();
            return dateTime.Year;
        }

        public static int Get_Month_From_TimeStamp(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp / 1000).ToLocalTime();
            return dateTime.Month;
        }

        public static int GetLastestID(List<dynamic> list )
        {
            if (list.Count == 0) { return 1; }
            if (list.Count == 1) { return 2; }

            list = list.OrderBy(q => q.ID).ToList();

            int maxID = list.Max(t => t.ID);
           
            for (int i = 1; i <= maxID; i++) { 
                if(list[i-1].ID != i)
                {
                    return i;
                }
            }

            return maxID + 1 ;
        }

        public static double Get_TimeStamp_YoY_Month_From_TimeStamp(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            string currentDate = Convert_TimeStamp_To_DateString(unixTimeStamp);
            string[] arrCurrentDate = currentDate.Split('-');

            string previousYoY = "01-" + arrCurrentDate[1] + "-" + (int.Parse(arrCurrentDate[2])-1).ToString();
            return Convert_DDMMYYYY_To_Timestamp(previousYoY);
        }
        public static double Get_Value_By_TimeStamp_From_List(double timeStamp,List<dynamic> list)
        {
            for(int i = 0; i < list.Count; i++)
            {
                if (timeStamp == list[i].TimeStamp)
                {
                    return list[i].Value;
                }
            }
            return double.NaN;
        }
        public static bool Check_Exist_List_Data(double timeStamp, List<dynamic> list)
        {
            for(int i = 0; i < list.Count; i++)
            {
                if (timeStamp == list[i].TimeStamp)
                {
                    return true;
                }
            }
            return false;
        }

    }
}
