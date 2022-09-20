using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Model
{
    class Row
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Key_ID { get; set; }
        public int Level { get; set; }
        public string Unit { get; set; }

        public int Stt { get; set; }

        public int ID_Table { get; set; }
        public string ID_String { get; set; }
        public int YAxis { get; set; }

        public List<Row> Rows { get; set; }
        public List<Row_Value> Row_Values {get;set;}
    }
}
