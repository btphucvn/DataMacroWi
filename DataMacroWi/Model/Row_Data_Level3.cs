using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Model
{
    class Row_Data_Level3
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string KeyID { get; set; }
        public string Unit { get; set; }
        public int Stt { get; set; }

        public int IdRowDataLevel2 { get; set; }
    }
}
