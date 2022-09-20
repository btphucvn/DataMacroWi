using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Model
{
    class Table
    {
        public int Id { get; set; }
        public string KeyID { get; set; }
        public string Name { get; set; }

        public string DateType { get; set; }

        public string TableType { get; set; }

        public int Stt { get; set; }
        public string Unit { get; set; }
        public string ValueType { get; set; }
        public int IdMacroType { get; set; }

    }
}
