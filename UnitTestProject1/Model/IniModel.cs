using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestProject1.Model
{
    public class IniModel
    {
        public string IniType { get; set; }
        public List<Detail> IniDetail { get; set; }

    }
    public class Detail
    {
        public string IniName { get; set; }
        public List<string> Inivalue { get; set; }
    }


}
