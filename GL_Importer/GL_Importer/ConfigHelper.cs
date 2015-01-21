using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace GL_Importer
{
    class ConfigHelper
    {
        public int Date { get; set; }
        public int seg1 { get; set; }
        public int seg2 { get; set; }
        public int seg3 { get; set; }
        public int Amount { get; set; }
        public int lineItem { get; set; }

        public ConfigHelper()
        {
                Date = Int32.Parse(ConfigurationManager.AppSettings["Date"]);
                seg1 = Int32.Parse(ConfigurationManager.AppSettings["Seg1"]);
                seg2 = Int32.Parse(ConfigurationManager.AppSettings["Seg2"]);
                seg3 = Int32.Parse(ConfigurationManager.AppSettings["Seg3"]);
                Amount = Int32.Parse(ConfigurationManager.AppSettings["Amount"]);
                lineItem = Int32.Parse(ConfigurationManager.AppSettings["lineItem"]);
        }
 
    }
}
