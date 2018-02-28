using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppDevReportGenerator
{
    public class Cell
    {
        public int ColumnNumber { get; set; }
        public string Value { get; set; }
        public int ColumnWidth { get; set; }
        public string ExportName { get; set; }
    }
}
