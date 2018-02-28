using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppDevReportGenerator
{
    public class Row
    {
        public List<Cell> Cells { get; set; }
        public bool AddedToReport { get; set; }
        public bool IsDivider { get; set; }
        public string Background { get; set; }
        public int Index { get; set; }
    }
}
