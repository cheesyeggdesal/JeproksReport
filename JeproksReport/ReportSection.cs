using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class ReportSection
    {
        public string DataSource { get; set; }
        
        public List<ReportRow> Rows { get; set; }

        public ReportRow AddRow()
        {
            if (Rows == null) Rows = new List<ReportRow>();
            var newRow = new ReportRow();
            this.Rows.Add(newRow);
            return newRow;
        }
    }
}
