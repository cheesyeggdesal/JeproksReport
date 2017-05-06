using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public class ReportSection
    {
        [XmlAttribute]
        public string DataSource { get; set; }

        [XmlElement("ReportRow")]
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
