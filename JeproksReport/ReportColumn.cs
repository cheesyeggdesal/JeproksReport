using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public class ReportColumn
    {
        [XmlAttribute]
        public double Width { get; set; }

        public ReportColumn()
        {
            this.Width = 100;
        }

        public ReportColumn(int colWidth)
        {
            this.Width = colWidth;
        }
    }
}
