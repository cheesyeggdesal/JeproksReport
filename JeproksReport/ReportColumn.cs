using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class ReportColumn
    {
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
