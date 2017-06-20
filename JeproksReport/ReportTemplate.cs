using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class ReportTemplate
    {
        public PageOptions PageOptions { get; set; }

        public List<ReportColumn> Columns { get; set; }

        public List<ReportSection> Sections { get; set; }
        public ReportSection Header { get; set; }

        public ReportTemplate()
        {
            PageOptions = new PageOptions();
        }

        public void InitializeColumns(params int[] colWidths)
        {
            this.Columns = colWidths.Select(col => new ReportColumn(col)).ToList();
        }

        public ReportSection AddSection()
        {
            var newSection = new ReportSection();
            return this.AddSection(newSection);
        }

        public ReportSection AddSection(ReportSection section)
        {
            if (this.Sections == null) this.Sections = new List<ReportSection>();
            this.Sections.Add(section);
            return section;
        }
    }
}
