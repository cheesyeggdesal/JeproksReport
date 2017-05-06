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

        private List<ReportSection> _sections;
        public List<ReportSection> Sections
        {
            get
            {
                return this._sections;
            }
        }

        public ReportTemplate()
        {
            PageOptions = new PageOptions();
            this.Columns = new List<ReportColumn>();
        }

        public ReportSection AddSection()
        {
            var newSection = new ReportSection();
            return this.AddSection(newSection);
        }

        public ReportSection AddSection(ReportSection section)
        {
            if (this._sections == null) this._sections = new List<ReportSection>();
            this._sections.Add(section);
            return section;
        }
    }
}
