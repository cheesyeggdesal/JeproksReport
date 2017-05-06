using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public class ReportTemplate
    {
        public PageOptions PageOptions { get; set; }
        
        public List<ReportColumn> Columns { get; set; }

        public List<ReportSection> Sections { get; set; }

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
            if (this.Sections == null) this.Sections = new List<ReportSection>();
            this.Sections.Add(section);
            return section;
        }

        public void Save(string fileName)
        {
            using (var streamwriter = new StreamWriter(fileName))
            {
                var serializer = new XmlSerializer(typeof(ReportTemplate));
                serializer.Serialize(streamwriter, this);
            };
        }

        public static ReportTemplate Load(string filename)
        {
            using (var reader = new StreamReader(filename))
            {
                var serializer = new XmlSerializer(typeof(ReportTemplate));
                return (ReportTemplate)serializer.Deserialize(reader);
            }
        }
    }
}
