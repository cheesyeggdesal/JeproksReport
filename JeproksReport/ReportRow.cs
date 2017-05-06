using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public class ReportRow
    {
        [XmlAttribute]
        public string BackColor { get; set; }

        [XmlArrayItem(typeof(DataField))]
        [XmlArrayItem(typeof(Image))]
        [XmlArrayItem(typeof(Label))]
        //[XmlArrayItem(typeof(TextBox))]
        //[XmlArrayItem(typeof(ReportObject), ElementName = "Rectangle")]
        public List<ReportObject> Objects { get; set; }
        public bool ShouldSerializeObjects()
        {
            return this.Objects.Count > 0;
        }
        public ReportRow()
        {
            this.Objects = new List<ReportObject>();
        }
    }
}
