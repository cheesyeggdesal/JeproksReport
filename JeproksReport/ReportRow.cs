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

        public ReportRow ParseDataFields(object datarow)
        {
            var newobjs = this.Objects.Select(obj =>
            {
                if (obj is DataField)
                {
                    var datafield = obj as DataField;

                    string value = "";

                    var prop = datarow.GetType().GetProperty(datafield.Field);
                    var datavalue = prop.GetValue(datarow);

                    if (string.IsNullOrEmpty(datafield.Format))
                    {
                        value = datavalue.ToString();
                    }
                    else
                    {
                        if (datavalue != null)
                        {
                            value = string.Format("{0:" + datafield.Format + "}", datavalue);
                        }
                        else
                        {
                            value = "";
                        }
                    }


                    return new Label()
                    {
                        Alignment = datafield.Alignment,
                        BackColor = datafield.BackColor,
                        Bold = datafield.Bold,
                        Borders = datafield.Borders,
                        Color = datafield.Color,
                        FontFamily = datafield.FontFamily,
                        FontSize = datafield.FontSize,
                        Italic = datafield.Italic,
                        Value = value
                    };
                }
                return obj;
            });
            return new ReportRow()
            {
                BackColor = this.BackColor,
                Objects = newobjs.ToList()
            };
        }
    }
}
