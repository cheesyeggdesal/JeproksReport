using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class ReportRow
    {
        public string BackColor { get; set; }

        private List<ReportObject> _objects;
        public IReadOnlyCollection<ReportObject> Objects
        {
            get
            {
                return this._objects;
            }
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
                _objects = newobjs.ToList()
            };
        }
        public T Add<T>(T reportobj) where T : ReportObject
        {
            if (this._objects == null) this._objects = new List<ReportObject>();
            this._objects.Add(reportobj);
            return reportobj;
        }
        public Label AddLabel()
        {
            return this.AddLabel(new Label());
        }
        public Label AddLabel(Label label)
        {
            this.Add(label);
            return label;
        }
        public Label AddLabel(string value)
        {
            return this.AddLabel(new Label() { Value = value });
        }
        public Label AddLabel(string value, Font font, string color = "Black", string backcolor = null, StringAlignment alignment = StringAlignment.Near)
        {
            return this.AddLabel(new Label()
            {
                Alignment = alignment,
                Value = value,
                FontSize = font.Size,
                FontFamily = font.FontFamily.Name,
                Color = color,
                BackColor = backcolor
            });
        }
        
        public DataField AddDataField(string field, string format = null, string color = "Black", string backcolor = null, StringAlignment alignment = StringAlignment.Near)
        {
            return this.Add(new DataField()
            {
                Alignment = alignment,
                Field = field,
                Format = format,
                Color = color,
                BackColor = backcolor
            });
        }
        public DataField AddDataField(string field, Font font, string format = null, string color = "Black", string backcolor = null, StringAlignment alignment = StringAlignment.Near)
        {
            return this.Add(new DataField()
            {
                Alignment = alignment,
                Field = field,
                Format = format,
                FontSize = font.Size,
                FontFamily = font.FontFamily.Name,
                Color = color,
                BackColor = backcolor
            });
        }
    }
}
