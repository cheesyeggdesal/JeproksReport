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

        public List<ReportObject> Objects { get; set; }

        public ReportRow()
        {
            this.Objects = new List<ReportObject>();
        }

        public ReportRow FromData(object data, Dictionary<string, string> lastValues)
        {
            var newobjs = this.Objects.Select(obj =>
            {
                if (obj is DataField)
                {
                    var datafield = obj as DataField;

                    string value = "";

                    var prop = data.GetType().GetProperty(datafield.Field);
                    var datavalue = prop.GetValue(data);

                    if (string.IsNullOrEmpty(datafield.Format))
                    {
                        value = (datavalue ?? "").ToString();
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
                    if (datafield.SuppressIfDuplicate)
                    {
                        if (!lastValues.ContainsKey(datafield.Field))
                        {
                            lastValues.Add(datafield.Field, null);
                        }
                        else
                        {
                            if (lastValues[datafield.Field] == value)
                            {
                                value = "";
                            }
                            lastValues[datafield.Field] = value;
                        }
                    }
                    return new Label()
                    {
                        ColSpan = datafield.ColSpan,
                        RowSpan = datafield.RowSpan,
                        Style = new TextBaseStyle()
                        {
                            Alignment = datafield.Style.Alignment,
                            VerticalAlignment = datafield.Style.VerticalAlignment,
                            BackColor = datafield.Style.BackColor,
                            Bold = datafield.Style.Bold,
                            Borders = datafield.Style.Borders,
                            Color = datafield.Style.Color,
                            FontFamily = datafield.Style.FontFamily,
                            FontSize = datafield.Style.FontSize,
                            Italic = datafield.Style.Italic
                        },
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

        public Label AddLabel(string value)
        {
            var label = new Label();
            label.Value = value;

            this.Objects.Add(label);
            return label;
        }

        public Label AddLabel(string value, Font font, string color = DefaultValues.COLOR_TEXT,
            int colspan = 1, int rowspan = 1, StringAlignment alignment = StringAlignment.Near)
        {
            var label = new Label();
            label.Value = value;
            label.Style.FontFamily = font.FontFamily.Name;
            label.Style.FontSize = font.Size;
            label.Style.Bold = font.Bold;
            label.Style.Italic = font.Italic;
            label.ColSpan = colspan;
            label.RowSpan = rowspan;
            label.Style.Color = color;
            label.Style.Alignment = alignment;
            this.Objects.Add(label);
            return label;
        }
        public Label AddLabel(string value, int colspan = 1, int rowspan = 1)
        {
            var label = new Label();
            label.Value = value;
            label.ColSpan = colspan;
            label.RowSpan = rowspan;
            this.Objects.Add(label);
            return label;
        }
        public Label AddLabel(string value, TextBaseStyle style, int colspan = 1, int rowspan = 1)
        {
            var label = new Label();
            label.Style = style;
            label.Value = value;
            label.ColSpan = colspan;
            label.RowSpan = rowspan;
            this.Objects.Add(label);
            return label;
        }

        public DataField AddDataField(string field)
        {
            var datafield = new DataField();
            datafield.Field = field;

            this.Objects.Add(datafield);
            return datafield;
        }

        public DataField AddDataField(string field, Font font, string format = null,
            string color = DefaultValues.COLOR_TEXT, int colspan = 1, int rowspan = 1,
            StringAlignment alignment = StringAlignment.Near)
        {
            var datafield = new DataField();
            datafield.Field = field;
            datafield.Format = format;
            datafield.Style.FontFamily = font.FontFamily.Name;
            datafield.Style.FontSize = font.Size;
            datafield.Style.Bold = font.Bold;
            datafield.Style.Italic = font.Italic;
            datafield.ColSpan = colspan;
            datafield.RowSpan = rowspan;
            datafield.Style.Color = color;
            datafield.Style.Alignment = alignment;
            this.Objects.Add(datafield);
            return datafield;
        }
        public DataField AddDataField(string field, int colspan = 1, int rowspan = 1)
        {
            var datafield = new DataField();
            datafield.Field = field;
            datafield.ColSpan = colspan;
            datafield.RowSpan = rowspan;
            this.Objects.Add(datafield);
            return datafield;
        }
        public DataField AddDataField(string field, TextBaseStyle style, string format = null, int colspan = 1, int rowspan = 1)
        {
            var datafield = new DataField();
            datafield.Style = style;
            datafield.Field = field;
            datafield.Format = format;
            datafield.ColSpan = colspan;
            datafield.RowSpan = rowspan;
            this.Objects.Add(datafield);
            return datafield;
        }
    }
}
