using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public class ReportObject
    {
        [XmlAttribute]
        public string BackColor { get; set; }
        public Borders Borders { get; set; }
        public bool ShouldSerializeBorders()
        {
            return !(this.Borders.Top.Width == 0) ||
                   !(this.Borders.Right.Width == 0) ||
                   !(this.Borders.Bottom.Width == 0) ||
                   !(this.Borders.Left.Width == 0);
        }

        public ReportObject()
        {
            this.Borders = new Borders(0, 0, 0, 0);
        }
    }

    public class Borders
    {
        public Border Top { get; set; }
        public bool ShouldSerializeTop()
        {
            return !(Top.Width == 0);
        }
        public Border Right { get; set; }
        public bool ShouldSerializeRight()
        {
            return !(Right.Width == 0);
        }
        public Border Bottom { get; set; }
        public bool ShouldSerializeBottom()
        {
            return !(Bottom.Width == 0);
        }
        public Border Left { get; set; }
        public bool ShouldSerializeLeft()
        {
            return !(Left.Width == 0);
        }

        public Borders() { }

        public Borders(int tw, int rw, int bw, int lw)
        {
            this.Top = new Border() { Visible = tw != 0, Width = tw };
            this.Right = new Border() { Visible = rw != 0, Width = rw };
            this.Bottom = new Border() { Visible = bw != 0, Width = bw };
            this.Left = new Border() { Visible = lw != 0, Width = lw };
        }


    }
    public class Border
    {
        [XmlAttribute]
        public string Color { get; set; }
        public bool ShouldSerializeColor()
        {
            return this.Color != DefaultValues.COLOR_BORDER;
        }
        [XmlAttribute]
        public bool Visible { get; set; }
        [XmlAttribute]
        public double Width { get; set; }

        public Border()
        {
            this.Color = DefaultValues.COLOR_BORDER;
        }
    }

    public class TextBase : ReportObject
    {
        [XmlAttribute]
        public string FontFamily { get; set; }
        public bool ShouldSerializeFontFamily()
        {
            return FontFamily != DefaultValues.FONT_FAMILY;
        }
        [XmlAttribute]
        public double FontSize { get; set; }
        public bool ShouldSerializeFontSize()
        {
            return this.FontSize != DefaultValues.FONT_SIZE;
        }
        [XmlAttribute]
        public bool Bold { get; set; }
        public bool ShouldSerializeBold()
        {
            return this.Bold;
        }
        [XmlAttribute]
        public bool Italic { get; set; }
        public bool ShouldSerializeItalic()
        {
            return this.Italic;
        }
        [XmlAttribute]
        public string Color { get; set; }
        public bool ShouldSerializeColor()
        {
            return this.Color != DefaultValues.COLOR_TEXT;
        }

        [XmlAttribute]
        public StringAlignment Alignment { get; set; }
        public bool ShouldSerializeAlignment()
        {
            return this.Alignment != StringAlignment.Near;
        }

        public TextBase()
        {
            this.Color = DefaultValues.COLOR_TEXT;
            this.FontSize = DefaultValues.FONT_SIZE;
            this.FontFamily = DefaultValues.FONT_FAMILY;
            this.Alignment = StringAlignment.Near;
        }
    }

    public class Label : TextBase
    {
        [XmlText]
        public string Value { get; set; }
    }

    public class DataField : TextBase
    {
        [XmlAttribute]
        public string Field { get; set; }
        [XmlAttribute]
        public string Format { get; set; }
    }

    public class Image : ReportObject
    {
        [XmlAttribute]
        public string Src { get; set; }
    }
}
