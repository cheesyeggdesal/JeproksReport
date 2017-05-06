using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class ReportObject
    {
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
        public string Color { get; set; }
        public bool Visible { get; set; }
        public double Width { get; set; }

        public Border()
        {
            this.Color = DefaultValues.COLOR_BORDER;
        }
    }

    public class TextBase : ReportObject
    {
        public string FontFamily { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string Color { get; set; }
        public StringAlignment Alignment { get; set; }

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
        public string Value { get; set; }

        public Label() { }
        public Label(string value, Font font, StringAlignment alignment = StringAlignment.Near)
        {
            this.Value = value;
            this.FontSize = font.Size;
            this.FontFamily = font.FontFamily.Name;
            this.Bold = font.Bold;
            this.Italic = font.Italic;
            this.Alignment = alignment;
        }
    }

    public class DataField : TextBase
    {
        public string Field { get; set; }
        public string Format { get; set; }
    }

    public class Image : ReportObject
    {
        public string Src { get; set; }
    }
}
