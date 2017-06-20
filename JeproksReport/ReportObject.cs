using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public class Borders
    {
        public Border Top { get; set; }
        public Border Right { get; set; }
        public Border Bottom { get; set; }
        public Border Left { get; set; }

        public Borders() { }

        public Borders(double tw, double rw, double bw, double lw)
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

    public class ReportObject
    {
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }

        public ReportObject()
        {
            this.ColSpan = 1;
            this.RowSpan = 1;
        }
    }
    public class Style
    {
        public string BackColor { get; set; }
        public Borders Borders { get; set; }
        public Style()
        {
            this.Borders = new Borders(0, 0, 0, 0);
        }
    }

    public class TextBase : ReportObject
    {
        public TextBaseStyle Style { get; set; }

        public TextBase()
        {
            this.Style = new TextBaseStyle();
        }
    }
    public class TextBaseStyle : Style
    {
        public string FontFamily { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string Color { get; set; }
        public StringAlignment Alignment { get; set; }
        public StringAlignment VerticalAlignment { get; set; }

        public TextBaseStyle()
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
    }

    public class DataField : TextBase
    {
        public string Field { get; set; }
        public string Format { get; set; }
        public bool SuppressIfDuplicate { get; set; }
    }

    public class Image : ReportObject
    {
        public string Src { get; set; }
    }

    public static class Extensions
    {
        public static TextBaseStyle Align(this TextBaseStyle style, StringAlignment alignment)
        {
            return new TextBaseStyle()
            {
                Alignment = alignment,
                BackColor = style.BackColor,
                Bold = style.Bold,
                Borders = style.Borders,
                Color = style.Color,
                FontFamily = style.FontFamily,
                FontSize = style.FontSize,
                Italic = style.Italic,
                VerticalAlignment = style.VerticalAlignment
            };
        }
        public static TextBaseStyle SetBorder(this TextBaseStyle style, Borders borders)
        {
            return new TextBaseStyle()
            {
                Alignment = style.Alignment,
                BackColor = style.BackColor,
                Bold = style.Bold,
                Borders = borders,
                Color = style.Color,
                FontFamily = style.FontFamily,
                FontSize = style.FontSize,
                Italic = style.Italic,
                VerticalAlignment = style.VerticalAlignment
            };
        }
        public static TextBaseStyle SetBorder(this TextBaseStyle style,
            double top, double right, double bottom, double left)
        {
            return new TextBaseStyle()
            {
                Alignment = style.Alignment,
                BackColor = style.BackColor,
                Bold = style.Bold,
                Borders = new Borders(top, right, bottom, left),
                Color = style.Color,
                FontFamily = style.FontFamily,
                FontSize = style.FontSize,
                Italic = style.Italic,
                VerticalAlignment = style.VerticalAlignment
            };
        }
    }
}
