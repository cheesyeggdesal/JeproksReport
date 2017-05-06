using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport.PDFSharp
{
    public class PdfSharpWriter : IReportWriter
    {
        ReportTemplate _template;
        Dictionary<string, IEnumerable<object>> _datasource;
        PdfDocument _document;
        PdfPage _currentPage;
        XGraphics _currentGraphics;
        double _currentYPos;
        double _currentXPos;

        int _currentColumnIndex;

        public void Write(ReportTemplate reportTemplate, Stream stream, Dictionary<string, IEnumerable<object>> datasource = null)
        {
            this._template = reportTemplate;
            this._datasource = datasource;

            this._document = new PdfDocument();
            this.addPage();

            foreach (var section in reportTemplate.Sections)
            {
                this.RenderSection(section);
            }

            this._document.Save(stream, false);
        }

        void addPage()
        {
            this._currentPage = this._document.AddPage();
            this._currentPage.Orientation =
                    this._template.PageOptions.PageOrientation == PageOrientation.Portrait ?
                    PdfSharp.PageOrientation.Portrait :
                    PdfSharp.PageOrientation.Landscape;
            this._currentYPos = this._template.PageOptions.Margins.Top;
            this._currentXPos = this._template.PageOptions.Margins.Left;
            if (this._currentGraphics != null) this._currentGraphics.Dispose();
            this._currentGraphics = XGraphics.FromPdfPage(this._currentPage);
        }

        void RenderSection(ReportSection section)
        {
            if (!string.IsNullOrEmpty(section.DataSource))
            {
                if (this._datasource.ContainsKey(section.DataSource))
                {
                    foreach (var datarow in this._datasource[section.DataSource])
                    {
                        RenderRows(section.Rows.Select(row => row.ParseDataFields(datarow)));
                    }
                }
            }
            else
            {
                RenderRows(section.Rows);
            }
        }

        void RenderRows(IEnumerable<ReportRow> rows)
        {
            foreach (var row in rows)
            {
                _currentColumnIndex = 0;
                double highesttHeight = 0;

                if (row.Objects.Count > 0)
                {
                    highesttHeight = row.Objects.Select(x =>
                    {
                        var height = 0d;
                        var columnWidth = this._template.Columns.ElementAt(_currentColumnIndex++).Width;
                        //check if column will overlap the right margin
                        if (_currentXPos + columnWidth > this._currentPage.Width - this._template.PageOptions.Margins.Right)
                        {
                            var drawWidth = this._currentPage.Width - this._template.PageOptions.Margins.Right;
                            columnWidth -= (_currentXPos + columnWidth) - drawWidth;
                        }

                        if (_currentColumnIndex > this._template.Columns.Count) _currentColumnIndex = 1;

                        if (x is Label)
                        {
                            var xlabel = x as Label;
                            height = this.measureString(xlabel.Value, new XFont(xlabel.FontFamily, xlabel.FontSize), columnWidth).Height;
                        }
                        this._currentXPos += columnWidth;
                        return height;
                    }).Max();
                }

                if (this._currentYPos + highesttHeight > this._currentPage.Height - this._template.PageOptions.Margins.Bottom)
                {
                    addPage();
                }

                _currentColumnIndex = 0;
                _currentXPos = this._template.PageOptions.Margins.Left;
                foreach (var obj in row.Objects)
                {
                    var columnWidth = this._template.Columns.ElementAt(_currentColumnIndex++).Width;
                    //check if column will overlap the right margin
                    if (_currentXPos + columnWidth > this._currentPage.Width - this._template.PageOptions.Margins.Right)
                    {
                        var drawWidth = this._currentPage.Width - this._template.PageOptions.Margins.Right;
                        columnWidth -= (_currentXPos + columnWidth) - drawWidth;
                    }
                    if (_currentColumnIndex > this._template.Columns.Count) _currentColumnIndex = 1;

                    if (obj is TextBase)
                    {
                        var textbase = obj as TextBase;
                        XFont font = null;
                        if (textbase.Bold && textbase.Italic)
                        {
                            font = new XFont(textbase.FontFamily, textbase.FontSize, XFontStyle.BoldItalic);
                        }
                        else if (textbase.Bold)
                        {
                            font = new XFont(textbase.FontFamily, textbase.FontSize, XFontStyle.Bold);
                        }
                        else if (textbase.Italic)
                        {
                            font = new XFont(textbase.FontFamily, textbase.FontSize, XFontStyle.Italic);
                        }
                        else
                        {
                            font = new XFont(textbase.FontFamily, textbase.FontSize);
                        }

                        var brush = new XSolidBrush(getColorFromHtml(textbase.Color));
                        XStringFormat stringFormat = new XStringFormat();
                        
                        if (obj is Label)
                        {
                            var label = obj as Label;
                            var size = this.measureString(label.Value, font, columnWidth);

                            var rect = new XRect(this._currentXPos, this._currentYPos, size.Width, size.Height);

                            if (!string.IsNullOrEmpty(obj.BackColor))
                            {
                                var backbrush = new XSolidBrush(getColorFromHtml(obj.BackColor));
                                this._currentGraphics.DrawRectangle(backbrush, rect);
                            }

                            XTextFormatter textFormatter = new XTextFormatter(this._currentGraphics);
                            switch (label.Alignment)
                            {
                                case System.Drawing.StringAlignment.Near:
                                    textFormatter.Alignment = XParagraphAlignment.Left;
                                    break;
                                case System.Drawing.StringAlignment.Center:
                                    textFormatter.Alignment = XParagraphAlignment.Center;
                                    break;
                                case System.Drawing.StringAlignment.Far:
                                    textFormatter.Alignment = XParagraphAlignment.Right;
                                    break;
                            }
                            textFormatter.DrawString(label.Value ?? "", font, brush, rect, XStringFormats.TopLeft);
                        }
                    }
                    this._currentXPos += columnWidth;
                }


                this._currentYPos += highesttHeight;
                this._currentXPos = this._template.PageOptions.Margins.Left;

            }
        }

        XColor getColorFromHtml(string color)
        {
            return XColor.FromArgb(System.Drawing.ColorTranslator.FromHtml(color).ToArgb());
        }

        XSize measureString(string @string, XFont font, double width)
        {
            var size = this._currentGraphics.MeasureString(@string ?? "", font);
            var height = size.Width / width;
            var floorHeight = Math.Ceiling(height);
            return new XSize(width, (floorHeight * size.Height) + 1);
        }
    }
}
