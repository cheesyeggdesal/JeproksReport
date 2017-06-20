using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport.iTextSharpWriter
{
    public class iTextSharpWriter : IReportWriter
    {
        Document _document;
        ReportTemplate _template;
        Dictionary<string, IEnumerable<object>> _datasource;

        public global::iTextSharp.text.Rectangle GetPageSize()
        {
            switch (this._template.PageOptions.PageSize)
            {
                case PageSize.Letter:
                    return global::iTextSharp.text.PageSize.LETTER;
                case PageSize.A4:
                default:
                    return global::iTextSharp.text.PageSize.A4;
            }
        }

        public void Write(ReportTemplate reportTemplate, System.IO.Stream stream, Dictionary<string, IEnumerable<object>> datasource = null)
        {
            this._template = reportTemplate;
            this._datasource = datasource;

            var pagesize = this.GetPageSize();

            if (this._template.PageOptions.PageOrientation == PageOrientation.Landscape)
            {
                pagesize = pagesize.Rotate();
            }
            this._document = new Document(pagesize,
                (float)this._template.PageOptions.Margins.Left,
                (float)this._template.PageOptions.Margins.Right,
                (float)this._template.PageOptions.Margins.Top + this.getHeaderHeight(pagesize),
                (float)this._template.PageOptions.Margins.Bottom);
            PdfWriter writer = PdfWriter.GetInstance(this._document, stream);
            writer.PageEvent = new PageXOfY(this._template);
            writer.CloseStream = false;
            this._document.Open();
            this._document.Add(new Chunk());
            RenderSections();
            this._document.Close();
        }

        void RenderSections()
        {
            SectionRenderer sectionrenderer = new SectionRenderer(this._template);
            var table = new PdfPTable(this._template.Columns.Count);

            table.WidthPercentage = 100;
            table.SetWidths(this._template.Columns.Select(x => (float)x.Width).ToArray());

            for (int x = 0; x < this._template.Sections.Count; x++)
            {
                var section = this._template.Sections[x];

                if (!string.IsNullOrEmpty(section.DataSource))
                {
                    if (this._datasource.ContainsKey(section.DataSource)
                        && this._datasource[section.DataSource] != null)
                    {
                        var lastValues = new Dictionary<string, string>();

                        foreach (var datarow in this._datasource[section.DataSource])
                        {
                            sectionrenderer.Render(section.Rows.Select(row => 
                            row.FromData(datarow, lastValues)).ToList(), table);
                        }
                    }
                }
                else
                {
                    sectionrenderer.Render(section.Rows, table);
                }
            }

            this._document.Add(table);
        }

        float getHeaderHeight(global::iTextSharp.text.Rectangle pageSize)
        {
            if (this._template.Header == null) return 0;
            SectionRenderer renderer = new SectionRenderer(this._template);
            PdfPTable tblHeader = new PdfPTable(this._template.Columns.Count);
            tblHeader.WidthPercentage = 100;
            tblHeader.SetWidths(this._template.Columns.Select(x => (float)x.Width).ToArray());
            tblHeader.TotalWidth = pageSize.Width - ((float)(this._template.PageOptions.Margins.Left - this._template.PageOptions.Margins.Left));
            renderer.Render(_template.Header.Rows, tblHeader);
            return tblHeader.CalculateHeights();
        }
    }

    public class SectionRenderer
    {
        ReportTemplate _template;

        public SectionRenderer(ReportTemplate template)
        {
            this._template = template;
        }

        public void Render(List<ReportRow> rows, PdfPTable table)
        {
            if (rows == null) return;
            int totalRowSpan = 0;
            for (int y = 0; y < rows.Count(); y++)
            {
                int totalColSpan = 0;
                for (int z = 0; z < this._template.Columns.Count; z++)
                {
                    if (z >= rows[y].Objects.Count)
                    {
                        if (totalColSpan > 0)
                        {
                            totalColSpan--;
                            continue;
                        }
                        if (totalRowSpan > 0)
                        {
                            totalRowSpan--;
                            continue;
                        }
                        var dummyCell = new PdfPCell();
                        dummyCell.BorderWidth = 0;
                        table.AddCell(dummyCell);
                        continue;
                    }

                    var repObj = rows[y].Objects[z] as ReportObject;
                    PdfPCell pdfCell = null;
                    Style style = null;

                    if (repObj is TextBase)
                    {

                        var txtBase = repObj as TextBase;
                        style = txtBase.Style;
                        IElement paragraph = null;
                        if (repObj is Label)
                        {
                            paragraph = new Paragraph((repObj as Label).Value);
                        }
                        else
                        {
                            paragraph = new Paragraph();

                        }

                        if (paragraph is Paragraph)
                        {
                            applyTextBaseStyle(paragraph as Paragraph, txtBase);
                            pdfCell = new PdfPCell(paragraph as Paragraph);
                        }

                        switch (txtBase.Style.Alignment)
                        {
                            case StringAlignment.Center:
                                pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                                break;
                            case StringAlignment.Far:
                                pdfCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                break;
                            case StringAlignment.Near:
                                pdfCell.HorizontalAlignment = Element.ALIGN_LEFT;
                                break;
                        }

                        switch (txtBase.Style.VerticalAlignment)
                        {
                            case StringAlignment.Center:
                                pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                break;
                            case StringAlignment.Far:
                                pdfCell.VerticalAlignment = Element.ALIGN_BOTTOM;
                                break;
                            case StringAlignment.Near:
                                pdfCell.VerticalAlignment = Element.ALIGN_TOP;
                                break;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(style.BackColor))
                    {
                        pdfCell.BackgroundColor = new BaseColor(System.Drawing.ColorTranslator.FromHtml(style.BackColor));
                    }

                    pdfCell.BorderWidthTop = (float)style.Borders.Top.Width;
                    pdfCell.BorderWidthRight = (float)style.Borders.Right.Width;
                    pdfCell.BorderWidthBottom = (float)style.Borders.Bottom.Width;
                    pdfCell.BorderWidthLeft = (float)style.Borders.Left.Width;

                    pdfCell.Colspan = repObj.ColSpan;
                    totalColSpan += repObj.ColSpan - 1;
                    pdfCell.Rowspan = repObj.RowSpan;
                    totalRowSpan += repObj.RowSpan - 1;
                    table.AddCell(pdfCell);
                }
            }
        }

        void applyTextBaseStyle(Paragraph paragraph, TextBase txtBase)
        {
            if (!string.IsNullOrEmpty(txtBase.Style.Color))
            {
                paragraph.Font.Color = new BaseColor(System.Drawing.ColorTranslator.FromHtml(txtBase.Style.Color));
            }
            paragraph.Font.Size = (float)txtBase.Style.FontSize;
            paragraph.Font.SetFamily(txtBase.Style.FontFamily);
            var style = "normal";
            if (txtBase.Style.Bold) style = "bold";
            paragraph.Font.SetStyle(style);
        }
    }

    public class PageXOfY : PdfPageEventHelper
    {
        PdfTemplate _total;
        ReportTemplate _template;
        public PageXOfY(ReportTemplate template)
        {
            this._template = template;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            this._total = writer.DirectContent.CreateTemplate(30, 16);
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            RenderFooter(writer, document);
            RenderHeader(writer, document);
        }

        void RenderHeader(PdfWriter writer, Document document)
        {
            if (this._template.Header == null) return;
            PdfContentByte cb = writer.DirectContent;
            SectionRenderer renderer = new SectionRenderer(this._template);
            PdfPTable tblHeader = new PdfPTable(this._template.Columns.Count);
            tblHeader.WidthPercentage = 100;
            tblHeader.SetWidths(this._template.Columns.Select(x => (float)x.Width).ToArray());
            tblHeader.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
            renderer.Render(_template.Header.Rows, tblHeader);
            tblHeader.WriteSelectedRows(0, -1,
                document.LeftMargin, document.PageSize.Height - ((float)this._template.PageOptions.Margins.Top),
                cb);
        }

        void RenderFooter(PdfWriter writer, Document document)
        {
            PdfContentByte cb = writer.DirectContent;

            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            table.SetWidths(new int[] { 100, 5 });
            table.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
            var paraPageXOf = new Paragraph(string.Format("Page {0} of", writer.PageNumber));
            paraPageXOf.Font.Size = 8;
            paraPageXOf.Font.SetFamily("Arial");
            var cellPageXOf = new PdfPCell(paraPageXOf);
            cellPageXOf.BorderWidth = 0;
            cellPageXOf.BorderWidthTop = .5f;
            cellPageXOf.HorizontalAlignment = Element.ALIGN_RIGHT;
            table.AddCell(cellPageXOf);
            var cellTotalPage = new PdfPCell(global::iTextSharp.text.Image.GetInstance(_total));
            cellTotalPage.HorizontalAlignment = Element.ALIGN_LEFT;
            cellTotalPage.BorderWidth = 0;
            cellTotalPage.BorderWidthTop = .5f;
            table.AddCell(cellTotalPage);
            table.WriteSelectedRows(0, -1, document.LeftMargin, document.BottomMargin, cb);
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            var phrase = new Phrase(writer.PageNumber.ToString());
            phrase.Font.SetFamily("Arial");
            phrase.Font.Size = 8;
            ColumnText.ShowTextAligned(_total, Element.ALIGN_LEFT,
                    phrase,
                    0, 6, 0);
        }
    }
}
