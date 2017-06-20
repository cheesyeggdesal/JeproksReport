using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport.ExcelWriter
{
    public class ExcelWriter : IReportWriter
    {
        Dictionary<string, IEnumerable<object>> _datasource;
        IXLWorksheet _currentWorksheet;
        int _currentRowIndex;

        public void Write(ReportTemplate reportTemplate, Stream stream, Dictionary<string, IEnumerable<object>> datasouce = null)
        {
            this._datasource = datasouce;
            
            var workbook = new XLWorkbook();
            workbook.PageOptions.PageOrientation =
                reportTemplate.PageOptions.PageOrientation == PageOrientation.Portrait ?
                XLPageOrientation.Portrait : XLPageOrientation.Landscape;

            workbook.PageOptions.Margins.Bottom = reportTemplate.PageOptions.Margins.Bottom / 72;
            workbook.PageOptions.Margins.Top = reportTemplate.PageOptions.Margins.Top / 72;
            workbook.PageOptions.Margins.Left = reportTemplate.PageOptions.Margins.Left / 72;
            workbook.PageOptions.Margins.Right = reportTemplate.PageOptions.Margins.Right / 72;

            this._currentWorksheet = workbook.Worksheets.Add("Report");
            this._currentRowIndex = 1;
            this.InitializeColumns(reportTemplate.Columns);
            foreach (var section in reportTemplate.Sections)
            {
                this.RenderSection(section);
            }

            workbook.SaveAs(stream);
        }

        void InitializeColumns(IEnumerable<ReportColumn> reportColumns)
        {
            for (int i = 0; i < reportColumns.Count(); i++)
            {
                var width = getExcelPoint(reportColumns.ElementAt(i).Width);
                this._currentWorksheet.Column(i + 1).Width = width;
            }
        }

        double getExcelPoint(double p)
        {
            return (p - 12) / 7d + 1;
        }

        void RenderSection(ReportSection section)
        {
            if (!string.IsNullOrEmpty(section.DataSource))
            {
                if (this._datasource.ContainsKey(section.DataSource)
                        && this._datasource[section.DataSource] != null)
                {
                    var lastValues = new Dictionary<string, string>();
                    foreach (var datarow in this._datasource[section.DataSource])
                    {
                        RenderRows(section.Rows.Select(row => row.FromData(datarow, lastValues)));
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
                int currentColumnIndex = 1;
                foreach (var obj in row.Objects)
                {
                    var cell = this._currentWorksheet.Row(this._currentRowIndex).Cell(currentColumnIndex);
                    currentColumnIndex++;
                    renderObject(obj, cell);
                }
                this._currentRowIndex++;
            }
        }

        void renderObject(ReportObject obj, IXLCell cell)
        {

            if (obj is TextBase)
            {
                var textbase = obj as TextBase;
                renderReportObject(textbase.Style, cell);
                cell.Style.Font.FontName = textbase.Style.FontFamily;
                cell.Style.Font.FontSize = textbase.Style.FontSize;
                cell.Style.Font.FontColor = XLColor.FromName(textbase.Style.Color);
                cell.Style.Font.Bold = (obj as TextBase).Style.Bold;
                cell.Style.Font.Italic = (obj as TextBase).Style.Italic;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                switch ((obj as TextBase).Style.Alignment)
                {
                    case System.Drawing.StringAlignment.Center:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        break;
                    case System.Drawing.StringAlignment.Far:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        break;
                    case System.Drawing.StringAlignment.Near:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        break;
                    default:
                        break;
                }

                var value = "";

                if (obj is Label)
                {
                    value = ((Label)obj).Value;
                }

                cell.Value = value;
            }
        }

        void renderReportObject(Style style, IXLCell cell)
        {
            if (style.Borders.Top != null && style.Borders.Top.Visible)
            {
                cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.TopBorderColor = XLColor.FromName(style.Borders.Top.Color);
            }
            if (style.Borders.Right != null && style.Borders.Right.Visible)
            {
                cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.RightBorderColor = XLColor.FromName(style.Borders.Right.Color);
            }
            if (style.Borders.Bottom != null && style.Borders.Bottom.Visible)
            {
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.BottomBorderColor = XLColor.FromName(style.Borders.Bottom.Color);
            }
            if (style.Borders.Left != null && style.Borders.Left.Visible)
            {
                cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.LeftBorderColor = XLColor.FromName(style.Borders.Left.Color);
            }

            if (!string.IsNullOrEmpty(style.BackColor))
            {
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml(style.BackColor);
            }
        }
    }
}
