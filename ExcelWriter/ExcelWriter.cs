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

            using (var workbook = new XLWorkbook())
            {
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
            renderReportObject(obj, cell);
            if (obj is TextBase)
            {
                var textbase = obj as TextBase;
                cell.Style.Font.FontName = textbase.FontFamily;
                cell.Style.Font.FontSize = textbase.FontSize;
                cell.Style.Font.FontColor = XLColor.FromName(textbase.Color);
                cell.Style.Font.Bold = (obj as TextBase).Bold;
                cell.Style.Font.Italic = (obj as TextBase).Italic;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                switch ((obj as TextBase).Alignment)
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

        void renderReportObject(ReportObject obj, IXLCell cell)
        {
            if (obj.Borders.Top != null && obj.Borders.Top.Visible)
            {
                cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.TopBorderColor = XLColor.FromName(obj.Borders.Top.Color);
            }
            if (obj.Borders.Right != null && obj.Borders.Right.Visible)
            {
                cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.RightBorderColor = XLColor.FromName(obj.Borders.Right.Color);
            }
            if (obj.Borders.Bottom != null && obj.Borders.Bottom.Visible)
            {
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.BottomBorderColor = XLColor.FromName(obj.Borders.Bottom.Color);
            }
            if (obj.Borders.Left != null && obj.Borders.Left.Visible)
            {
                cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.LeftBorderColor = XLColor.FromName(obj.Borders.Left.Color);
            }

            if (!string.IsNullOrEmpty(obj.BackColor))
            {
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml(obj.BackColor);
            }
        }
    }
}
