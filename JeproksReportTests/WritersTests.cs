using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using JeproksReport;
using JeproksReport.PDFSharp;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using JeproksReport.ExcelWriter;

namespace JeproksReportTests
{
    [TestClass]
    public class WritersTests
    {
        ReportTemplate generateSimpleTemplate()
        {
            var template = new ReportTemplate();

            template.Columns.Add(new ReportColumn(75));
            template.Columns.Add(new ReportColumn(300));
            template.Columns.Add(new ReportColumn(50));
            template.Columns.Add(new ReportColumn(100));

            var header = template.AddSection();
            var currentrow = header.AddRow();
            currentrow.Objects.Add(new Label()
            {
                Bold = true,
                Value = "Product List",
                FontSize = 12
            });
            currentrow = header.AddRow();
            currentrow.Objects.Add(new Label()
            {
                Value = string.Format("Date Printed : ")
            });
            currentrow.Objects.Add(new Label()
            {
                Value = string.Format("{0:MMM dd, yyyy}", DateTime.Now)
            });
            header.AddRow().Objects.Add(new Label() { Value = " " }); //blank row
            currentrow = header.AddRow();
            currentrow.Objects.Add(new Label()
            {
                Bold = true,
                BackColor = "#0063B1",
                Color = "White",
                Value = "Product Id"
            });
            currentrow.Objects.Add(new Label()
            {
                Bold = true,
                BackColor = "#0063B1",
                Color = "White",
                Value = "Product Name"
            });
            currentrow.Objects.Add(new Label()
            {
                Alignment = System.Drawing.StringAlignment.Far,
                Bold = true,
                BackColor = "#0063B1",
                Color = "White",
                Value = "Quantity "
            });
            currentrow.Objects.Add(new Label()
            {
                Alignment = System.Drawing.StringAlignment.Far,
                Bold = true,
                BackColor = "#0063B1",
                Color = "White",
                Value = "Price "
            });

            var details = template.AddSection();
            details.DataSource = "products";

            currentrow = details.AddRow();
            currentrow.Objects.Add(new DataField()
            {
                Field = "ProductId"
            });
            currentrow.Objects.Add(new DataField()
            {
                Field = "ProductName"
            });
            currentrow.Objects.Add(new DataField()
            {
                Alignment = StringAlignment.Far,
                Field = "Quantity"
            });
            currentrow.Objects.Add(new DataField()
            {
                Alignment = StringAlignment.Far,
                Field = "Price",
                Format = "N2"
            });
            return template;
        }

        IEnumerable<object> GetProductList()
        {
            var rnd = new Random();
            return Enumerable.Range(1, 100).Select(p => new
            {
                ProductId = p,
                ProductName = "Apple Pen " + p,
                Quantity = rnd.Next(50),
                Price = rnd.NextDouble() * 123
            });
        }

        [TestMethod]
        public void should_write_simple_pdfsharp_report()
        {
            var reportfilename = ".\\output_simplepdfsharpreport.pdf";

            var datasource = new Dictionary<string, IEnumerable<object>>();
            datasource.Add("products", this.GetProductList());

            var template = this.generateSimpleTemplate();
            var writer = new PdfSharpWriter();
            using (var file = File.Create(reportfilename))
            {
                writer.Write(template, file, datasource);
            }
        }

        [TestMethod]
        public void should_write_simple_excel_report()
        {
            var reportfilename = ".\\output_simpleexcelreport.xlsx";

            var datasource = new Dictionary<string, IEnumerable<object>>();
            datasource.Add("products", this.GetProductList());

            var template = this.generateSimpleTemplate();
            var writer = new ExcelWriter();
            using (var file = File.Create(reportfilename))
            {
                writer.Write(template, file, datasource);
            }
        }
    }
}
