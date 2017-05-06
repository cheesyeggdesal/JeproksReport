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
            currentrow.AddLabel("Product List", new Font("Arial", 12, FontStyle.Bold));
            
            currentrow = header.AddRow();
            currentrow.AddLabel("Date Printed :");
            currentrow.AddLabel(string.Format("{0:MMM dd, yyyy}", DateTime.Now));

            header.AddRow().AddLabel(" "); //blank row

            Font tableheaderfont = new Font("Arial", 9, FontStyle.Bold);
            currentrow = header.AddRow();
            currentrow.AddLabel("Product Id", tableheaderfont, "White", "#0063B1");
            currentrow.AddLabel("Product Name", tableheaderfont, "White", "#0063B1");
            currentrow.AddLabel("Quantity", tableheaderfont, "White", "#0063B1", StringAlignment.Far);
            currentrow.AddLabel("Price", tableheaderfont, "White", "#0063B1", StringAlignment.Far);

            var details = template.AddSection();
            details.DataSource = "products";

            currentrow = details.AddRow();
            currentrow.AddDataField("ProductId");
            currentrow.AddDataField("ProductName");
            currentrow.AddDataField("Quantity", format: "N2", alignment: StringAlignment.Far);
            currentrow.AddDataField("Price", format: "N2", alignment: StringAlignment.Far);
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
