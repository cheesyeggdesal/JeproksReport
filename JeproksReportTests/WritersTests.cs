using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using JeproksReport;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using JeproksReport.ExcelWriter;
using JeproksReport.iTextSharpWriter;

namespace JeproksReportTests
{
    [TestClass]
    public class WritersTests
    {
        ReportTemplate generateSimpleTemplate()
        {
            var template = new ReportTemplate();
            template.InitializeColumns(75, 300, 50, 100);

            var header = new ReportSection();
            template.Header = header;
            var currentrow = header.AddRow();
            currentrow.AddLabel("Product List", new Font("Arial", 12, FontStyle.Bold));
            
            currentrow = header.AddRow();
            currentrow.AddLabel("Date Printed :");
            currentrow.AddLabel(string.Format("{0:MMM dd, yyyy}", DateTime.Now));

            header.AddRow().AddLabel(" "); //blank row
            
            currentrow = header.AddRow();
            var headerStyle = new TextBaseStyle()
            {
                Color = "White",
                BackColor = "#0063B1",
                Bold = true
            };

            currentrow.AddLabel("Product Id", headerStyle);
            currentrow.AddLabel("Product Name", headerStyle);
            currentrow.AddLabel("Quantity", headerStyle.Align(StringAlignment.Far));
            currentrow.AddLabel("Price", headerStyle.Align(StringAlignment.Far));

            var details = template.AddSection();
            details.DataSource = "products";

            currentrow = details.AddRow();
            var numberStyle = new TextBaseStyle() { Alignment = StringAlignment.Far };
            currentrow.AddDataField("ProductId");
            currentrow.AddDataField("ProductName");
            currentrow.AddDataField("Quantity", numberStyle, "N2");
            currentrow.AddDataField("Price", numberStyle, "N2");
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
        public void should_write_simple_itextsharp_report()
        {
            var reportfilename = ".\\output_simplepdfsharpreport.pdf";

            var datasource = new Dictionary<string, IEnumerable<object>>();
            datasource.Add("products", this.GetProductList());

            var template = this.generateSimpleTemplate();
            var writer = new iTextSharpWriter();
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
