using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using JeproksReport;

namespace JeproksReportTests
{
    [TestClass]
    public class WritersTests
    {
        ReportTemplate generateSimpleTemplate()
        {
            var template = new ReportTemplate();

            template.Columns.Add(new ReportColumn(100));
            template.Columns.Add(new ReportColumn(200));

            var header = template.AddSection();
            var currentrow = header.AddRow();
            currentrow.Objects.Add(new Label()
            {

            });
            return template;
        }

        [TestMethod]
        public void should_write_simple_pdfsharp_report()
        {
            var template = this.generateSimpleTemplate();
        }
    }
}
