# Basic Usage
```csharp
//generate report template
ReportTemplate productListTemplate()
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

//create pdf report using pdfsharpwriter
void WritePDFReport()
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

//create excel report using excelwriter
void WriteExcelReport()
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
```