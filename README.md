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
void WritePDFReport()
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