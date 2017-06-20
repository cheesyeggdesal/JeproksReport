# Basic Usage
```csharp
//generate report template
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

//create pdf report using pdfsharpwriter
void WriteExcelReport()
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

# Screenshots
## PDFWriter
![image](https://cloud.githubusercontent.com/assets/3628837/25771398/7d94e3c6-3282-11e7-852c-2c24966cbbe9.png)
## ExcelWriter
![image](https://cloud.githubusercontent.com/assets/3628837/25771405/a07569ce-3282-11e7-9e00-2ff605b8e3b1.png)
