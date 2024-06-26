<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613578/15.1.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E4997)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/ExportToDataTableExample/Form1.cs) (VB: [Form1.vb](./VB/ExportToDataTableExample/Form1.vb))
<!-- default file list end -->
# How to export a cell range to a DataTable

This example illustrates how you can export a cell range to a System.Data.DataTable object.

The following steps are required:

1) Add a reference to the **DevExpress.Docs.dll** assembly to your Spreadsheet project. The distribution of this assembly requires <a href="https://www.devexpress.com/products/net/office-file-api/">a license to the DevExpress Office File API or DevExpress Universal Subscription</a>.

2) Use the **DevExpress.Spreadsheet.Worksheet.CreateDataTableExporter** method to create a **DevExpress.Spreadsheet.Export.DataTableExporter** instance.

3) Call the DataTableExporter's **Export** method.

You can use the **Worksheet.CreateDataTable** method to create an empty DataTable from a cell range. This method obtains column names from the range headings, and determines the column data types based on the first row of the specified range.
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-export-cell-range-to-a-datatable&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-export-cell-range-to-a-datatable&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
