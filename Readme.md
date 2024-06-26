<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613578/19.2.2%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E4997)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# How to Export a Cell Range from a Spreadsheet Document to a DataTable

This example illustrates how you can export a cell range to a [System.Data.DataTable](https://learn.microsoft.com/en-us/dotnet/api/system.data.datatable) object.

Use the [Worksheet.CreateDataTableExporter](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.WorksheetExtensions.CreateDataTableExporter(DevExpress.Spreadsheet.Worksheet-DevExpress.Spreadsheet.CellRange-System.Data.DataTable-System.Boolean)) method to create a [DataTableExporter](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Export.DataTableExporter) instance and call the DataTableExporter's **Export** method.

You can use the [Worksheet.CreateDataTable](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.WorksheetExtensions.CreateDataTable.overloads) method to create an empty **DataTable** from a cell range. This method obtains column names from the range headings, and determines the column data types based on the first row of the specified range.

**Note:** Add a reference to the **DevExpress.Docs** assembly to your Spreadsheet project. The distribution of this assembly requires [a license to the DevExpress Office File API or DevExpress Universal Subscription](https://www.devexpress.com/products/net/office-file-api/).

# Files to Look At

* [Form1.cs](./CS/ExportToDataTableExample/Form1.cs) (VB: [Form1.vb](./VB/ExportToDataTableExample/Form1.vb))

# Documentation

* [Import and Export Spreadsheet Data](https://docs.devexpress.com/WindowsForms/16457/controls-and-libraries/spreadsheet/examples/import-and-export-data)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-export-cell-range-to-a-datatable&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-export-cell-range-to-a-datatable&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
