# How to export cell range to a DataTable


<p>This example illustrates how you can export worksheet cell range to a System.Data.DataTable object.</p>
<p>The following steps are required:</p>
<p>1) Add a reference to the <strong>DevExpress.Docs.vX.Y.dll</strong> assembly to your project containing the SpreadsheetControl. Note that distribution of this assembly requires <a href="https://www.devexpress.com/Products/NET/Document-Server/pricing.xml">a license to the DevExpress Document Server or the DevExpress Universal Subscription</a>.<br /> 2) Create the <strong>DevExpress.Spreadsheet.Export.DataTableExporter</strong> instance using the <strong>DevExpress.Spreadsheet.Worksheet.CreateDataTableExporter</strong> method.<br /> 3) Call the <strong>Export </strong>method of the DataTableExporter.</p>
<p>You can create an empty DataTable by using the <strong>CreateDataTable </strong>method of the DataTableExporter. The column names are obtained from headings of the cell range, and the column data types are extracted from cell data types of the first row of a range.<br /> The DataTableExporter contains various options that enables you to specify how cell data are processed before storing them in a DataTable.</p>

<br/>


