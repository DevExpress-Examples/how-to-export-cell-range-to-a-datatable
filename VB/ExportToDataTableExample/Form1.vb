Imports Microsoft.VisualBasic
#Region "#usings"
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Export
#End Region ' #usings
Imports System
Imports System.Data
Imports System.Windows.Forms

Namespace ExportToDataTableExample
	Partial Public Class Form1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm
		Public Sub New()
			InitializeComponent()
			spreadsheetControl1.LoadDocument("TopTradingPartners.xlsx")
			ribbonControl1.SelectedPage = exportDataExampleRibbonPage
		End Sub

		Private Sub barButtonItemRangeToDataTable_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItemRangeToDataTable.ItemClick
'			#Region "#SimpleDataExport"
			Dim worksheet As Worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet
			Dim range As Range = worksheet.Selection
			Dim rangeHasHeaders As Boolean = Me.barCheckItemHasHeaders1.Checked
			' Create a data table with column names obtained from the first row in a range if it has headers.
			' Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
			Dim dataTable As DataTable = worksheet.CreateDataTable(range, rangeHasHeaders)
			' Create the exporter that obtains data from the specified range, 
			'skips header row if required and populates the specified data table. 
			Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(range, dataTable, rangeHasHeaders)

			' Perform the export.
			exporter.Export()
'			#End Region ' #SimpleDataExport
			' A custom method that displays the resulting data table.
			ShowResult(dataTable)
		End Sub

'			#Region "#DataExportWithOptions"
		Private Sub barButtonItemUseExporterOptions_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItemUseExporterOptions.ItemClick
			Dim worksheet As Worksheet = spreadsheetControl1.Document.Worksheets(0)
			Dim range As Range = worksheet.Tables(0).Range
			' Create a data table with column names obtained from the first row in a range if it has headers.
			' Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
			Dim dataTable As DataTable = worksheet.CreateDataTable(range, True)
			' Create the exporter that obtains data from the specified range, 
			'skips the header row (if required) and populates the specified data table. 
			Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(range, dataTable, True)
			' Specify exporter options.
			exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = barCheckItemSkipErrors.Checked
			AddHandler exporter.CellValueConversionError, AddressOf exporter_CellValueConversionError

			' Perform the export.
			exporter.Export()
			' A custom method that displays the resulting data table.
			ShowResult(dataTable)
		End Sub

		Private Sub exporter_CellValueConversionError(ByVal sender As Object, ByVal e As CellValueConversionErrorEventArgs)
			MessageBox.Show("Error in cell " & e.Cell.GetReferenceA1())
			e.DataTableValue = Nothing
			e.Action = DataTableExporterAction.Continue
		End Sub
        '			#End Region ' #DataExportWithOptions

'			#Region "#DataExportWithCustomConverter"
		Private Sub barButtonItem1_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonItem1.ItemClick
            Dim worksheet As Worksheet = spreadsheetControl1.Document.Worksheets(0)
            Dim range As Range = worksheet.Tables(0).Range
            ' Create a data table with column names obtained from the first row in a range.
            ' Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            Dim dataTable As DataTable = worksheet.CreateDataTable(range, True)
            ' Change the data type of the "As Of" column to text.
            dataTable.Columns("As Of").DataType = System.Type.GetType("System.String")
            ' Create the exporter that obtains data from the specified range, 
            'skips header row if required and populates the specified data table. 
            Dim exporter As DataTableExporter = worksheet.CreateDataTableExporter(range, dataTable, True)
            ' Specify a custom converter for the "As Of" column.
            Dim toDateStringConverter As New DateTimeToStringConverter()
            exporter.Options.CustomConverters.Add("As Of", toDateStringConverter)
            ' Set the export value for empty cell.
            toDateStringConverter.EmptyCellValue = "N/A"
            ' Specify that empty cells and cells with errors should be processed.
            exporter.Options.ConvertEmptyCells = True
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = False
            ' Perform the export.
            exporter.Export()
            ' A custom method that displays the resulting data table.
            ShowResult(dataTable)
		End Sub

		' A custom converter that converts DateTime values to "Month-Year" text strings.
		Private Class DateTimeToStringConverter
			Implements ICellValueToColumnTypeConverter
			Private privateSkipErrorValues As Boolean
			Public Property SkipErrorValues() As Boolean
				Get
					Return privateSkipErrorValues
				End Get
				Set(ByVal value As Boolean)
					privateSkipErrorValues = value
				End Set
			End Property
			Private privateEmptyCellValue As CellValue
            Public Property EmptyCellValue() As CellValue Implements ICellValueToColumnTypeConverter.EmptyCellValue
                Get
                    Return privateEmptyCellValue
                End Get
                Set(ByVal value As CellValue)
                    privateEmptyCellValue = value
                End Set
            End Property

            Public Function Convert(ByVal readOnlyCell As Cell, ByVal cellValue As CellValue, ByVal dataColumnType As Type, <System.Runtime.InteropServices.Out()> ByRef result As Object) _
                As ConversionResult Implements ICellValueToColumnTypeConverter.Convert
                result = DBNull.Value
                Dim converted As ConversionResult = ConversionResult.Success
                If cellValue.IsEmpty Then
                    result = EmptyCellValue
                    Return converted
                End If
                If cellValue.IsError Then
                    ' You can return an error, subsequently the exporter throws an exception if the CellValueConversionError event is unhandled.
                    'return SkipErrorValues ? ConversionResult.Success : ConversionResult.Error;
                    result = "N/A"
                    Return ConversionResult.Success
                End If
                result = String.Format("{0:MMMM-yyyy}", cellValue.DateTimeValue)
                Return converted
            End Function
		End Class
        '		#End Region ' #DataExportWithCustomConverter


		#Region "Show Result Form"
		Private Function ShowResult(ByVal result As DataTable) As Form
			Dim newForm As New Form()
			newForm.Width = 600
			newForm.Height = 300

			Dim grid As New DevExpress.XtraGrid.GridControl()
			grid.Dock = DockStyle.Fill
			grid.DataSource = result

			newForm.Controls.Add(grid)
			grid.ForceInitialize()
			CType(grid.FocusedView, DevExpress.XtraGrid.Views.Grid.GridView).OptionsView.ShowGroupPanel = False

			newForm.ShowDialog(Me)
			Return newForm
		End Function
		#End Region
	End Class
End Namespace
