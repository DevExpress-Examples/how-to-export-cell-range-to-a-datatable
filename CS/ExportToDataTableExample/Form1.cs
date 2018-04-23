#region #usings
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
#endregion #usings
using System;
using System.Data;
using System.Windows.Forms;

namespace ExportToDataTableExample
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            spreadsheetControl1.LoadDocument("TopTradingPartners.xlsx");
            ribbonControl1.SelectedPage = exportDataExampleRibbonPage;
        }

        private void barButtonItemRangeToDataTable_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #SimpleDataExport
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet;
            Range range = worksheet.Selection;
            bool rangeHasHeaders = this.barCheckItemHasHeaders1.Checked;
            // Create a data table with column names obtained from the first row in a range if it has headers.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, rangeHasHeaders);
            // Create the exporter that obtains data from the specified range, 
            // skips the header row (if required) and populates the specified data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, rangeHasHeaders);

            // Perform the export.
            exporter.Export();
            #endregion #SimpleDataExport
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

            #region #DataExportWithOptions
        private void barButtonItemUseExporterOptions_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
            Range range = worksheet.Tables[0].Range;
            // Create a data table with column names obtained from the first row in a range if it has headers.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, true);
            // Create the exporter that obtains data from the specified range, 
            //skips header row if required and populates the specified data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            // Specify exporter options.
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = barCheckItemSkipErrors.Checked;
            exporter.CellValueConversionError += exporter_CellValueConversionError;

            // Perform the export.
            exporter.Export();
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            MessageBox.Show("Error in cell " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }
            #endregion #DataExportWithOptions

            #region #DataExportWithCustomConverter
        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
            Range range = worksheet.Tables[0].Range;
            // Create a data table with column names obtained from the first row in a range.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, true);
            // Change the data type of the "As Of" column to text.
            dataTable.Columns["As Of"].DataType = System.Type.GetType("System.String");
            // Create the exporter that obtains data from the specified range, 
            //skips header row if required and populates the specified data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            // Specify a custom converter for the "As Of" column.
            DateTimeToStringConverter toDateStringConverter = new DateTimeToStringConverter();
            exporter.Options.CustomConverters.Add("As Of", toDateStringConverter);
            // Set the export value for empty cell.
            toDateStringConverter.EmptyCellValue = "N/A";
            // Specify that empty cells and cells with errors should be processed.
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;
            // Perform the export.
            exporter.Export();
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        // A custom converter that converts DateTime values to "Month-Year" text strings.
        class DateTimeToStringConverter : ICellValueToColumnTypeConverter
        {
            public bool SkipErrorValues { get; set; }
            public CellValue EmptyCellValue { get; set; }

            public ConversionResult Convert(Cell readOnlyCell, CellValue cellValue, Type dataColumnType, out object result)
            {
                result = DBNull.Value; 
                ConversionResult converted = ConversionResult.Success;
                if (cellValue.IsEmpty) {
                    result = EmptyCellValue;
                    return converted;
                }
                if (cellValue.IsError) {
                    // You can return an error, subsequently the exporter throws an exception if the CellValueConversionError event is unhandled.
                    //return SkipErrorValues ? ConversionResult.Success : ConversionResult.Error;
                    result = "N/A";
                    return ConversionResult.Success;
                }
                result =  String.Format("{0:MMMM-yyyy}",cellValue.DateTimeValue);
                return converted;
            }
        }
        #endregion #DataExportWithCustomConverter


        #region Show Result Form
        Form ShowResult(DataTable result)
        {
            Form newForm = new Form();
            newForm.Width = 600;
            newForm.Height = 300;

            DevExpress.XtraGrid.GridControl grid = new DevExpress.XtraGrid.GridControl();
            grid.Dock = DockStyle.Fill;
            grid.DataSource = result;

            newForm.Controls.Add(grid);
            grid.ForceInitialize();
            ((DevExpress.XtraGrid.Views.Grid.GridView)grid.FocusedView).OptionsView.ShowGroupPanel = false;

            newForm.ShowDialog(this);
            return newForm;
        }
        #endregion





    }
}
