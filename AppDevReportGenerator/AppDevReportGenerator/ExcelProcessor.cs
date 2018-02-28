using AppDevReportGenerator;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace AppDevReportGenerator
{
    public class ProgressUpdatedArgs:EventArgs
    {
        private string message;
        public string Message
        {
            get { return message; }
            set { message = value; }
        }

        public ProgressUpdatedArgs(string Message)
        {
            this.Message = Message;
        }
    }

    public delegate void ProgressUpdatedHandler(object sender, ProgressUpdatedArgs e);

    public class ExcelProcessor
    {
        public event ProgressUpdatedHandler ProgressUpdated;
        public event EventHandler ProcessingComplete;

        protected virtual void OnProgressUpdated(string message)
        {
            if(ProgressUpdated != null)
            {
                ProgressUpdated(this, new ProgressUpdatedArgs(message));
            }
        }
        protected virtual void OnProcessingComplete(EventArgs e)
        {
            if (ProcessingComplete != null)
            {
                ProcessingComplete(this, e);
            }
        }

        public async void ExcelFromReport(Report ActiveReport)
        {
            #region Prepare Excel
            Global.ExcelApplication = new Excel.Application() { Visible = false };
            Global.ExcelApplication.UserControl = false;
            Global.ExcelApplication.DisplayAlerts = false;
            Excel.Workbook book = Global.ExcelApplication.Workbooks.Open(ActiveReport.GetSourceFileName());
            Excel.Worksheet sheet = book.Sheets[1];
            Excel.Range range = sheet.UsedRange;
            #endregion

            #region Load Report Rows Collection
            int dividerrowcount = 0;
            ActiveReport.Rows = new List<Row>();
            await Task.Run(() =>  // Run async so the UI can be updated while processing
            {
                try
                {
                    #region Add Rows to Rows Collection
                    int index = 1;
                    int actualrows = range.Rows.Count - (ActiveReport.FirstRowIndex - 1);

                    string lastoutofcycle = "No";
                    for (int i = ActiveReport.FirstRowIndex; i <= range.Rows.Count; i++)
                    {
                        Row row = new Row { Cells = new List<Cell>() };
                        foreach (ReportField field in ActiveReport.Fields)
                        {
                            string textvalue = ActiveReport.GetExcelText(range.Cells[i, field.SourceIndex].Value2, field);
                            row.Cells.Add(new Cell { ColumnNumber = field.ExportIndex, Value = textvalue, ColumnWidth = field.ColumnWidth, ExportName = field.ExportName });
                        }
                        if (!ActiveReport.IsRowEmpty(range, i))  // sometimes there is a blank line at the end of the excel file, this accounts for it
                        {
                            #region Out Of Cycle Divider
                            var divider = row.Cells.Where(x => x.ExportName == "OutOfCycle").FirstOrDefault();
                            string outofcycle = string.Empty;
                            if (divider != null)
                            {
                                outofcycle = divider.Value;
                            }
                            if (outofcycle == "Yes" && lastoutofcycle == "No")  // Identify the first row that is out of cycle
                            {
                                Row dividerrow = new Row { Cells = new List<Cell>() };
                                foreach (ReportField field in ActiveReport.Fields)  // Add empty cells for each field so row is same length as others (just makes things easier if all same length)
                                {
                                    dividerrow.Cells.Add(new Cell { ColumnNumber = field.ExportIndex, Value = string.Empty, ColumnWidth = field.ColumnWidth });
                                }
                                dividerrow.IsDivider = true;
                                ActiveReport.Rows.Add(dividerrow);
                                dividerrowcount = 1;
                                OnProgressUpdated($"Found Divider Row at Index: {index}");
                            }
                            lastoutofcycle = outofcycle;
                            #endregion
                            ActiveReport.Rows.Add(row);
                        }

                        OnProgressUpdated($"Added Row: {index} of {actualrows}");
                        index++;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error loading source file: {ex.StackTrace}");
                    MessageBox.Show($"Error loading source file: {ex.Message}", "Error Loading File");
                }
                finally
                {
                    book.Close(true, Missing.Value, Missing.Value);
                    Global.ReleaseExcelProcesses();
                }
            });
            #endregion

            #region Create Excel Report

            #region Prepare Excel
            Global.ExcelApplication = new Excel.Application() { Visible = false };
            Global.ExcelApplication.UserControl = false;
            Global.ExcelApplication.DisplayAlerts = false;
            book = Global.ExcelApplication.Workbooks.Add(Missing.Value);
            sheet = book.Worksheets.get_Item(1);
            range = sheet.UsedRange;
            #endregion

            await Task.Run(() =>
            {
                try
                {
                    ColorConverter colorconverter = new ColorConverter();

                    #region Add Report Title Row to Excel
                    sheet.Cells[1, 1] = $"Report: {ActiveReport.Name} - Prepared {DateTime.Now.ToShortDateString()}";
                    int columncount = ActiveReport.Fields.Where(x => x.ExportIndex >= 0).Count();
                    sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, columncount]].Merge();
                    sheet.Cells[1, 1].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    #endregion

                    #region Add Column Headers Rows to Excel
                    foreach (ReportField field in ActiveReport.Fields)
                    {
                        if (field.ExportIndex < 0) { continue; }
                        sheet.Cells[2, field.ExportIndex] = field.Name;
                    }
                    Excel.Range headers = sheet.Range[
                        sheet.Cells[2, 1],
                        sheet.Cells[2, ActiveReport.Fields.Where(x => x.ExportIndex > 0).Count()]
                    ];
                    headers.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)colorconverter.ConvertFromString(ActiveReport.HeaderBackground));
                    #endregion

                    #region Add Rows to Excel
                    int rownumber = ActiveReport.FirstRowIndex;  // This variable is used to get the row content
                    int rowcount = 1;  // This variable is used for presentation of current row (not based on the active report and accounts for the divider row)
                    foreach (Row row in ActiveReport.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            if (cell.ColumnNumber < 1) { continue; }
                            sheet.Cells[rownumber, cell.ColumnNumber] = cell.Value;
                            sheet.Columns[cell.ColumnNumber].ColumnWidth = cell.ColumnWidth;
                            sheet.Cells[rownumber, cell.ColumnNumber].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            #region Add Divider Row to Excel
                            if (row.IsDivider)
                            {
                                Excel.Range divider = sheet.Range[
                                    sheet.Cells[rownumber, 1],
                                    sheet.Cells[rownumber, ActiveReport.Fields.Where(x => x.ExportIndex > 0).Count()]
                                ];
                                divider.Interior.Color = ColorTranslator.ToOle((System.Drawing.Color)colorconverter.ConvertFromString(ActiveReport.DividerBackground));
                            }
                            #endregion
                        }
                        if (!row.IsDivider)
                        {
                            OnProgressUpdated($"Exported Row {rowcount} of {ActiveReport.Rows.Count - dividerrowcount}");
                            rowcount++;
                        }
                        rownumber++;
                    }
                    #endregion

                    #region Apply Sheet Styles
                    range = sheet.UsedRange;
                    range.WrapText = true;
                    range.Rows.AutoFit();
                    range.Columns.AutoFit();
                    range.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    #endregion

                    #region Set Page/Printing Options
                    var _with1 = sheet.PageSetup;
                    _with1.PaperSize = Excel.XlPaperSize.xlPaperLegal;
                    _with1.Orientation = Excel.XlPageOrientation.xlLandscape;
                    _with1.FitToPagesWide = 1;
                    _with1.FitToPagesTall = false;
                    _with1.Zoom = false;
                    #endregion

                    #region Save Report
                    book.SaveAs(
                        ActiveReport.ExportFile,
                        Excel.XlFileFormat.xlWorkbookDefault,
                        Type.Missing,
                        Type.Missing,
                        false,
                        false,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing
                    );
                    #endregion
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error saving report: {ex.StackTrace}");
                    MessageBox.Show($"Error saving report: {ex.Message}", "Error Saving Report");
                }
                finally
                {
                    book.Close(true, Missing.Value, Missing.Value);
                    Global.ReleaseExcelProcesses();
                    OnProcessingComplete(new EventArgs());
                }
            });
            #endregion
        }
    }
}
