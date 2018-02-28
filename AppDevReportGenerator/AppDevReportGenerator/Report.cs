using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows;
using System.Windows.Input;
using System.Diagnostics;

namespace AppDevReportGenerator
{
    public class Report
    {
        public string Name { get; set; }
        public string ReportDefinitionFile { get; set; }
        public List<string> DefinitionHeaders { get; set; }
        public string SourceFolder { get; set; }
        public string SourceFile { get; set; }
        public string ExportFolder { get; set; }
        public string ExportFile { get; set; }
        public DateTime SourceDate { get; set; }
        public DateTime ExportDate { get; set; }
        public bool Enabled { get; set; }
        public string Stripe { get; set; }
        public List<ReportField> Fields { get; set; }
        public List<Row> Rows { get; set; }
        public string Divider { get; set; }
        public string DividerBackground { get; set; }
        public string HeaderBackground { get; set; }
        public int FirstRowIndex { get; set; }

        public Report() { }

        public Report(string jsonfile)
        {
            DeserializeReport(jsonfile);
        }

        private void DeserializeReport(string jsonfile)
        {
            Report source = new Report();            
            using (StreamReader file = File.OpenText(jsonfile))
            {
                JsonSerializer serializer = new JsonSerializer();
                source = (Report)serializer.Deserialize(file, typeof(Report));
                SourceFile = source.SourceFile;
                Name = source.Name;
                SourceFolder = source.SourceFolder;
                ExportFolder = source.ExportFolder;
                SourceDate = new FileInfo(jsonfile).CreationTime;
                Enabled = source.Enabled;
                Stripe = source.Stripe;
                Divider = source.Divider;
                DividerBackground = source.DividerBackground;
                HeaderBackground = source.HeaderBackground;
                FirstRowIndex = source.FirstRowIndex;
                Fields = source.Fields.OrderBy(x => x.ExportIndex).ToList<ReportField>();
            }
            ReportDefinitionFile = jsonfile;
            ExportFile = GetExportFileName();
            FileInfo fi = new FileInfo(SourceFile);
            ExportDate = fi.LastWriteTime;
        }

        public void GetReportHeaders()
        {
            #region Prepare Excel
            Global.ExcelApplication = new Excel.Application() { Visible = false };
            Global.ExcelApplication.UserControl = false;
            Global.ExcelApplication.DisplayAlerts = false;
            Excel.Workbook book = Global.ExcelApplication.Workbooks.Open(GetSourceFileName());
            Excel.Worksheet sheet = book.Sheets[1];
            Excel.Range range = sheet.UsedRange;
            #endregion

            DefinitionHeaders = new List<string>();
            Excel.Range headerrange = sheet.Range[
                sheet.Cells[2, 1],
                sheet.Cells[2, range.Columns.Count]
            ];
            foreach (Excel.Range cell in headerrange)
            {
                if(cell == null || cell.Value2 == null) { continue; }
                DefinitionHeaders.Add(cell.Value2.ToString());
            }

            book.Close(true, Missing.Value, Missing.Value);
            Global.ReleaseExcelProcesses();
        }

        public ReportField GetDividerField()
        {
            return Fields.Where(x => x.ExportName.ToLower() == Divider.ToLower()).First();
        }

        public bool IsRowEmpty(Excel.Range range, int rowindex)
        {
            string rowcontent = string.Empty;
            foreach(ReportField field in Fields)
            {
                string exceltext = GetExcelText(range.Cells[rowindex, field.SourceIndex].Value2,field);
                if(exceltext == field.NullValue) { exceltext = string.Empty; }
                rowcontent += string.IsNullOrEmpty(exceltext) ? string.Empty : exceltext;
            }
            return string.IsNullOrEmpty(rowcontent);
        }

        public Row GetDividerRow(Excel.Range range)
        {
            if (string.IsNullOrEmpty(Divider)) { return null; }  // report has no divider

            Row dividerrow = new Row { IsDivider = true, AddedToReport = true, Cells = new List<Cell>() };
            bool match = false;
            string comparevalue = string.Empty;            
            int i = FirstRowIndex;

            foreach(Row row in Rows)
            {
                ReportField dividerfield = GetDividerField();
                if (match == false)
                {
                    string currentvalue = GetExcelText(range.Cells[i, dividerfield.SourceIndex].Value2, dividerfield);
                    if (!string.IsNullOrEmpty(comparevalue) && comparevalue != currentvalue)
                    {
                        foreach (ReportField field in Fields.Where(x => x.ExportIndex > 0))
                        {
                            dividerrow.Cells.Add(new Cell { ColumnNumber = field.ExportIndex, Value = string.Empty, ColumnWidth = field.ColumnWidth });
                        }
                        match = true;
                        dividerrow.Index = i- FirstRowIndex;
                        return dividerrow;
                    }
                    else
                    {
                        comparevalue = currentvalue;
                    }
                }
                i++;
            }
            return null;
        }

        public void LoadReportRows()
        {
            Rows = new List<Row>();

            #region Prepare Excel
            Global.ExcelApplication = new Excel.Application() { Visible = false };
            Global.ExcelApplication.UserControl = false;
            Global.ExcelApplication.DisplayAlerts = false;
            Excel.Workbook book = Global.ExcelApplication.Workbooks.Open(GetSourceFileName());
            Excel.Worksheet sheet = book.Sheets[1];
            Excel.Range range = sheet.UsedRange;
            #endregion

            try
            {
                #region Add Normal Rows to Rows
                for (int i = FirstRowIndex; i <= range.Rows.Count; i++)
                {                        
                    Row row = new Row { Cells = new List<Cell>() };
                    foreach (ReportField field in Fields)
                    {
                        string textvalue = GetExcelText(range.Cells[i, field.SourceIndex].Value2, field);
                        row.Cells.Add(new Cell { ColumnNumber = field.ExportIndex, Value = textvalue, ColumnWidth = field.ColumnWidth });
                    }
                    if (!IsRowEmpty(range, i))  // sometimes there is a blank line at the end of the excel file, this accounts for it
                    {
                        Rows.Add(row);
                    }
                }
                #endregion

                #region Add Divider Row
                Row dividerrow = GetDividerRow(range);
                if(dividerrow != null)
                {
                    Rows.Insert(dividerrow.Index, dividerrow);
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
        }

        public string GetSourceFileName() { return $"{SourceFolder}\\{SourceFile}"; }

        public string GetExportFileName(){ return $"{ExportFolder}\\{Name}_{DateTime.Now.AddDays(1).ToString("MMddyy")}.xlsx"; }

        public string GetExcelText(object value, ReportField field)
        {
            string result = field.NullValue;
            if (value != null)
            {
                switch (field.DataType)
                {
                    case "string":
                        result = value.ToString().Replace("|", "\r\n");  // Unfortunately, newlines do not export from ServicePro properly, users must use this token to indicate where a newline should be inserted
                        break;
                    case "date":
                        result = Convert.ToDateTime(value.ToString()).ToString("MM/dd/yyyy");
                        break;
                    default:  // int and bool
                        result = value.ToString();
                        break;
                }
            }
            return result;
        }
    }
}
