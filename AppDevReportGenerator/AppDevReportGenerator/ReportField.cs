namespace AppDevReportGenerator
{
    public class ReportField
    {
        public string Name { get; set; }
        public int SourceIndex { get; set; }
        public int ExportIndex { get; set; }
        public int ColumnWidth { get; set; }
        public string DataType { get; set; }
        public string NullValue { get; set; }
        public string ExportName { get; set; }
    }
}
