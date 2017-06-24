namespace xw
{
    using OfficeOpenXml;

    public static class ExcelWorksheetExtensions
    {
        public static void Pretty(this ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns();
        }
    }
}