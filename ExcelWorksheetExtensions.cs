namespace xw
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using OfficeOpenXml;
    using XbrlTable;

    public static class ExcelWorksheetExtensions
    {
        public static TableSize AddTable(this ExcelWorksheet worksheet, Table table, int row, int col)
        {
            worksheet.Cells[row, col].Value = table.Code;
            var size = worksheet.PlaceTableAxes(table, row, col);
            //PlaceNames(table, worksheet, row, col);
            return size;
        }

        public static TableSize PlaceTableAxes(this ExcelWorksheet worksheet, Table table, int row, int col)
        {
            worksheet.PlaceZAxis(table, row, col + 2);

            var columns = table.GetColumns(); ;
            worksheet.PlaceColumns(columns, row + 2, col + 2);

            var rows = table.GetRows();
            worksheet.PlaceRows(rows, row + 3, col + 1);

            worksheet.PlaceCellNames(rows, columns, table.Code, row, col);

            return new TableSize(columns.Count, rows.Count);
        }

        private static void PlaceCellNames(this ExcelWorksheet worksheet, List<Tuple<string, int>> rows, List<Tuple<string, int>> columns, string tableCode, int row, int col)
        {
            var cells = rows.SelectMany(r => columns.Select(c => (new[] { r, c }))).ToList();
            cells.ForEach(rc => worksheet.Cells[row + 3 + rc.First().Item2, col + 2 + rc.Last().Item2].Value = GetCellName(tableCode, rc));
            cells.ForEach(rc => worksheet.Names.Add(GetCellName(tableCode, rc), worksheet.Cells[row + 3 + rc.First().Item2, col + 2 + rc.Last().Item2]));
        }

        private static string GetCellName(string tableCode, Tuple<string, int>[] rc)
        {
            var result = $"{CleanTableCode(tableCode)}_{CleanOrdinateCode(rc.First().Item1)}_{CleanOrdinateCode(rc.Last().Item1)}";
            result = rc.Last().Item1.StartsWith("*")
                ? $"{result}_xkey"
                : rc.First().Item1.StartsWith("*")
                    ? $"{result}_ykey"
                    : result;
            return result;
        }

        private static string CleanTableCode(string value)
        {
            return string.IsNullOrEmpty(value)
                ? ""
                : value.Replace(" ", "").ToUpper();
        }

        private static string CleanOrdinateCode(string value)
        {
            return string.IsNullOrEmpty(value)
                 ? ""
                 : value.Replace(" ", "").Replace("*", "").ToUpper();
        }

        private static void PlaceColumns(this ExcelWorksheet worksheet, List<Tuple<string, int>> columns, int row, int col)
            => columns.ForEach(o => worksheet.Cells[row, col + o.Item2].Value = o.Item1);

        private static void PlaceRows(this ExcelWorksheet worksheet, List<Tuple<string, int>> rows, int row, int col)
            => rows.ForEach(o => worksheet.Cells[row + o.Item2, col].Value = o.Item1);

        private static void PlaceZAxis(this ExcelWorksheet worksheet, Table table, int row, int col)
        {
            table.
            Axes.
            Where(a => a.Direction == Direction.Z).
            Where(a => a.IsOpen).
            OrderBy(a => a.Order).
            SelectMany(a => a.Ordinates).
            OrderBy(o => o.Path).
            Select(o => $"*{o.Code}").
            Select((o, i) => Tuple.Create(o, col + i)).
            ToList().
            ForEach(o => worksheet.Cells[row, o.Item2].Value = o.Item1);
        }

        public static void Pretty(this ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns();
        }
    }
}