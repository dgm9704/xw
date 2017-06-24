namespace xw
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using OfficeOpenXml;
    using XbrlTable;

    public static class TableExtensions
    {
        public static List<Tuple<string, int>> GetColumns(this Table table)
        {
            return
                table.
                Axes.
                Where(a => a.IsOpen).
                Where(a => a.Direction == Direction.Y).
                SelectMany(a => a.Ordinates).
                OrderBy(o => o.Path).
                Select(o => $"*{o.Code}").
                Concat(
                    table.Axes.
                    Where(a => !a.IsOpen).
                    FirstOrDefault(a => a.Direction == Direction.X).
                    Ordinates.
                    OrderBy(o => o.Path).
                    Select(o => o.Code)).
                Select((o, i) => Tuple.Create(o, i)).
                ToList();
        }

        public static List<Tuple<string, int>> GetRows(this Table table)
        {
            return
                table.
                Axes.
                Where(a => a.IsOpen).
                Where(a => a.Direction == Direction.X).
                SelectMany(a => a.Ordinates).
                OrderBy(o => o.Path).
                Select(o => $"*{o.Code}").
                Concat(
                    table.
                    Axes.
                    Where(a => !a.IsOpen).
                    Where(a => a.Direction == Direction.Y).
                    DefaultIfEmpty(Axis.DefaultYAxis).
                    First().
                    Ordinates.
                    OrderBy(o => o.Path).
                    Select(o => o.Code)).
                Select((o, i) => Tuple.Create(o, i)).
                ToList();
        }

        private static void PlaceZAxis(this Table table, ExcelWorksheet worksheet, ExcelCoordinate start)
        {
            table.
            Axes.
            Where(a => a.Direction == Direction.Z).
            Where(a => a.IsOpen).
            OrderBy(a => a.Order).
            SelectMany(a => a.Ordinates).
            OrderBy(o => o.Path).
            Select(o => $"*{o.Code}").
            Select((o, i) => Tuple.Create(o, start.Column + i)).
            ToList().
            ForEach(o => worksheet.Cells[start.Row, o.Item2].Value = o.Item1);
        }

        public static TableSize PlaceTableAxes(this Table table, ExcelWorksheet worksheet, ExcelCoordinate start)
        {
            table.PlaceZAxis(worksheet, start.Offset(0, 2));

            var columns = table.GetColumns();
            PlaceColumns(worksheet, columns, start.Offset(2, 2));

            var rows = table.GetRows();
            PlaceRows(worksheet, rows, start.Offset(3, 1));

            var size = new TableSize(columns.Count, rows.Count);
            PlaceCellNames(worksheet, rows, columns, table.Code, start.Offset(3, 2));
            PlaceDataArea(worksheet, table.Code, start.Offset(3, 2), size);

            return size;
        }

        public static ExcelCoordinate WriteToWorksheet(this Table table, ExcelWorksheet worksheet, ExcelCoordinate start)
        {
            worksheet.Cells[start.Row, start.Column].Value = table.Code;
            var size = table.PlaceTableAxes(worksheet, start);
            return start.Add(size);
        }

        private static void PlaceCellNames(ExcelWorksheet worksheet, List<Tuple<string, int>> rows, List<Tuple<string, int>> columns, string tableCode, ExcelCoordinate start)
        {
            var cells = rows.SelectMany(r => columns.Select(c => (new[] { r, c }))).ToList();
            cells.ForEach(rc => worksheet.Cells[start.Row + rc.First().Item2, start.Column + rc.Last().Item2].Value = GetCellName(tableCode, rc));
            cells.ForEach(rc => worksheet.Names.Add(GetCellName(tableCode, rc), worksheet.Cells[start.Row + rc.First().Item2, start.Column + rc.Last().Item2]));
        }

        private static void PlaceDataArea(ExcelWorksheet worksheet, string tableCode, ExcelCoordinate start, TableSize size)
            => worksheet.Names.Add(GetDataAreaName(tableCode), worksheet.Cells[start.Row, start.Column, start.Row + size.Rows - 1, start.Column + size.Columns - 1]);

        private static string GetDataAreaName(string tableCode)
            => $"{CleanTableCode(tableCode)}_DataArea";

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

        private static void PlaceColumns(ExcelWorksheet worksheet, List<Tuple<string, int>> columns, ExcelCoordinate start)
            => columns.ForEach(o => worksheet.Cells[start.Row, start.Column + o.Item2].Value = o.Item1);

        private static void PlaceRows(ExcelWorksheet worksheet, List<Tuple<string, int>> rows, ExcelCoordinate start)
            => rows.ForEach(o => worksheet.Cells[start.Row + o.Item2, start.Column].Value = o.Item1);

    }
}