namespace xw
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using XbrlTable;

    public static class TableExtensions
    {
        private static Axis DefaultYAxis = new Axis(0, Direction.Y, false, new OrdinateCollection { new Ordinate("999", "0", new Signature()) });

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
                    DefaultIfEmpty(DefaultYAxis).
                    First().
                    Ordinates.
                    OrderBy(o => o.Path).
                    Select(o => o.Code)).
                Select((o, i) => Tuple.Create(o, i)).
                ToList();
        }
    }
}