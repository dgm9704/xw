namespace xw
{
    using System;

    public struct ExcelCoordinate
    {
        public int Row { get; }
        public int Column { get; }

        public ExcelCoordinate(int row, int column)
        {
            if (row < 1)
                throw new ArgumentOutOfRangeException(nameof(row));

            if (column < 1)
                throw new ArgumentOutOfRangeException(nameof(column));

            Row = row;
            Column = column;
        }

        public ExcelCoordinate Offset(int rows, int columns)
        {
            return new ExcelCoordinate(Row + rows, Column + columns);
        }

        public ExcelCoordinate Add(TableSize size)
        {
            return Offset(size.Rows, size.Columns);
        }
    }
}