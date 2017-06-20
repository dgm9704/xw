namespace xw
{
    public struct TableSize
    {
        public int Columns { get; }
        public int Rows { get; }

        public TableSize(int columns, int rows)
        {
            Columns = columns;
            Rows = rows;
        }
    }
}