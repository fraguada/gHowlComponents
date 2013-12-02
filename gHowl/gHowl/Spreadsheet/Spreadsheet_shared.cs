namespace gHowl.Spreadsheet
{
    class CellAddress
    {
        public uint Row;
        public uint Col;
        public string IdString;
        public string DataString;

        /**
         * Constructs a CellAddress representing the specified {@code row} and
         * {@code col}. The IdString will be set in 'RnCn' notation.
         */
        public CellAddress(uint row, uint col)
        {
            this.Row = row;
            this.Col = col;
            this.IdString = string.Format("R{0}C{1}", row, col);
        }

        public CellAddress(uint row, uint col, string data)
        {
            this.Row = row;
            this.Col = col;
            this.DataString = data;
            this.IdString = string.Format("R{0}C{1}", row, col);
        }
    }



}