namespace OpenXmlEx.SubClasses
{
    public sealed record OpenXmlMergedCellEx
    {
        public uint StartCell { get; }
        public uint StartRow { get; }
        public uint EndCell { get; }
        public uint EndRow { get; }

        public OpenXmlMergedCellEx(uint StartCell, uint StartRow, uint EndCell, uint EndRow)
        {
            this.StartCell = StartCell;
            this.StartRow = StartRow;
            this.EndCell = EndCell;
            this.EndRow = EndRow;
        }
    }
}
