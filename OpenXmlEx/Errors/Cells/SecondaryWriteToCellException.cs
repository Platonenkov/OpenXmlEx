namespace OpenXmlEx.Errors.Cells
{
    internal class SecondaryWriteToCellException : CellException
    {
        public SecondaryWriteToCellException(string message, uint rowNumder, uint cellNumber, string excelAddress, string method_name)
            : base(message, rowNumder, cellNumber, excelAddress, method_name)
        {
        }
    }
}