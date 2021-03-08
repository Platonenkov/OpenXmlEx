namespace OpenXmlEx.Errors
{
    public class SecondaryWriteToCellException : CellException
    {
        public SecondaryWriteToCellException(string message, uint rowNumder, uint cellNumber, string excelAddress) 
            : base(message, rowNumder, cellNumber, excelAddress)
        {
        }
    }
}