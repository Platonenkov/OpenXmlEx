namespace OpenXmlEx.Errors
{
    public class RowNotClosedException : RowException
    {
        public RowNotClosedException(string message, uint rowNumder) : base(message, rowNumder)
        {
        }
    }
}