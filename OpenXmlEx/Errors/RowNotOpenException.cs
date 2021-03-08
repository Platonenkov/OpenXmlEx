namespace OpenXmlEx.Errors
{
    public class RowNotOpenException : RowException
    {
        public RowNotOpenException(string message, uint rowNumder) : base(message, rowNumder)
        {
        }
    }
}