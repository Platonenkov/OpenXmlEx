namespace OpenXmlEx.Errors.Rows
{
    public class RowNotOpenException : RowException
    {
        public RowNotOpenException(string message, uint rowNumder, string method_name) : base(message, rowNumder, method_name)
        {
        }
    }
}