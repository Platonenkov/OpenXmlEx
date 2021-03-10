namespace OpenXmlEx.Errors.Rows
{
    internal class RowNotClosedException : RowException
    {
        public RowNotClosedException(string message, uint rowNumder,string method_name) : base(message, rowNumder,method_name)
        {
        }
    }
}