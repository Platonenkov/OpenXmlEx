using System;

namespace OpenXmlEx.Errors.Sheets
{
    public class SheetException : Exception
    {
        public string SheetName { get; }
        public string MethodName { get; }
        public SheetException(string message, string sheet_name, string method_name) : base(message)
        {
            MethodName = method_name;
            SheetName = sheet_name;
        }
    }
}
