using System;

namespace OpenXmlEx.Errors.Actions
{
    internal class FilterException : Exception
    {
        public string MethodName { get; }
        public string SheetName { get; }
        public FilterException(string message,string method_name) : base(message)
        {
            MethodName = method_name;
        }
        public FilterException(string message,string sheet_name, string method_name) : base(message)
        {
            SheetName = sheet_name;
            MethodName = method_name;
        }
    }
}