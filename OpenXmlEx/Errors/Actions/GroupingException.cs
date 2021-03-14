using System;

namespace OpenXmlEx.Errors.Actions
{
    public class GroupingException : Exception
    {
        public string MethodName { get; }
        public string SheetName { get; }
        public GroupingException(string message,string method_name) : base(message)
        {
            MethodName = method_name;
        }
        public GroupingException(string message,string sheet_name, string method_name) : base(message)
        {
            SheetName = sheet_name;
            MethodName = method_name;
        }
    }
}
