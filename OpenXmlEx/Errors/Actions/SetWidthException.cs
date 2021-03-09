using System;

namespace OpenXmlEx.Errors.Actions
{
    public class SetWidthException : Exception
    {
        public string SheetName { get; }

        public SetWidthException(string message,string sheet_name) : base(message)
        {
            SheetName = sheet_name;
        }
    }
}