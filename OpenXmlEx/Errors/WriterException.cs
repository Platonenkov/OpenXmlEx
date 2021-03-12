using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.Errors
{
    internal class WriterException : Exception
    {
        public string MethodName { get; }
        public string SheetName { get; }
        public WriterException(string message, string method_name) : base(message)
        {
            MethodName = method_name;
        }
        public WriterException(string message, string sheet_name, string method_name) : base(message)
        {
            SheetName = sheet_name;
            MethodName = method_name;
        }
    }
}
