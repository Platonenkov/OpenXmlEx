using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.Errors.Actions
{
    public class GroupingException : Exception
    {
        public string SheetName { get; }

        public GroupingException(string message,string sheet_name) : base(message)
        {
            SheetName = sheet_name;
        }
    }
}
