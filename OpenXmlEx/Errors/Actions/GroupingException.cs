﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.Errors.Actions
{
    internal class GroupingException : Exception
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
