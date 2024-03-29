﻿using System;

namespace OpenXmlEx.Errors.Actions
{
    public class SetWidthException : Exception
    {
        public string SheetName { get; }
        public string MethodName { get; }

        public SetWidthException(string message, string method_name) : base(message)
        {
            MethodName = method_name;
        }
        public SetWidthException(string message,string sheet_name,string method_name) : base(message)
        {
            SheetName = sheet_name;
            MethodName = method_name;
        }
    }
}