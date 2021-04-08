using System;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx.Errors.Actions
{
    public class MergeCellException : Exception
    {
        /// <summary> пересечение с диапазоном </summary>
        public OpenXmlMergedCellEx InRange;
        /// <summary> Входной диапазон </summary>
        public OpenXmlMergedCellEx Range { get; }
        /// <summary> имя метода вызвавшего ошибку </summary>
        public string MethodName { get; }

        public MergeCellException(string message, OpenXmlMergedCellEx range, OpenXmlMergedCellEx inRange, string method_name) : base(message)
        {
            InRange = inRange;
            Range = range;
            MethodName = method_name;
        }
    }
}