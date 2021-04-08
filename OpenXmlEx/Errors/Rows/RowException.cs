using System;

namespace OpenXmlEx.Errors.Rows
{
    public class RowException : Exception
    {
        /// <summary> Номер строки вызвавшей ошибку </summary>
        public uint RowNumder { get; set; }
        public string MethodName { get; }

        public RowException(string message, uint rowNumder, string method_name)
            : base(message)
        {
            RowNumder = rowNumder;
            MethodName = method_name;
        }

        #region Overrides of Exception
        /// <summary>
        /// Вывод информации об ошибки
        /// </summary>
        /// <returns></returns>
        public override string ToString() => $"In row number - {RowNumder}\n{base.ToString()}";

        #endregion
    }
}
