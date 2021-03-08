using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.Errors
{
    public class RowException : Exception
    {
        /// <summary> Номер строки вызвавшей ошибку </summary>
        public uint RowNumder { get; set; }

        public RowException(string message, uint rowNumder)
            : base(message)
        {
            RowNumder = rowNumder;
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
