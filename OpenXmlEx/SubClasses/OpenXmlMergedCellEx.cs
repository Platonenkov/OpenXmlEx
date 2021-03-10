using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.SubClasses
{
    public sealed record OpenXmlMergedCellEx(uint StartCell, uint StartRow, uint EndCell, uint EndRow);
}
