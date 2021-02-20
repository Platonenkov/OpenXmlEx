using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    public static class OpenXmlExStyles
    {

        public static Stylesheet GetStylesheet()
        {
            return new Stylesheet(
                new Fonts(OpenXmlExStyleFont.Fonts.Values),
                new Fills(OpenXmlExStyleFill.Fills.Values), 
                new Borders(OpenXmlExStyleBorderGrand.Borders.Values),
                new CellFormats(OpenXmlExStyleCell.CellsFormats.Values));
        }
    }
}
