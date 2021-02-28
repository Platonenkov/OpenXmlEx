namespace OpenXmlEx
{
    public static class ColorExtensions
    {
        #region Colors

        public static string ToHexConverter(this System.Drawing.Color c)
            => "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");

        public static string RGBConverter(this System.Drawing.Color c)
            => "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";

        #endregion
    }
}