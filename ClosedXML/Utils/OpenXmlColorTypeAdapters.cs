// Keep this file CodeMaid organised and cleaned
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Utils
{
    internal interface IColorTypeAdapter
    {
        public BooleanValue Auto { get; set; }
        public UInt32Value Indexed { get; set; }
        public HexBinaryValue Rgb { get; set; }
        public UInt32Value Theme { get; set; }
        public DoubleValue Tint { get; set; }
    }

    internal class ColorTypeAdapter : IColorTypeAdapter
    {
        public ColorTypeAdapter(ColorType colorType)
        {
            this.ColorType = colorType;
        }

        #region ColorType

        public ColorType ColorType { get; }

        #endregion ColorType

        public BooleanValue Auto { get => ColorType.Auto; set => ColorType.Auto = value; }
        public UInt32Value Indexed { get => ColorType.Indexed; set => ColorType.Indexed = value; }
        public HexBinaryValue Rgb { get => ColorType.Rgb; set => ColorType.Rgb = value; }
        public UInt32Value Theme { get => ColorType.Theme; set => ColorType.Theme = value; }
        public DoubleValue Tint { get => ColorType.Tint; set => ColorType.Tint = value; }
    }

    internal class X14ColorTypeAdapter : IColorTypeAdapter
    {
        public X14ColorTypeAdapter(X14.ColorType colorType)
        {
            this.ColorType = colorType;
        }

        #region ColorType

        public X14.ColorType ColorType { get; }

        #endregion ColorType

        public BooleanValue Auto { get => ColorType.Auto; set => ColorType.Auto = value; }
        public UInt32Value Indexed { get => ColorType.Indexed; set => ColorType.Indexed = value; }
        public HexBinaryValue Rgb { get => ColorType.Rgb; set => ColorType.Rgb = value; }
        public UInt32Value Theme { get => ColorType.Theme; set => ColorType.Theme = value; }
        public DoubleValue Tint { get => ColorType.Tint; set => ColorType.Tint = value; }
    }
}
