namespace ClosedXML.Excel
{
    internal class XLDrawingStyle : IXLDrawingStyle
    {
        public XLDrawingStyle()
        {
            //Font = new XLDrawingFont(this);
            Alignment = new XLDrawingAlignment(this);
            ColorsAndLines = new XLDrawingColorsAndLines(this);
            Size = new XLDrawingSize(this);
            Protection = new XLDrawingProtection(this);
            Properties = new XLDrawingProperties(this);
            Margins = new XLDrawingMargins(this);
            Web = new XLDrawingWeb(this);
        }

        //public IXLDrawingFont Font { get; private set; }
        public IXLDrawingAlignment Alignment { get; private set; }

        public IXLDrawingColorsAndLines ColorsAndLines { get; private set; }
        public IXLDrawingSize Size { get; private set; }
        public IXLDrawingProtection Protection { get; private set; }
        public IXLDrawingProperties Properties { get; private set; }
        public IXLDrawingMargins Margins { get; private set; }
        public IXLDrawingWeb Web { get; private set; }

        public static IXLDrawingStyle DefaultCommentStyle
        {
            get
            {
                var defaultCommentStyle = new XLDrawingStyle();

                defaultCommentStyle
                    .Margins.SetLeft(0.1)
                    .Margins.SetRight(0.1)
                    .Margins.SetTop(0.05)
                    .Margins.SetBottom(0.05)
                    .Margins.SetAutomatic()
                    .Size.SetHeight(59.25)
                    .Size.SetWidth(19.2)
                    .ColorsAndLines.SetLineColor(XLColor.Black)
                    .ColorsAndLines.SetFillColor(XLColor.FromArgb(255, 255, 225))
                    .ColorsAndLines.SetLineDash(XLDashStyle.Solid)
                    .ColorsAndLines.SetLineStyle(XLLineStyle.Single)
                    .ColorsAndLines.SetLineWeight(0.75)
                    .ColorsAndLines.SetFillTransparency(1)
                    .ColorsAndLines.SetLineTransparency(1)
                    .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Left)
                    .Alignment.SetVertical(XLDrawingVerticalAlignment.Top)
                    .Alignment.SetDirection(XLDrawingTextDirection.LeftToRight)
                    .Alignment.SetOrientation(XLDrawingTextOrientation.LeftToRight)
                    .Properties.SetPositioning(XLDrawingAnchor.Absolute)
                    .Protection.SetLocked()
                    .Protection.SetLockText();

                return defaultCommentStyle;
            }
        }
    }
}
