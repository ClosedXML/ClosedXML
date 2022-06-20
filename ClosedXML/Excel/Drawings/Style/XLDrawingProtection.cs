namespace ClosedXML.Excel
{
    internal class XLDrawingProtection : IXLDrawingProtection
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingProtection(IXLDrawingStyle style)
        {
            _style = style;
        }
        public bool Locked { get; set; }	public IXLDrawingStyle SetLocked() { Locked = true; return _style; }	public IXLDrawingStyle SetLocked(bool value) { Locked = value; return _style; }
        public bool LockText { get; set; }	public IXLDrawingStyle SetLockText() { LockText = true; return _style; }	public IXLDrawingStyle SetLockText(bool value) { LockText = value; return _style; }

    }
}
