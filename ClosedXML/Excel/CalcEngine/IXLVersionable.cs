namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// An interface for marking entities that track <see cref="XLWorkbook.RecalculationCounter"/> when modified.
    /// Serves for determining when <see cref="XLCell.CachedValue"/> has to be re-evaluated.
    /// </summary>
    internal interface IXLVersionable
    {
        long ModifiedAtVersion { get; }
    }
}
