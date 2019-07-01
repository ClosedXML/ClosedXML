// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    /// <summary>
    /// Relative or absolute reference to a cell or a range
    /// </summary>
    internal interface IXLReference
    {
        #region Public Methods

        string ToStringA1(IXLAddress baseAddress);

        string ToStringR1C1();

        #endregion Public Methods
    }
}
