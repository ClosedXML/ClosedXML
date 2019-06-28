using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotSourceReference : IEquatable<IXLPivotSourceReference>
    {
        IXLRange SourceRange { get; set; }
        IXLTable SourceTable { get; set; }
        XLPivotTableSourceType SourceType { get; }
    }
}
