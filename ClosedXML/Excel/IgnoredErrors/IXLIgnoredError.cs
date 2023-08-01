namespace ClosedXML.Excel
{
    public enum XLIgnoredErrorType
    {
        CalculatedColumn = 1,
        EmptyCellReference,
        EvalError,
        Formula,
        FormulaRange,
        ListDataValidation,
        NumberAsText,
        TwoDigitTextYear,
        UnlockedFormula
    }

    public interface IXLIgnoredError
    {
        XLIgnoredErrorType Type { get; }
        IXLRange Range { get; }
    }
}
