namespace ClosedXML.Excel.CalcEngine
{
    internal static class CalcEngineHelpers
    {
        internal static bool TryExtractRange(Expression expression, out IXLRange? range, out XLError calculationErrorType)
        {
            range = null;
            calculationErrorType = default;

            if (expression is not XObjectExpression objectExpression)
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            if (objectExpression.Value is not CellRangeReference cellRangeReference)
            {
                calculationErrorType = XLError.NoValueAvailable;
                return false;
            }

            range = cellRangeReference.Range;
            return true;
        }
    }
}
