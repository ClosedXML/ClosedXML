using System;

namespace ClosedXML.Excel.CalcEngine.Functions;

internal interface ITally
{
    OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>;
}
