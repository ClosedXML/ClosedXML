namespace ClosedXML.Excel.CalcEngine.Functions;

internal interface ITallyState<out TState>
{
    TState Tally(double number);
}
