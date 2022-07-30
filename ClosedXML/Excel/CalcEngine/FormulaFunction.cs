using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FormulaFunction
    {
        private readonly CalcEngineFunction _method;
        private readonly FunctionFlags _flags;

        public FormulaFunction(CalcEngineFunction method, int parmMin, int parmMax, FunctionFlags flags)
        {
            _method = method;
            ParmMin = parmMin;
            ParmMax = parmMax;
            _flags = flags;
        }

        public int ParmMin { get; }

        public int ParmMax { get; }

        public AnyValue CallFunction(CalcContext ctx, params AnyValue?[] args)
        {
            return _method(ctx, args);
        }
    }
}
