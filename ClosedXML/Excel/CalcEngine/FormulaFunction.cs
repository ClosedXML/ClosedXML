using OneOf;
using System;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FormulaFunction
    {
        private readonly CalcEngineFunction _method;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="method"></param>
        /// <param name="parmMin">Minimum amount of parameters, useful only for variable number of parameters function.</param>
        /// <param name="parmMax">Minimum amount of parameters, useful only for variable number of parameters function.</param>
        public FormulaFunction(CalcEngineFunction method, int parmMin, int parmMax)
        {
            _method = method;
            ParmMin = parmMin;
            ParmMax = parmMax;
        }

        public int ParmMin { get; }

        public int ParmMax { get; }

        public AnyValue CallFunction(CalcContext ctx, params AnyValue?[] args)
        {
            return _method(ctx, args);
        }
    }
}
