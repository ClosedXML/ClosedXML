using System.Collections.Generic;
using System.Reflection;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FunctionRegistry
    {
        private readonly Dictionary<string, FormulaFunction> _fnTbl = new();

        public void RegisterFunction(string functionName, MethodInfo fn)
        {
            _fnTbl.Add(functionName, new FormulaFunction(fn, 0, int.MaxValue));
        }

        public void RegisterFunction(string functionName, MethodInfo fn, int parmMin, int parmMax)
        {
            _fnTbl.Add(functionName, new FormulaFunction(fn, parmMin, parmMax));
        }

        public bool TryGetFunc(string name, out FormulaFunction func)
        {
            return _fnTbl.TryGetValue(name, out func);
        }
    }
}
