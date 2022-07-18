namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Function definition class (keeps function name, parameter counts, and delegate).
    /// </summary>
    internal class FunctionDefinition
    {
        // ** fields
        public int ParmMin, ParmMax;
        public CalcEngineFunction Function;

        // ** ctor
        public FunctionDefinition(string name, int parmMin, int parmMax, CalcEngineFunction function)
        {
            Name = name;
            ParmMin = parmMin;
            ParmMax = parmMax;
            Function = function;
        }

        public string Name { get; }
    }
}
