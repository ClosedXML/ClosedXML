namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>A non-state representation of a formula that can be used by many cells.</summary>
    internal class Formula
    {
        public Formula(string text, ValueNode root, FormulaFlags flags)
        {
            AstRoot = root;
            Text = text;
            Flags = flags;
        }

        /// <summary>Text of the formula.</summary>
        public string Text { get; }

        public ValueNode AstRoot { get; }

        public FormulaFlags Flags { get; }
    }
}
