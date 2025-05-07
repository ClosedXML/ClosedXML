using System;

namespace ClosedXML.IO.CodeGen.Model;

public readonly record struct Occurrences(int? Min, int? Max)
{
    internal ElementsCount Elements
    {
        get
        {
            var min = Min ?? 1;
            var max = Max ?? 1;
            return (min, max) switch
            {
                (0, 1) => ElementsCount.ZeroToOne,
                (0, int.MaxValue) => ElementsCount.ZeroToMany,
                (1, 1) => ElementsCount.OneToOne,
                (1, int.MaxValue) => ElementsCount.OneToMany,
                _ => throw new NotSupportedException()
            };
        }
    }
};
