using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLIgnoredErrors : IXLIgnoredErrors
    {
        private readonly List<XLIgnoredError> _ignoredErrors = new();

        public XLWorksheet Worksheet { get; internal set; }

        public XLIgnoredErrors(XLWorksheet targetSheet)
        {
            Worksheet = targetSheet;
        }

        public XLIgnoredErrors(XLWorksheet targetSheet, XLIgnoredErrors ignoredErrors)
        {
            Worksheet = targetSheet;

            _ignoredErrors.AddRange(ignoredErrors._ignoredErrors.Select(x =>
                    new XLIgnoredError(x.Type, targetSheet.Range(((XLRangeAddress)x.Range.RangeAddress).WithoutWorksheet()))));
        }

        public IEnumerator<IXLIgnoredError> GetEnumerator()
        {
            return _ignoredErrors.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _ignoredErrors.GetEnumerator();
        }

        public void Clear()
        {
            _ignoredErrors.Clear();
        }

        public void Add(XLIgnoredError ignoredError)
        {
            _ignoredErrors.Add(ignoredError);
        }

        public void Add(XLIgnoredErrorType type, IXLRange range)
        {
            Add(new XLIgnoredError(type, range));
        }
    }
}
