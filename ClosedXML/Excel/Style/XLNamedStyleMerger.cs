using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Class responsible for providing correct style names when copined from one workbook to
    /// another.
    ///  * If target workbook does not contain style with the specified name the original name is preserved.
    ///  * If target workbook does have style with the same name and it is equivalent to the source style
    ///    the same name is used.
    ///  * If styles differ merger tries "Name 1", "Name 2" and so on until it finds either not-used style
    ///    name or equivalent style.
    /// </summary>
    internal class XLNamedStyleMerger
    {
        private readonly XLWorkbook _sourceWorkbook;
        private readonly XLWorkbook _targetWorkbook;
        private readonly Dictionary<string, string> _styleMapping;
        private readonly bool _sameWorkbook;

        public XLNamedStyleMerger(XLWorkbook sourceWorkbook, XLWorkbook targetWorkbook)
        {
            _sourceWorkbook = sourceWorkbook ?? throw new ArgumentNullException(nameof(sourceWorkbook));
            _targetWorkbook = targetWorkbook ?? throw new ArgumentNullException(nameof(targetWorkbook));
            _sameWorkbook = ReferenceEquals(_sourceWorkbook, targetWorkbook);
            _styleMapping = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
        }

        public string GetTargetStyleName(XLStyleValue styleInSourceWorkbook)
        {
            if (_sameWorkbook)
                return styleInSourceWorkbook.Name;

            if (_styleMapping.ContainsKey(styleInSourceWorkbook.Name))
                return _styleMapping[styleInSourceWorkbook.Name];

            var sourceStyleName = styleInSourceWorkbook.Name;
            int i = 1;
            string res = null;
            do
            {
                var styleInTargetWorkbook = _targetWorkbook.NamedStyles[sourceStyleName];

                if (styleInTargetWorkbook == null)
                {
                    res = sourceStyleName;
                }

                if (styleInTargetWorkbook == styleInSourceWorkbook)
                {
                    // Names may differ in case
                    res = styleInTargetWorkbook.Name;
                }

                sourceStyleName = $"{styleInSourceWorkbook.Name} {i}";
                i++;
            } while (res == null);

            _styleMapping.Add(styleInSourceWorkbook.Name, res);

            return res;
        }
    }
}
