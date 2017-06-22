using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel
{
    public sealed class XLTableTheme
    {
        public static readonly XLTableTheme None = new XLTableTheme("None");
        public static readonly XLTableTheme TableStyleMedium28 = new XLTableTheme("TableStyleMedium28");
        public static readonly XLTableTheme TableStyleMedium27 = new XLTableTheme("TableStyleMedium27");
        public static readonly XLTableTheme TableStyleMedium26 = new XLTableTheme("TableStyleMedium26");
        public static readonly XLTableTheme TableStyleMedium25 = new XLTableTheme("TableStyleMedium25");
        public static readonly XLTableTheme TableStyleMedium24 = new XLTableTheme("TableStyleMedium24");
        public static readonly XLTableTheme TableStyleMedium23 = new XLTableTheme("TableStyleMedium23");
        public static readonly XLTableTheme TableStyleMedium22 = new XLTableTheme("TableStyleMedium22");
        public static readonly XLTableTheme TableStyleMedium21 = new XLTableTheme("TableStyleMedium21");
        public static readonly XLTableTheme TableStyleMedium20 = new XLTableTheme("TableStyleMedium20");
        public static readonly XLTableTheme TableStyleMedium19 = new XLTableTheme("TableStyleMedium19");
        public static readonly XLTableTheme TableStyleMedium18 = new XLTableTheme("TableStyleMedium18");
        public static readonly XLTableTheme TableStyleMedium17 = new XLTableTheme("TableStyleMedium17");
        public static readonly XLTableTheme TableStyleMedium16 = new XLTableTheme("TableStyleMedium16");
        public static readonly XLTableTheme TableStyleMedium15 = new XLTableTheme("TableStyleMedium15");
        public static readonly XLTableTheme TableStyleMedium14 = new XLTableTheme("TableStyleMedium14");
        public static readonly XLTableTheme TableStyleMedium13 = new XLTableTheme("TableStyleMedium13");
        public static readonly XLTableTheme TableStyleMedium12 = new XLTableTheme("TableStyleMedium12");
        public static readonly XLTableTheme TableStyleMedium11 = new XLTableTheme("TableStyleMedium11");
        public static readonly XLTableTheme TableStyleMedium10 = new XLTableTheme("TableStyleMedium10");
        public static readonly XLTableTheme TableStyleMedium9 = new XLTableTheme("TableStyleMedium9");
        public static readonly XLTableTheme TableStyleMedium8 = new XLTableTheme("TableStyleMedium8");
        public static readonly XLTableTheme TableStyleMedium7 = new XLTableTheme("TableStyleMedium7");
        public static readonly XLTableTheme TableStyleMedium6 = new XLTableTheme("TableStyleMedium6");
        public static readonly XLTableTheme TableStyleMedium5 = new XLTableTheme("TableStyleMedium5");
        public static readonly XLTableTheme TableStyleMedium4 = new XLTableTheme("TableStyleMedium4");
        public static readonly XLTableTheme TableStyleMedium3 = new XLTableTheme("TableStyleMedium3");
        public static readonly XLTableTheme TableStyleMedium2 = new XLTableTheme("TableStyleMedium2");
        public static readonly XLTableTheme TableStyleMedium1 = new XLTableTheme("TableStyleMedium1");
        public static readonly XLTableTheme TableStyleLight21 = new XLTableTheme("TableStyleLight21");
        public static readonly XLTableTheme TableStyleLight20 = new XLTableTheme("TableStyleLight20");
        public static readonly XLTableTheme TableStyleLight19 = new XLTableTheme("TableStyleLight19");
        public static readonly XLTableTheme TableStyleLight18 = new XLTableTheme("TableStyleLight18");
        public static readonly XLTableTheme TableStyleLight17 = new XLTableTheme("TableStyleLight17");
        public static readonly XLTableTheme TableStyleLight16 = new XLTableTheme("TableStyleLight16");
        public static readonly XLTableTheme TableStyleLight15 = new XLTableTheme("TableStyleLight15");
        public static readonly XLTableTheme TableStyleLight14 = new XLTableTheme("TableStyleLight14");
        public static readonly XLTableTheme TableStyleLight13 = new XLTableTheme("TableStyleLight13");
        public static readonly XLTableTheme TableStyleLight12 = new XLTableTheme("TableStyleLight12");
        public static readonly XLTableTheme TableStyleLight11 = new XLTableTheme("TableStyleLight11");
        public static readonly XLTableTheme TableStyleLight10 = new XLTableTheme("TableStyleLight10");
        public static readonly XLTableTheme TableStyleLight9 = new XLTableTheme("TableStyleLight9");
        public static readonly XLTableTheme TableStyleLight8 = new XLTableTheme("TableStyleLight8");
        public static readonly XLTableTheme TableStyleLight7 = new XLTableTheme("TableStyleLight7");
        public static readonly XLTableTheme TableStyleLight6 = new XLTableTheme("TableStyleLight6");
        public static readonly XLTableTheme TableStyleLight5 = new XLTableTheme("TableStyleLight5");
        public static readonly XLTableTheme TableStyleLight4 = new XLTableTheme("TableStyleLight4");
        public static readonly XLTableTheme TableStyleLight3 = new XLTableTheme("TableStyleLight3");
        public static readonly XLTableTheme TableStyleLight2 = new XLTableTheme("TableStyleLight2");
        public static readonly XLTableTheme TableStyleLight1 = new XLTableTheme("TableStyleLight1");
        public static readonly XLTableTheme TableStyleDark11 = new XLTableTheme("TableStyleDark11");
        public static readonly XLTableTheme TableStyleDark10 = new XLTableTheme("TableStyleDark10");
        public static readonly XLTableTheme TableStyleDark9 = new XLTableTheme("TableStyleDark9");
        public static readonly XLTableTheme TableStyleDark8 = new XLTableTheme("TableStyleDark8");
        public static readonly XLTableTheme TableStyleDark7 = new XLTableTheme("TableStyleDark7");
        public static readonly XLTableTheme TableStyleDark6 = new XLTableTheme("TableStyleDark6");
        public static readonly XLTableTheme TableStyleDark5 = new XLTableTheme("TableStyleDark5");
        public static readonly XLTableTheme TableStyleDark4 = new XLTableTheme("TableStyleDark4");
        public static readonly XLTableTheme TableStyleDark3 = new XLTableTheme("TableStyleDark3");
        public static readonly XLTableTheme TableStyleDark2 = new XLTableTheme("TableStyleDark2");
        public static readonly XLTableTheme TableStyleDark1 = new XLTableTheme("TableStyleDark1");

        public string Name { get; private set; }

        public XLTableTheme(string name)
        {
            this.Name = name;
        }

        private static IEnumerable<XLTableTheme> allThemes;

        public static IEnumerable<XLTableTheme> GetAllThemes()
        {
            return (allThemes ?? (allThemes = typeof(XLTableTheme).GetFields(BindingFlags.Static | BindingFlags.Public)
                .Where(fi => fi.FieldType.Equals(typeof(XLTableTheme)))
                .Select(fi => (XLTableTheme)fi.GetValue(null))
                .ToArray()));
        }

        public static XLTableTheme FromName(string name)
        {
            return GetAllThemes().FirstOrDefault(s => s.Name == name);
        }

        #region Overrides

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            XLTableTheme theme = obj as XLTableTheme;
            if (theme == null)
            {
                return false;
            }
            return this.Name.Equals(theme.Name);
        }

        public override int GetHashCode()
        {
            return this.Name.GetHashCode();
        }

        public override string ToString()
        {
            return this.Name;
        }

        #endregion Overrides
    }
}