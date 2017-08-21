#if _NETSTANDARD_
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing;

namespace ClosedXML.NetStandard
{
    public class XLColorTranslator
    {
        private static IDictionary<string, Color> NAMED_COLOR_MAP = new Dictionary<string, Color>{
             
            //X11 Color Names (W3C color names)
            //Red color names
            {"indianred", Color.FromArgb(205, 92, 92)},
            {"lightcoral",Color.FromArgb(240, 128, 128)},
            {"salmon",Color.FromArgb(250, 128, 114)},
            {"darksalmon",Color.FromArgb(233, 150, 122)},
            {"crimson",Color.FromArgb(220, 20, 60)},
            {"red",Color.FromArgb(255, 0, 0)},
            {"firebrick",Color.FromArgb(178, 34, 34)},
            {"darkred",Color.FromArgb(139, 0, 0)},
            //Pink color names
            {"pink",Color.FromArgb(255, 192, 203)},
            {"lightpink",Color.FromArgb(255, 182, 193)},
            {"hotpink",Color.FromArgb(255, 105, 180)},
            {"deeppink",Color.FromArgb(255, 20, 147)},
            {"mediumvioletred",Color.FromArgb(199, 21, 133)},
            {"palevioletred",Color.FromArgb(219, 112, 147)},
            //Orange color names
            {"lightsalmon",Color.FromArgb(255, 160, 122)},
            {"coral",Color.FromArgb(255, 127, 80)},
            {"tomato",Color.FromArgb(255, 99, 71)},
            {"orangered",Color.FromArgb(255, 69, 0)},
            {"darkorange",Color.FromArgb(255, 140, 0)},
            {"orange",Color.FromArgb(255, 165, 0)},
            //Yellow color names
            {"gold",Color.FromArgb(255, 215, 0)},
            {"yellow",Color.FromArgb(255, 255, 0)},
            {"lightyellow",Color.FromArgb(255, 255, 224)},
            {"lemonchiffon",Color.FromArgb(255, 250, 205)},
            {"lightgoldenrodyellow",Color.FromArgb(250, 250, 210)},
            {"papayawhip",Color.FromArgb(255, 239, 213)},
            {"moccasin",Color.FromArgb(255, 228, 181)},
            {"peachpuff",Color.FromArgb(255, 218, 185)},
            {"palegoldenrod",Color.FromArgb(238, 232, 170)},
            {"khaki",Color.FromArgb(240, 230, 140)},
            {"darkkhaki",Color.FromArgb(189, 183, 107)},
            //Purple color names
            {"lavender",Color.FromArgb(230, 230, 250)},
            {"thistle",Color.FromArgb(216, 191, 216)},
            {"plum",Color.FromArgb(221, 160, 221)},
            {"violet",Color.FromArgb(238, 130, 238)},
            {"orchid",Color.FromArgb(218, 112, 214)},
            {"fuchsia",Color.FromArgb(255, 0, 255)},
            {"magenta",Color.FromArgb(255, 0, 255)},
            {"mediumorchid",Color.FromArgb(186, 85, 211)},
            {"mediumpurple",Color.FromArgb(147, 112, 219)},
            {"blueviolet",Color.FromArgb(138, 43, 226)},
            {"darkviolet",Color.FromArgb(148, 0, 211)},
            {"darkorchid",Color.FromArgb(153, 50, 204)},
            {"darkmagenta",Color.FromArgb(139, 0, 139)},
            {"purple",Color.FromArgb(128, 0, 128)},
            {"indigo",Color.FromArgb(75, 0, 130)},
            {"slateblue",Color.FromArgb(106, 90, 205)},
            {"darkslateblue",Color.FromArgb(72, 61, 139)},
            //Green color names
            {"greenyellow",Color.FromArgb(173, 255, 47)},
            {"chartreuse",Color.FromArgb(127, 255, 0)},
            {"lawngreen",Color.FromArgb(124, 252, 0)},
            {"lime",Color.FromArgb(0, 255, 0)},
            {"limegreen",Color.FromArgb(50, 205, 50)},
            {"palegreen",Color.FromArgb(152, 251, 152)},
            {"lightgreen",Color.FromArgb(144, 238, 144)},
            {"mediumspringgreen",Color.FromArgb(0, 250, 154)},
            {"springgreen",Color.FromArgb(0, 255, 127)},
            {"mediumseagreen",Color.FromArgb(60, 179, 113)},
            {"seagreen",Color.FromArgb(46, 139, 87)},
            {"forestgreen",Color.FromArgb(34, 139, 34)},
            {"green",Color.FromArgb(0, 128, 0)},
            {"darkgreen",Color.FromArgb(0, 100, 0)},
            {"yellowgreen",Color.FromArgb(154, 205, 50)},
            {"olivedrab",Color.FromArgb(107, 142, 35)},
            {"olive",Color.FromArgb(128, 128, 0)},
            {"darkolivegreen",Color.FromArgb(85, 107, 47)},
            {"mediumaquamarine",Color.FromArgb(102, 205, 170)},
            {"darkseagreen",Color.FromArgb(143, 188, 143)},
            {"lightseagreen",Color.FromArgb(32, 178, 170)},
            {"darkcyan",Color.FromArgb(0, 139, 139)},
            {"teal",Color.FromArgb(0, 128, 128)},
            //Blue color names
            {"aqua",Color.FromArgb(0, 255, 255)},
            {"cyan",Color.FromArgb(0, 255, 255)},
            {"lightcyan",Color.FromArgb(224, 255, 255)},
            {"paleturquoise",Color.FromArgb(175, 238, 238)},
            {"aquamarine",Color.FromArgb(127, 255, 212)},
            {"turquoise",Color.FromArgb(64, 224, 208)},
            {"mediumturquoise",Color.FromArgb(72, 209, 204)},
            {"darkturquoise",Color.FromArgb(0, 206, 209)},
            {"cadetblue",Color.FromArgb(95, 158, 160)},
            {"steelblue",Color.FromArgb(70, 130, 180)},
            {"lightsteelblue",Color.FromArgb(176, 196, 222)},
            {"powderblue",Color.FromArgb(176, 224, 230)},
            {"lightblue",Color.FromArgb(173, 216, 230)},
            {"skyblue",Color.FromArgb(135, 206, 235)},
            {"lightskyblue",Color.FromArgb(135, 206, 250)},
            {"deepskyblue",Color.FromArgb(0, 191, 255)},
            {"dodgerblue",Color.FromArgb(30, 144, 255)},
            {"cornflowerblue",Color.FromArgb(100, 149, 237)},
            {"mediumslateblue",Color.FromArgb(123, 104, 238)},
            {"royalblue",Color.FromArgb(65, 105, 225)},
            {"blue",Color.FromArgb(0, 0, 255)},
            {"mediumblue",Color.FromArgb(0, 0, 205)},
            {"darkblue",Color.FromArgb(0, 0, 139)},
            {"navy",Color.FromArgb(0, 0, 128)},
            {"midnightblue",Color.FromArgb(25, 25, 112)},
            //Brown color names
            {"cornsilk",Color.FromArgb(255, 248, 220)},
            {"blanchedalmond",Color.FromArgb(255, 235, 205)},
            {"bisque",Color.FromArgb(255, 228, 196)},
            {"navajowhite",Color.FromArgb(255, 222, 173)},
            {"wheat",Color.FromArgb(245, 222, 179)},
            {"burlywood",Color.FromArgb(222, 184, 135)},
            {"tan",Color.FromArgb(210, 180, 140)},
            {"rosybrown",Color.FromArgb(188, 143, 143)},
            {"sandybrown",Color.FromArgb(244, 164, 96)},
            {"goldenrod",Color.FromArgb(218, 165, 32)},
            {"darkgoldenrod",Color.FromArgb(184, 134, 11)},
            {"peru",Color.FromArgb(205, 133, 63)},
            {"chocolate",Color.FromArgb(210, 105, 30)},
            {"saddlebrown",Color.FromArgb(139, 69, 19)},
            {"sienna",Color.FromArgb(160, 82, 45)},
            {"brown",Color.FromArgb(165, 42, 42)},
            {"maroon",Color.FromArgb(128, 0, 0)},
            //White color names
            {"white",Color.FromArgb(255, 255, 255)},
            {"snow",Color.FromArgb(255, 250, 250)},
            {"honeydew",Color.FromArgb(240, 255, 240)},
            {"mintcream",Color.FromArgb(245, 255, 250)},
            {"azure",Color.FromArgb(240, 255, 255)},
            {"aliceblue",Color.FromArgb(240, 248, 255)},
            {"ghostwhite",Color.FromArgb(248, 248, 255)},
            {"whitesmoke",Color.FromArgb(245, 245, 245)},
            {"seashell",Color.FromArgb(255, 245, 238)},
            {"beige",Color.FromArgb(245, 245, 220)},
            {"oldlace",Color.FromArgb(253, 245, 230)},
            {"floralwhite",Color.FromArgb(255, 250, 240)},
            {"ivory",Color.FromArgb(255, 255, 240)},
            {"antiquewhite",Color.FromArgb(250, 235, 215)},
            {"linen",Color.FromArgb(250, 240, 230)},
            {"lavenderblush",Color.FromArgb(255, 240, 245)},
            {"mistyrose",Color.FromArgb(255, 228, 225)},
            //Grey color names
            {"gainsboro",Color.FromArgb(220, 220, 220)},
            {"lightgrey",Color.FromArgb(211, 211, 211)},
            {"silver",Color.FromArgb(192, 192, 192)},
            {"darkgray",Color.FromArgb(169, 169, 169)},
            {"gray",Color.FromArgb(128, 128, 128)},
            {"dimgray",Color.FromArgb(105, 105, 105)},
            {"lightslategray",Color.FromArgb(119, 136, 153)},
            {"slategray",Color.FromArgb(112, 128, 144)},
            {"darkslategray",Color.FromArgb(47, 79, 79)},
            {"black",Color.FromArgb(0, 0, 0)},
        };

        private static Regex HexParser = new Regex("^#([0-9a-f]{2})?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$");
        private static Regex ShortHexParser = new Regex("^#([0-9a-f]{1})?([0-9a-f]{1})([0-9a-f]{1})([0-9a-f]{1})$");

        public static Color FromKnownColor(KnownColor knownColor)
        {
            var s = knownColor.ToString().ToLowerInvariant();
            return XLColorTranslator.FromHtml(s);
        }

        public static Color FromHtml(string htmlColor)
        {
            string id = htmlColor.Trim().ToLowerInvariant();
            if (NAMED_COLOR_MAP.ContainsKey(id))
            {
                return NAMED_COLOR_MAP[id];
            }
            else
            {
                var m = HexParser.Match(id);
                if (m.Value == String.Empty)
                {
                    m = ShortHexParser.Match(id);
                    if (m.Value == String.Empty)
                    {
                        throw new ArgumentException("Invalid HTML color: " + htmlColor);
                    }
                }
                if (String.IsNullOrWhiteSpace(m.Groups[1].Value))
                    return Color.FromArgb(
                        Convert.ToInt32(m.Groups[2].Value.PadRight(2, m.Groups[2].Value[0]), 16),
                        Convert.ToInt32(m.Groups[3].Value.PadRight(2, m.Groups[3].Value[0]), 16),
                        Convert.ToInt32(m.Groups[4].Value.PadRight(2, m.Groups[4].Value[0]), 16));
                else
                    return Color.FromArgb(
                        Convert.ToInt32(m.Groups[1].Value.PadRight(2, m.Groups[1].Value[0]), 16),
                        Convert.ToInt32(m.Groups[2].Value.PadRight(2, m.Groups[2].Value[0]), 16),
                        Convert.ToInt32(m.Groups[3].Value.PadRight(2, m.Groups[3].Value[0]), 16),
                        Convert.ToInt32(m.Groups[4].Value.PadRight(2, m.Groups[4].Value[0]), 16));
            }
        }
    }
}
#endif
