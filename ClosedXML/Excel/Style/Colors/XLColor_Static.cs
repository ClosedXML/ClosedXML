using ClosedXML.Excel.Caching;
using ClosedXML.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        private static readonly XLColorRepository Repository = new XLColorRepository(key => new XLColor(key));

        private static readonly Dictionary<Color, XLColor> ByColor = new Dictionary<Color, XLColor>();
        private static readonly Object ByColorLock = new Object();

        internal static XLColor FromKey(ref XLColorKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        public static XLColor FromColor(Color color)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Color,
                Color = color
            };
            return FromKey(ref key);
        }

        public static XLColor FromArgb(Int32 argb)
        {
            return FromColor(Color.FromArgb(argb));
        }

        public static XLColor FromArgb(Int32 r, Int32 g, Int32 b)
        {
            return FromColor(Color.FromArgb(r, g, b));
        }

        public static XLColor FromArgb(Int32 a, Int32 r, Int32 g, Int32 b)
        {
            return FromColor(Color.FromArgb(a, r, g, b));
        }

#if NETFRAMEWORK
        public static XLColor FromKnownColor(KnownColor color)
        {
            return FromColor(Color.FromKnownColor(color));
        }
#endif

        public static XLColor FromName(String name)
        {
            return FromColor(Color.FromName(name));
        }

        public static XLColor FromHtml(String htmlColor)
        {
            return FromColor(ColorStringParser.ParseFromHtml(htmlColor));
        }

        public static XLColor FromIndex(Int32 index)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Indexed,
                Indexed = index
            };
            return FromKey(ref key);
        }

        public static XLColor FromTheme(XLThemeColor themeColor)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Theme,
                ThemeColor = themeColor
            };
            return FromKey(ref key);
        }

        public static XLColor FromTheme(XLThemeColor themeColor, Double themeTint)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Theme,
                ThemeColor = themeColor,
                ThemeTint = themeTint
            };
            return FromKey(ref key);
        }

        private static Dictionary<Int32, XLColor> _indexedColors;

        public static Dictionary<Int32, XLColor> IndexedColors
        {
            get
            {
                if (_indexedColors == null)
                {
                    var retVal = new Dictionary<Int32, XLColor>
                        {
                            {0, FromHtml("#FF000000")},
                            {1, FromHtml("#FFFFFFFF")},
                            {2, FromHtml("#FFFF0000")},
                            {3, FromHtml("#FF00FF00")},
                            {4, FromHtml("#FF0000FF")},
                            {5, FromHtml("#FFFFFF00")},
                            {6, FromHtml("#FFFF00FF")},
                            {7, FromHtml("#FF00FFFF")},
                            {8, FromHtml("#FF000000")},
                            {9, FromHtml("#FFFFFFFF")},
                            {10, FromHtml("#FFFF0000")},
                            {11, FromHtml("#FF00FF00")},
                            {12, FromHtml("#FF0000FF")},
                            {13, FromHtml("#FFFFFF00")},
                            {14, FromHtml("#FFFF00FF")},
                            {15, FromHtml("#FF00FFFF")},
                            {16, FromHtml("#FF800000")},
                            {17, FromHtml("#FF008000")},
                            {18, FromHtml("#FF000080")},
                            {19, FromHtml("#FF808000")},
                            {20, FromHtml("#FF800080")},
                            {21, FromHtml("#FF008080")},
                            {22, FromHtml("#FFC0C0C0")},
                            {23, FromHtml("#FF808080")},
                            {24, FromHtml("#FF9999FF")},
                            {25, FromHtml("#FF993366")},
                            {26, FromHtml("#FFFFFFCC")},
                            {27, FromHtml("#FFCCFFFF")},
                            {28, FromHtml("#FF660066")},
                            {29, FromHtml("#FFFF8080")},
                            {30, FromHtml("#FF0066CC")},
                            {31, FromHtml("#FFCCCCFF")},
                            {32, FromHtml("#FF000080")},
                            {33, FromHtml("#FFFF00FF")},
                            {34, FromHtml("#FFFFFF00")},
                            {35, FromHtml("#FF00FFFF")},
                            {36, FromHtml("#FF800080")},
                            {37, FromHtml("#FF800000")},
                            {38, FromHtml("#FF008080")},
                            {39, FromHtml("#FF0000FF")},
                            {40, FromHtml("#FF00CCFF")},
                            {41, FromHtml("#FFCCFFFF")},
                            {42, FromHtml("#FFCCFFCC")},
                            {43, FromHtml("#FFFFFF99")},
                            {44, FromHtml("#FF99CCFF")},
                            {45, FromHtml("#FFFF99CC")},
                            {46, FromHtml("#FFCC99FF")},
                            {47, FromHtml("#FFFFCC99")},
                            {48, FromHtml("#FF3366FF")},
                            {49, FromHtml("#FF33CCCC")},
                            {50, FromHtml("#FF99CC00")},
                            {51, FromHtml("#FFFFCC00")},
                            {52, FromHtml("#FFFF9900")},
                            {53, FromHtml("#FFFF6600")},
                            {54, FromHtml("#FF666699")},
                            {55, FromHtml("#FF969696")},
                            {56, FromHtml("#FF003366")},
                            {57, FromHtml("#FF339966")},
                            {58, FromHtml("#FF003300")},
                            {59, FromHtml("#FF333300")},
                            {60, FromHtml("#FF993300")},
                            {61, FromHtml("#FF993366")},
                            {62, FromHtml("#FF333399")},
                            {63, FromHtml("#FF333333")},
                            {64, FromColor(Color.Transparent)}
                        };
                    _indexedColors = retVal;
                }
                return _indexedColors;
            }
        }

        internal static bool IsNullOrTransparent(XLColor color)
        {
            return color == null
                || !color.HasValue
                || IsTransparent(color.Key);
        }

        internal static bool IsTransparent(in XLColorKey colorKey)
        {
            return colorKey == NoColor.Key
                || (colorKey.ColorType == XLColorType.Indexed && colorKey.Indexed == 64);
        }

        public static XLColor NoColor { get; } = new XLColor();

        public static XLColor AliceBlue { get { return FromColor(Color.AliceBlue); } }

        public static XLColor AntiqueWhite { get { return FromColor(Color.AntiqueWhite); } }

        public static XLColor Aqua { get { return FromColor(Color.Aqua); } }

        public static XLColor Aquamarine { get { return FromColor(Color.Aquamarine); } }

        public static XLColor Azure { get { return FromColor(Color.Azure); } }

        public static XLColor Beige { get { return FromColor(Color.Beige); } }

        public static XLColor Bisque { get { return FromColor(Color.Bisque); } }

        public static XLColor Black { get { return FromColor(Color.Black); } }

        public static XLColor BlanchedAlmond { get { return FromColor(Color.BlanchedAlmond); } }

        public static XLColor Blue { get { return FromColor(Color.Blue); } }

        public static XLColor BlueViolet { get { return FromColor(Color.BlueViolet); } }

        public static XLColor Brown { get { return FromColor(Color.Brown); } }

        public static XLColor BurlyWood { get { return FromColor(Color.BurlyWood); } }

        public static XLColor CadetBlue { get { return FromColor(Color.CadetBlue); } }

        public static XLColor Chartreuse { get { return FromColor(Color.Chartreuse); } }

        public static XLColor Chocolate { get { return FromColor(Color.Chocolate); } }

        public static XLColor Coral { get { return FromColor(Color.Coral); } }

        public static XLColor CornflowerBlue { get { return FromColor(Color.CornflowerBlue); } }

        public static XLColor Cornsilk { get { return FromColor(Color.Cornsilk); } }

        public static XLColor Crimson { get { return FromColor(Color.Crimson); } }

        public static XLColor Cyan { get { return FromColor(Color.Cyan); } }

        public static XLColor DarkBlue { get { return FromColor(Color.DarkBlue); } }

        public static XLColor DarkCyan { get { return FromColor(Color.DarkCyan); } }

        public static XLColor DarkGoldenrod { get { return FromColor(Color.DarkGoldenrod); } }

        public static XLColor DarkGray { get { return FromColor(Color.DarkGray); } }

        public static XLColor DarkGreen { get { return FromColor(Color.DarkGreen); } }

        public static XLColor DarkKhaki { get { return FromColor(Color.DarkKhaki); } }

        public static XLColor DarkMagenta { get { return FromColor(Color.DarkMagenta); } }

        public static XLColor DarkOliveGreen { get { return FromColor(Color.DarkOliveGreen); } }

        public static XLColor DarkOrange { get { return FromColor(Color.DarkOrange); } }

        public static XLColor DarkOrchid { get { return FromColor(Color.DarkOrchid); } }

        public static XLColor DarkRed { get { return FromColor(Color.DarkRed); } }

        public static XLColor DarkSalmon { get { return FromColor(Color.DarkSalmon); } }

        public static XLColor DarkSeaGreen { get { return FromColor(Color.DarkSeaGreen); } }

        public static XLColor DarkSlateBlue { get { return FromColor(Color.DarkSlateBlue); } }

        public static XLColor DarkSlateGray { get { return FromColor(Color.DarkSlateGray); } }

        public static XLColor DarkTurquoise { get { return FromColor(Color.DarkTurquoise); } }

        public static XLColor DarkViolet { get { return FromColor(Color.DarkViolet); } }

        public static XLColor DeepPink { get { return FromColor(Color.DeepPink); } }

        public static XLColor DeepSkyBlue { get { return FromColor(Color.DeepSkyBlue); } }

        public static XLColor DimGray { get { return FromColor(Color.DimGray); } }

        public static XLColor DodgerBlue { get { return FromColor(Color.DodgerBlue); } }

        public static XLColor Firebrick { get { return FromColor(Color.Firebrick); } }

        public static XLColor FloralWhite { get { return FromColor(Color.FloralWhite); } }

        public static XLColor ForestGreen { get { return FromColor(Color.ForestGreen); } }

        public static XLColor Fuchsia { get { return FromColor(Color.Fuchsia); } }

        public static XLColor Gainsboro { get { return FromColor(Color.Gainsboro); } }

        public static XLColor GhostWhite { get { return FromColor(Color.GhostWhite); } }

        public static XLColor Gold { get { return FromColor(Color.Gold); } }

        public static XLColor Goldenrod { get { return FromColor(Color.Goldenrod); } }

        public static XLColor Gray { get { return FromColor(Color.Gray); } }

        public static XLColor Green { get { return FromColor(Color.Green); } }

        public static XLColor GreenYellow { get { return FromColor(Color.GreenYellow); } }

        public static XLColor Honeydew { get { return FromColor(Color.Honeydew); } }

        public static XLColor HotPink { get { return FromColor(Color.HotPink); } }

        public static XLColor IndianRed { get { return FromColor(Color.IndianRed); } }

        public static XLColor Indigo { get { return FromColor(Color.Indigo); } }

        public static XLColor Ivory { get { return FromColor(Color.Ivory); } }

        public static XLColor Khaki { get { return FromColor(Color.Khaki); } }

        public static XLColor Lavender { get { return FromColor(Color.Lavender); } }

        public static XLColor LavenderBlush { get { return FromColor(Color.LavenderBlush); } }

        public static XLColor LawnGreen { get { return FromColor(Color.LawnGreen); } }

        public static XLColor LemonChiffon { get { return FromColor(Color.LemonChiffon); } }

        public static XLColor LightBlue { get { return FromColor(Color.LightBlue); } }

        public static XLColor LightCoral { get { return FromColor(Color.LightCoral); } }

        public static XLColor LightCyan { get { return FromColor(Color.LightCyan); } }

        public static XLColor LightGoldenrodYellow { get { return FromColor(Color.LightGoldenrodYellow); } }

        public static XLColor LightGray { get { return FromColor(Color.LightGray); } }

        public static XLColor LightGreen { get { return FromColor(Color.LightGreen); } }

        public static XLColor LightPink { get { return FromColor(Color.LightPink); } }

        public static XLColor LightSalmon { get { return FromColor(Color.LightSalmon); } }

        public static XLColor LightSeaGreen { get { return FromColor(Color.LightSeaGreen); } }

        public static XLColor LightSkyBlue { get { return FromColor(Color.LightSkyBlue); } }

        public static XLColor LightSlateGray { get { return FromColor(Color.LightSlateGray); } }

        public static XLColor LightSteelBlue { get { return FromColor(Color.LightSteelBlue); } }

        public static XLColor LightYellow { get { return FromColor(Color.LightYellow); } }

        public static XLColor Lime { get { return FromColor(Color.Lime); } }

        public static XLColor LimeGreen { get { return FromColor(Color.LimeGreen); } }

        public static XLColor Linen { get { return FromColor(Color.Linen); } }

        public static XLColor Magenta { get { return FromColor(Color.Magenta); } }

        public static XLColor Maroon { get { return FromColor(Color.Maroon); } }

        public static XLColor MediumAquamarine { get { return FromColor(Color.MediumAquamarine); } }

        public static XLColor MediumBlue { get { return FromColor(Color.MediumBlue); } }

        public static XLColor MediumOrchid { get { return FromColor(Color.MediumOrchid); } }

        public static XLColor MediumPurple { get { return FromColor(Color.MediumPurple); } }

        public static XLColor MediumSeaGreen { get { return FromColor(Color.MediumSeaGreen); } }

        public static XLColor MediumSlateBlue { get { return FromColor(Color.MediumSlateBlue); } }

        public static XLColor MediumSpringGreen { get { return FromColor(Color.MediumSpringGreen); } }

        public static XLColor MediumTurquoise { get { return FromColor(Color.MediumTurquoise); } }

        public static XLColor MediumVioletRed { get { return FromColor(Color.MediumVioletRed); } }

        public static XLColor MidnightBlue { get { return FromColor(Color.MidnightBlue); } }

        public static XLColor MintCream { get { return FromColor(Color.MintCream); } }

        public static XLColor MistyRose { get { return FromColor(Color.MistyRose); } }

        public static XLColor Moccasin { get { return FromColor(Color.Moccasin); } }

        public static XLColor NavajoWhite { get { return FromColor(Color.NavajoWhite); } }

        public static XLColor Navy { get { return FromColor(Color.Navy); } }

        public static XLColor OldLace { get { return FromColor(Color.OldLace); } }

        public static XLColor Olive { get { return FromColor(Color.Olive); } }

        public static XLColor OliveDrab { get { return FromColor(Color.OliveDrab); } }

        public static XLColor Orange { get { return FromColor(Color.Orange); } }

        public static XLColor OrangeRed { get { return FromColor(Color.OrangeRed); } }

        public static XLColor Orchid { get { return FromColor(Color.Orchid); } }

        public static XLColor PaleGoldenrod { get { return FromColor(Color.PaleGoldenrod); } }

        public static XLColor PaleGreen { get { return FromColor(Color.PaleGreen); } }

        public static XLColor PaleTurquoise { get { return FromColor(Color.PaleTurquoise); } }

        public static XLColor PaleVioletRed { get { return FromColor(Color.PaleVioletRed); } }

        public static XLColor PapayaWhip { get { return FromColor(Color.PapayaWhip); } }

        public static XLColor PeachPuff { get { return FromColor(Color.PeachPuff); } }

        public static XLColor Peru { get { return FromColor(Color.Peru); } }

        public static XLColor Pink { get { return FromColor(Color.Pink); } }

        public static XLColor Plum { get { return FromColor(Color.Plum); } }

        public static XLColor PowderBlue { get { return FromColor(Color.PowderBlue); } }

        public static XLColor Purple { get { return FromColor(Color.Purple); } }

        public static XLColor Red { get { return FromColor(Color.Red); } }

        public static XLColor RosyBrown { get { return FromColor(Color.RosyBrown); } }

        public static XLColor RoyalBlue { get { return FromColor(Color.RoyalBlue); } }

        public static XLColor SaddleBrown { get { return FromColor(Color.SaddleBrown); } }

        public static XLColor Salmon { get { return FromColor(Color.Salmon); } }

        public static XLColor SandyBrown { get { return FromColor(Color.SandyBrown); } }

        public static XLColor SeaGreen { get { return FromColor(Color.SeaGreen); } }

        public static XLColor SeaShell { get { return FromColor(Color.SeaShell); } }

        public static XLColor Sienna { get { return FromColor(Color.Sienna); } }

        public static XLColor Silver { get { return FromColor(Color.Silver); } }

        public static XLColor SkyBlue { get { return FromColor(Color.SkyBlue); } }

        public static XLColor SlateBlue { get { return FromColor(Color.SlateBlue); } }

        public static XLColor SlateGray { get { return FromColor(Color.SlateGray); } }

        public static XLColor Snow { get { return FromColor(Color.Snow); } }

        public static XLColor SpringGreen { get { return FromColor(Color.SpringGreen); } }

        public static XLColor SteelBlue { get { return FromColor(Color.SteelBlue); } }

        public static XLColor Tan { get { return FromColor(Color.Tan); } }

        public static XLColor Teal { get { return FromColor(Color.Teal); } }

        public static XLColor Thistle { get { return FromColor(Color.Thistle); } }

        public static XLColor Tomato { get { return FromColor(Color.Tomato); } }

        public static XLColor Turquoise { get { return FromColor(Color.Turquoise); } }

        public static XLColor Violet { get { return FromColor(Color.Violet); } }

        public static XLColor Wheat { get { return FromColor(Color.Wheat); } }

        public static XLColor White { get { return FromColor(Color.White); } }

        public static XLColor WhiteSmoke { get { return FromColor(Color.WhiteSmoke); } }

        public static XLColor Yellow { get { return FromColor(Color.Yellow); } }

        public static XLColor YellowGreen { get { return FromColor(Color.YellowGreen); } }

        public static XLColor AirForceBlue { get { return FromHtml("#FF5D8AA8"); } }

        public static XLColor Alizarin { get { return FromHtml("#FFE32636"); } }

        public static XLColor Almond { get { return FromHtml("#FFEFDECD"); } }

        public static XLColor Amaranth { get { return FromHtml("#FFE52B50"); } }

        public static XLColor Amber { get { return FromHtml("#FFFFBF00"); } }

        public static XLColor AmberSaeEce { get { return FromHtml("#FFFF7E00"); } }

        public static XLColor AmericanRose { get { return FromHtml("#FFFF033E"); } }

        public static XLColor Amethyst { get { return FromHtml("#FF9966CC"); } }

        public static XLColor AntiFlashWhite { get { return FromHtml("#FFF2F3F4"); } }

        public static XLColor AntiqueBrass { get { return FromHtml("#FFCD9575"); } }

        public static XLColor AntiqueFuchsia { get { return FromHtml("#FF915C83"); } }

        public static XLColor AppleGreen { get { return FromHtml("#FF8DB600"); } }

        public static XLColor Apricot { get { return FromHtml("#FFFBCEB1"); } }

        public static XLColor Aquamarine1 { get { return FromHtml("#FF7FFFD0"); } }

        public static XLColor ArmyGreen { get { return FromHtml("#FF4B5320"); } }

        public static XLColor Arsenic { get { return FromHtml("#FF3B444B"); } }

        public static XLColor ArylideYellow { get { return FromHtml("#FFE9D66B"); } }

        public static XLColor AshGrey { get { return FromHtml("#FFB2BEB5"); } }

        public static XLColor Asparagus { get { return FromHtml("#FF87A96B"); } }

        public static XLColor AtomicTangerine { get { return FromHtml("#FFFF9966"); } }

        public static XLColor Auburn { get { return FromHtml("#FF6D351A"); } }

        public static XLColor Aureolin { get { return FromHtml("#FFFDEE00"); } }

        public static XLColor Aurometalsaurus { get { return FromHtml("#FF6E7F80"); } }

        public static XLColor Awesome { get { return FromHtml("#FFFF2052"); } }

        public static XLColor AzureColorWheel { get { return FromHtml("#FF007FFF"); } }

        public static XLColor BabyBlue { get { return FromHtml("#FF89CFF0"); } }

        public static XLColor BabyBlueEyes { get { return FromHtml("#FFA1CAF1"); } }

        public static XLColor BabyPink { get { return FromHtml("#FFF4C2C2"); } }

        public static XLColor BallBlue { get { return FromHtml("#FF21ABCD"); } }

        public static XLColor BananaMania { get { return FromHtml("#FFFAE7B5"); } }

        public static XLColor BattleshipGrey { get { return FromHtml("#FF848482"); } }

        public static XLColor Bazaar { get { return FromHtml("#FF98777B"); } }

        public static XLColor BeauBlue { get { return FromHtml("#FFBCD4E6"); } }

        public static XLColor Beaver { get { return FromHtml("#FF9F8170"); } }

        public static XLColor Bistre { get { return FromHtml("#FF3D2B1F"); } }

        public static XLColor Bittersweet { get { return FromHtml("#FFFE6F5E"); } }

        public static XLColor BleuDeFrance { get { return FromHtml("#FF318CE7"); } }

        public static XLColor BlizzardBlue { get { return FromHtml("#FFACE5EE"); } }

        public static XLColor Blond { get { return FromHtml("#FFFAF0BE"); } }

        public static XLColor BlueBell { get { return FromHtml("#FFA2A2D0"); } }

        public static XLColor BlueGray { get { return FromHtml("#FF6699CC"); } }

        public static XLColor BlueGreen { get { return FromHtml("#FF00DDDD"); } }

        public static XLColor BluePigment { get { return FromHtml("#FF333399"); } }

        public static XLColor BlueRyb { get { return FromHtml("#FF0247FE"); } }

        public static XLColor Blush { get { return FromHtml("#FFDE5D83"); } }

        public static XLColor Bole { get { return FromHtml("#FF79443B"); } }

        public static XLColor BondiBlue { get { return FromHtml("#FF0095B6"); } }

        public static XLColor BostonUniversityRed { get { return FromHtml("#FFCC0000"); } }

        public static XLColor BrandeisBlue { get { return FromHtml("#FF0070FF"); } }

        public static XLColor Brass { get { return FromHtml("#FFB5A642"); } }

        public static XLColor BrickRed { get { return FromHtml("#FFCB4154"); } }

        public static XLColor BrightCerulean { get { return FromHtml("#FF1DACD6"); } }

        public static XLColor BrightGreen { get { return FromHtml("#FF66FF00"); } }

        public static XLColor BrightLavender { get { return FromHtml("#FFBF94E4"); } }

        public static XLColor BrightMaroon { get { return FromHtml("#FFC32148"); } }

        public static XLColor BrightPink { get { return FromHtml("#FFFF007F"); } }

        public static XLColor BrightTurquoise { get { return FromHtml("#FF08E8DE"); } }

        public static XLColor BrightUbe { get { return FromHtml("#FFD19FE8"); } }

        public static XLColor BrilliantLavender { get { return FromHtml("#FFF4BBFF"); } }

        public static XLColor BrilliantRose { get { return FromHtml("#FFFF55A3"); } }

        public static XLColor BrinkPink { get { return FromHtml("#FFFB607F"); } }

        public static XLColor BritishRacingGreen { get { return FromHtml("#FF004225"); } }

        public static XLColor Bronze { get { return FromHtml("#FFCD7F32"); } }

        public static XLColor BrownTraditional { get { return FromHtml("#FF964B00"); } }

        public static XLColor BubbleGum { get { return FromHtml("#FFFFC1CC"); } }

        public static XLColor Bubbles { get { return FromHtml("#FFE7FEFF"); } }

        public static XLColor Buff { get { return FromHtml("#FFF0DC82"); } }

        public static XLColor BulgarianRose { get { return FromHtml("#FF480607"); } }

        public static XLColor Burgundy { get { return FromHtml("#FF800020"); } }

        public static XLColor BurntOrange { get { return FromHtml("#FFCC5500"); } }

        public static XLColor BurntSienna { get { return FromHtml("#FFE97451"); } }

        public static XLColor BurntUmber { get { return FromHtml("#FF8A3324"); } }

        public static XLColor Byzantine { get { return FromHtml("#FFBD33A4"); } }

        public static XLColor Byzantium { get { return FromHtml("#FF702963"); } }

        public static XLColor Cadet { get { return FromHtml("#FF536872"); } }

        public static XLColor CadetGrey { get { return FromHtml("#FF91A3B0"); } }

        public static XLColor CadmiumGreen { get { return FromHtml("#FF006B3C"); } }

        public static XLColor CadmiumOrange { get { return FromHtml("#FFED872D"); } }

        public static XLColor CadmiumRed { get { return FromHtml("#FFE30022"); } }

        public static XLColor CadmiumYellow { get { return FromHtml("#FFFFF600"); } }

        public static XLColor CalPolyPomonaGreen { get { return FromHtml("#FF1E4D2B"); } }

        public static XLColor CambridgeBlue { get { return FromHtml("#FFA3C1AD"); } }

        public static XLColor Camel { get { return FromHtml("#FFC19A6B"); } }

        public static XLColor CamouflageGreen { get { return FromHtml("#FF78866B"); } }

        public static XLColor CanaryYellow { get { return FromHtml("#FFFFEF00"); } }

        public static XLColor CandyAppleRed { get { return FromHtml("#FFFF0800"); } }

        public static XLColor CandyPink { get { return FromHtml("#FFE4717A"); } }

        public static XLColor CaputMortuum { get { return FromHtml("#FF592720"); } }

        public static XLColor Cardinal { get { return FromHtml("#FFC41E3A"); } }

        public static XLColor CaribbeanGreen { get { return FromHtml("#FF00CC99"); } }

        public static XLColor Carmine { get { return FromHtml("#FF960018"); } }

        public static XLColor CarminePink { get { return FromHtml("#FFEB4C42"); } }

        public static XLColor CarmineRed { get { return FromHtml("#FFFF0038"); } }

        public static XLColor CarnationPink { get { return FromHtml("#FFFFA6C9"); } }

        public static XLColor Carnelian { get { return FromHtml("#FFB31B1B"); } }

        public static XLColor CarolinaBlue { get { return FromHtml("#FF99BADD"); } }

        public static XLColor CarrotOrange { get { return FromHtml("#FFED9121"); } }

        public static XLColor Ceil { get { return FromHtml("#FF92A1CF"); } }

        public static XLColor Celadon { get { return FromHtml("#FFACE1AF"); } }

        public static XLColor CelestialBlue { get { return FromHtml("#FF4997D0"); } }

        public static XLColor Cerise { get { return FromHtml("#FFDE3163"); } }

        public static XLColor CerisePink { get { return FromHtml("#FFEC3B83"); } }

        public static XLColor Cerulean { get { return FromHtml("#FF007BA7"); } }

        public static XLColor CeruleanBlue { get { return FromHtml("#FF2A52BE"); } }

        public static XLColor Chamoisee { get { return FromHtml("#FFA0785A"); } }

        public static XLColor Champagne { get { return FromHtml("#FFF7E7CE"); } }

        public static XLColor Charcoal { get { return FromHtml("#FF36454F"); } }

        public static XLColor ChartreuseTraditional { get { return FromHtml("#FFDFFF00"); } }

        public static XLColor CherryBlossomPink { get { return FromHtml("#FFFFB7C5"); } }

        public static XLColor Chocolate1 { get { return FromHtml("#FF7B3F00"); } }

        public static XLColor ChromeYellow { get { return FromHtml("#FFFFA700"); } }

        public static XLColor Cinereous { get { return FromHtml("#FF98817B"); } }

        public static XLColor Cinnabar { get { return FromHtml("#FFE34234"); } }

        public static XLColor Citrine { get { return FromHtml("#FFE4D00A"); } }

        public static XLColor ClassicRose { get { return FromHtml("#FFFBCCE7"); } }

        public static XLColor Cobalt { get { return FromHtml("#FF0047AB"); } }

        public static XLColor ColumbiaBlue { get { return FromHtml("#FF9BDDFF"); } }

        public static XLColor CoolBlack { get { return FromHtml("#FF002E63"); } }

        public static XLColor CoolGrey { get { return FromHtml("#FF8C92AC"); } }

        public static XLColor Copper { get { return FromHtml("#FFB87333"); } }

        public static XLColor CopperRose { get { return FromHtml("#FF996666"); } }

        public static XLColor Coquelicot { get { return FromHtml("#FFFF3800"); } }

        public static XLColor CoralPink { get { return FromHtml("#FFF88379"); } }

        public static XLColor CoralRed { get { return FromHtml("#FFFF4040"); } }

        public static XLColor Cordovan { get { return FromHtml("#FF893F45"); } }

        public static XLColor Corn { get { return FromHtml("#FFFBEC5D"); } }

        public static XLColor CornellRed { get { return FromHtml("#FFB31B1B"); } }

        public static XLColor CosmicLatte { get { return FromHtml("#FFFFF8E7"); } }

        public static XLColor CottonCandy { get { return FromHtml("#FFFFBCD9"); } }

        public static XLColor Cream { get { return FromHtml("#FFFFFDD0"); } }

        public static XLColor CrimsonGlory { get { return FromHtml("#FFBE0032"); } }

        public static XLColor CyanProcess { get { return FromHtml("#FF00B7EB"); } }

        public static XLColor Daffodil { get { return FromHtml("#FFFFFF31"); } }

        public static XLColor Dandelion { get { return FromHtml("#FFF0E130"); } }

        public static XLColor DarkBrown { get { return FromHtml("#FF654321"); } }

        public static XLColor DarkByzantium { get { return FromHtml("#FF5D3954"); } }

        public static XLColor DarkCandyAppleRed { get { return FromHtml("#FFA40000"); } }

        public static XLColor DarkCerulean { get { return FromHtml("#FF08457E"); } }

        public static XLColor DarkChampagne { get { return FromHtml("#FFC2B280"); } }

        public static XLColor DarkChestnut { get { return FromHtml("#FF986960"); } }

        public static XLColor DarkCoral { get { return FromHtml("#FFCD5B45"); } }

        public static XLColor DarkElectricBlue { get { return FromHtml("#FF536878"); } }

        public static XLColor DarkGreen1 { get { return FromHtml("#FF013220"); } }

        public static XLColor DarkJungleGreen { get { return FromHtml("#FF1A2421"); } }

        public static XLColor DarkLava { get { return FromHtml("#FF483C32"); } }

        public static XLColor DarkLavender { get { return FromHtml("#FF734F96"); } }

        public static XLColor DarkMidnightBlue { get { return FromHtml("#FF003366"); } }

        public static XLColor DarkPastelBlue { get { return FromHtml("#FF779ECB"); } }

        public static XLColor DarkPastelGreen { get { return FromHtml("#FF03C03C"); } }

        public static XLColor DarkPastelPurple { get { return FromHtml("#FF966FD6"); } }

        public static XLColor DarkPastelRed { get { return FromHtml("#FFC23B22"); } }

        public static XLColor DarkPink { get { return FromHtml("#FFE75480"); } }

        public static XLColor DarkPowderBlue { get { return FromHtml("#FF003399"); } }

        public static XLColor DarkRaspberry { get { return FromHtml("#FF872657"); } }

        public static XLColor DarkScarlet { get { return FromHtml("#FF560319"); } }

        public static XLColor DarkSienna { get { return FromHtml("#FF3C1414"); } }

        public static XLColor DarkSpringGreen { get { return FromHtml("#FF177245"); } }

        public static XLColor DarkTan { get { return FromHtml("#FF918151"); } }

        public static XLColor DarkTangerine { get { return FromHtml("#FFFFA812"); } }

        public static XLColor DarkTaupe { get { return FromHtml("#FF483C32"); } }

        public static XLColor DarkTerraCotta { get { return FromHtml("#FFCC4E5C"); } }

        public static XLColor DartmouthGreen { get { return FromHtml("#FF00693E"); } }

        public static XLColor DavysGrey { get { return FromHtml("#FF555555"); } }

        public static XLColor DebianRed { get { return FromHtml("#FFD70A53"); } }

        public static XLColor DeepCarmine { get { return FromHtml("#FFA9203E"); } }

        public static XLColor DeepCarminePink { get { return FromHtml("#FFEF3038"); } }

        public static XLColor DeepCarrotOrange { get { return FromHtml("#FFE9692C"); } }

        public static XLColor DeepCerise { get { return FromHtml("#FFDA3287"); } }

        public static XLColor DeepChampagne { get { return FromHtml("#FFFAD6A5"); } }

        public static XLColor DeepChestnut { get { return FromHtml("#FFB94E48"); } }

        public static XLColor DeepFuchsia { get { return FromHtml("#FFC154C1"); } }

        public static XLColor DeepJungleGreen { get { return FromHtml("#FF004B49"); } }

        public static XLColor DeepLilac { get { return FromHtml("#FF9955BB"); } }

        public static XLColor DeepMagenta { get { return FromHtml("#FFCC00CC"); } }

        public static XLColor DeepPeach { get { return FromHtml("#FFFFCBA4"); } }

        public static XLColor DeepSaffron { get { return FromHtml("#FFFF9933"); } }

        public static XLColor Denim { get { return FromHtml("#FF1560BD"); } }

        public static XLColor Desert { get { return FromHtml("#FFC19A6B"); } }

        public static XLColor DesertSand { get { return FromHtml("#FFEDC9AF"); } }

        public static XLColor DogwoodRose { get { return FromHtml("#FFD71868"); } }

        public static XLColor DollarBill { get { return FromHtml("#FF85BB65"); } }

        public static XLColor Drab { get { return FromHtml("#FF967117"); } }

        public static XLColor DukeBlue { get { return FromHtml("#FF00009C"); } }

        public static XLColor EarthYellow { get { return FromHtml("#FFE1A95F"); } }

        public static XLColor Ecru { get { return FromHtml("#FFC2B280"); } }

        public static XLColor Eggplant { get { return FromHtml("#FF614051"); } }

        public static XLColor Eggshell { get { return FromHtml("#FFF0EAD6"); } }

        public static XLColor EgyptianBlue { get { return FromHtml("#FF1034A6"); } }

        public static XLColor ElectricBlue { get { return FromHtml("#FF7DF9FF"); } }

        public static XLColor ElectricCrimson { get { return FromHtml("#FFFF003F"); } }

        public static XLColor ElectricIndigo { get { return FromHtml("#FF6F00FF"); } }

        public static XLColor ElectricLavender { get { return FromHtml("#FFF4BBFF"); } }

        public static XLColor ElectricLime { get { return FromHtml("#FFCCFF00"); } }

        public static XLColor ElectricPurple { get { return FromHtml("#FFBF00FF"); } }

        public static XLColor ElectricUltramarine { get { return FromHtml("#FF3F00FF"); } }

        public static XLColor ElectricViolet { get { return FromHtml("#FF8F00FF"); } }

        public static XLColor Emerald { get { return FromHtml("#FF50C878"); } }

        public static XLColor EtonBlue { get { return FromHtml("#FF96C8A2"); } }

        public static XLColor Fallow { get { return FromHtml("#FFC19A6B"); } }

        public static XLColor FaluRed { get { return FromHtml("#FF801818"); } }

        public static XLColor Fandango { get { return FromHtml("#FFB53389"); } }

        public static XLColor FashionFuchsia { get { return FromHtml("#FFF400A1"); } }

        public static XLColor Fawn { get { return FromHtml("#FFE5AA70"); } }

        public static XLColor Feldgrau { get { return FromHtml("#FF4D5D53"); } }

        public static XLColor FernGreen { get { return FromHtml("#FF4F7942"); } }

        public static XLColor FerrariRed { get { return FromHtml("#FFFF2800"); } }

        public static XLColor FieldDrab { get { return FromHtml("#FF6C541E"); } }

        public static XLColor FireEngineRed { get { return FromHtml("#FFCE2029"); } }

        public static XLColor Flame { get { return FromHtml("#FFE25822"); } }

        public static XLColor FlamingoPink { get { return FromHtml("#FFFC8EAC"); } }

        public static XLColor Flavescent { get { return FromHtml("#FFF7E98E"); } }

        public static XLColor Flax { get { return FromHtml("#FFEEDC82"); } }

        public static XLColor FluorescentOrange { get { return FromHtml("#FFFFBF00"); } }

        public static XLColor FluorescentYellow { get { return FromHtml("#FFCCFF00"); } }

        public static XLColor Folly { get { return FromHtml("#FFFF004F"); } }

        public static XLColor ForestGreenTraditional { get { return FromHtml("#FF014421"); } }

        public static XLColor FrenchBeige { get { return FromHtml("#FFA67B5B"); } }

        public static XLColor FrenchBlue { get { return FromHtml("#FF0072BB"); } }

        public static XLColor FrenchLilac { get { return FromHtml("#FF86608E"); } }

        public static XLColor FrenchRose { get { return FromHtml("#FFF64A8A"); } }

        public static XLColor FuchsiaPink { get { return FromHtml("#FFFF77FF"); } }

        public static XLColor Fulvous { get { return FromHtml("#FFE48400"); } }

        public static XLColor FuzzyWuzzy { get { return FromHtml("#FFCC6666"); } }

        public static XLColor Gamboge { get { return FromHtml("#FFE49B0F"); } }

        public static XLColor Ginger { get { return FromHtml("#FFF9F9FF"); } }

        public static XLColor Glaucous { get { return FromHtml("#FF6082B6"); } }

        public static XLColor GoldenBrown { get { return FromHtml("#FF996515"); } }

        public static XLColor GoldenPoppy { get { return FromHtml("#FFFCC200"); } }

        public static XLColor GoldenYellow { get { return FromHtml("#FFFFDF00"); } }

        public static XLColor GoldMetallic { get { return FromHtml("#FFD4AF37"); } }

        public static XLColor GrannySmithApple { get { return FromHtml("#FFA8E4A0"); } }

        public static XLColor GrayAsparagus { get { return FromHtml("#FF465945"); } }

        public static XLColor GreenPigment { get { return FromHtml("#FF00A550"); } }

        public static XLColor GreenRyb { get { return FromHtml("#FF66B032"); } }

        public static XLColor Grullo { get { return FromHtml("#FFA99A86"); } }

        public static XLColor HalayaUbe { get { return FromHtml("#FF663854"); } }

        public static XLColor HanBlue { get { return FromHtml("#FF446CCF"); } }

        public static XLColor HanPurple { get { return FromHtml("#FF5218FA"); } }

        public static XLColor HansaYellow { get { return FromHtml("#FFE9D66B"); } }

        public static XLColor Harlequin { get { return FromHtml("#FF3FFF00"); } }

        public static XLColor HarvardCrimson { get { return FromHtml("#FFC90016"); } }

        public static XLColor HarvestGold { get { return FromHtml("#FFDA9100"); } }

        public static XLColor Heliotrope { get { return FromHtml("#FFDF73FF"); } }

        public static XLColor HollywoodCerise { get { return FromHtml("#FFF400A1"); } }

        public static XLColor HookersGreen { get { return FromHtml("#FF007000"); } }

        public static XLColor HotMagenta { get { return FromHtml("#FFFF1DCE"); } }

        public static XLColor HunterGreen { get { return FromHtml("#FF355E3B"); } }

        public static XLColor Iceberg { get { return FromHtml("#FF71A6D2"); } }

        public static XLColor Icterine { get { return FromHtml("#FFFCF75E"); } }

        public static XLColor Inchworm { get { return FromHtml("#FFB2EC5D"); } }

        public static XLColor IndiaGreen { get { return FromHtml("#FF138808"); } }

        public static XLColor IndianYellow { get { return FromHtml("#FFE3A857"); } }

        public static XLColor IndigoDye { get { return FromHtml("#FF00416A"); } }

        public static XLColor InternationalKleinBlue { get { return FromHtml("#FF002FA7"); } }

        public static XLColor InternationalOrange { get { return FromHtml("#FFFF4F00"); } }

        public static XLColor Iris { get { return FromHtml("#FF5A4FCF"); } }

        public static XLColor Isabelline { get { return FromHtml("#FFF4F0EC"); } }

        public static XLColor IslamicGreen { get { return FromHtml("#FF009000"); } }

        public static XLColor Jade { get { return FromHtml("#FF00A86B"); } }

        public static XLColor Jasper { get { return FromHtml("#FFD73B3E"); } }

        public static XLColor JazzberryJam { get { return FromHtml("#FFA50B5E"); } }

        public static XLColor Jonquil { get { return FromHtml("#FFFADA5E"); } }

        public static XLColor JuneBud { get { return FromHtml("#FFBDDA57"); } }

        public static XLColor JungleGreen { get { return FromHtml("#FF29AB87"); } }

        public static XLColor KellyGreen { get { return FromHtml("#FF4CBB17"); } }

        public static XLColor KhakiHtmlCssKhaki { get { return FromHtml("#FFC3B091"); } }

        public static XLColor LanguidLavender { get { return FromHtml("#FFD6CADD"); } }

        public static XLColor LapisLazuli { get { return FromHtml("#FF26619C"); } }

        public static XLColor LaSalleGreen { get { return FromHtml("#FF087830"); } }

        public static XLColor LaserLemon { get { return FromHtml("#FFFEFE22"); } }

        public static XLColor Lava { get { return FromHtml("#FFCF1020"); } }

        public static XLColor LavenderBlue { get { return FromHtml("#FFCCCCFF"); } }

        public static XLColor LavenderFloral { get { return FromHtml("#FFB57EDC"); } }

        public static XLColor LavenderGray { get { return FromHtml("#FFC4C3D0"); } }

        public static XLColor LavenderIndigo { get { return FromHtml("#FF9457EB"); } }

        public static XLColor LavenderPink { get { return FromHtml("#FFFBAED2"); } }

        public static XLColor LavenderPurple { get { return FromHtml("#FF967BB6"); } }

        public static XLColor LavenderRose { get { return FromHtml("#FFFBA0E3"); } }

        public static XLColor Lemon { get { return FromHtml("#FFFFF700"); } }

        public static XLColor LightApricot { get { return FromHtml("#FFFDD5B1"); } }

        public static XLColor LightBrown { get { return FromHtml("#FFB5651D"); } }

        public static XLColor LightCarminePink { get { return FromHtml("#FFE66771"); } }

        public static XLColor LightCornflowerBlue { get { return FromHtml("#FF93CCEA"); } }

        public static XLColor LightFuchsiaPink { get { return FromHtml("#FFF984EF"); } }

        public static XLColor LightMauve { get { return FromHtml("#FFDCD0FF"); } }

        public static XLColor LightPastelPurple { get { return FromHtml("#FFB19CD9"); } }

        public static XLColor LightSalmonPink { get { return FromHtml("#FFFF9999"); } }

        public static XLColor LightTaupe { get { return FromHtml("#FFB38B6D"); } }

        public static XLColor LightThulianPink { get { return FromHtml("#FFE68FAC"); } }

        public static XLColor LightYellow1 { get { return FromHtml("#FFFFFFED"); } }

        public static XLColor Lilac { get { return FromHtml("#FFC8A2C8"); } }

        public static XLColor LimeColorWheel { get { return FromHtml("#FFBFFF00"); } }

        public static XLColor LincolnGreen { get { return FromHtml("#FF195905"); } }

        public static XLColor Liver { get { return FromHtml("#FF534B4F"); } }

        public static XLColor Lust { get { return FromHtml("#FFE62020"); } }

        public static XLColor MacaroniAndCheese { get { return FromHtml("#FFFFBD88"); } }

        public static XLColor MagentaDye { get { return FromHtml("#FFCA1F7B"); } }

        public static XLColor MagentaProcess { get { return FromHtml("#FFFF0090"); } }

        public static XLColor MagicMint { get { return FromHtml("#FFAAF0D1"); } }

        public static XLColor Magnolia { get { return FromHtml("#FFF8F4FF"); } }

        public static XLColor Mahogany { get { return FromHtml("#FFC04000"); } }

        public static XLColor Maize { get { return FromHtml("#FFFBEC5D"); } }

        public static XLColor MajorelleBlue { get { return FromHtml("#FF6050DC"); } }

        public static XLColor Malachite { get { return FromHtml("#FF0BDA51"); } }

        public static XLColor Manatee { get { return FromHtml("#FF979AAA"); } }

        public static XLColor MangoTango { get { return FromHtml("#FFFF8243"); } }

        public static XLColor MaroonX11 { get { return FromHtml("#FFB03060"); } }

        public static XLColor Mauve { get { return FromHtml("#FFE0B0FF"); } }

        public static XLColor Mauvelous { get { return FromHtml("#FFEF98AA"); } }

        public static XLColor MauveTaupe { get { return FromHtml("#FF915F6D"); } }

        public static XLColor MayaBlue { get { return FromHtml("#FF73C2FB"); } }

        public static XLColor MeatBrown { get { return FromHtml("#FFE5B73B"); } }

        public static XLColor MediumAquamarine1 { get { return FromHtml("#FF66DDAA"); } }

        public static XLColor MediumCandyAppleRed { get { return FromHtml("#FFE2062C"); } }

        public static XLColor MediumCarmine { get { return FromHtml("#FFAF4035"); } }

        public static XLColor MediumChampagne { get { return FromHtml("#FFF3E5AB"); } }

        public static XLColor MediumElectricBlue { get { return FromHtml("#FF035096"); } }

        public static XLColor MediumJungleGreen { get { return FromHtml("#FF1C352D"); } }

        public static XLColor MediumPersianBlue { get { return FromHtml("#FF0067A5"); } }

        public static XLColor MediumRedViolet { get { return FromHtml("#FFBB3385"); } }

        public static XLColor MediumSpringBud { get { return FromHtml("#FFC9DC87"); } }

        public static XLColor MediumTaupe { get { return FromHtml("#FF674C47"); } }

        public static XLColor Melon { get { return FromHtml("#FFFDBCB4"); } }

        public static XLColor MidnightGreenEagleGreen { get { return FromHtml("#FF004953"); } }

        public static XLColor MikadoYellow { get { return FromHtml("#FFFFC40C"); } }

        public static XLColor Mint { get { return FromHtml("#FF3EB489"); } }

        public static XLColor MintGreen { get { return FromHtml("#FF98FF98"); } }

        public static XLColor ModeBeige { get { return FromHtml("#FF967117"); } }

        public static XLColor MoonstoneBlue { get { return FromHtml("#FF73A9C2"); } }

        public static XLColor MordantRed19 { get { return FromHtml("#FFAE0C00"); } }

        public static XLColor MossGreen { get { return FromHtml("#FFADDFAD"); } }

        public static XLColor MountainMeadow { get { return FromHtml("#FF30BA8F"); } }

        public static XLColor MountbattenPink { get { return FromHtml("#FF997A8D"); } }

        public static XLColor MsuGreen { get { return FromHtml("#FF18453B"); } }

        public static XLColor Mulberry { get { return FromHtml("#FFC54B8C"); } }

        public static XLColor Mustard { get { return FromHtml("#FFFFDB58"); } }

        public static XLColor Myrtle { get { return FromHtml("#FF21421E"); } }

        public static XLColor NadeshikoPink { get { return FromHtml("#FFF6ADC6"); } }

        public static XLColor NapierGreen { get { return FromHtml("#FF2A8000"); } }

        public static XLColor NaplesYellow { get { return FromHtml("#FFFADA5E"); } }

        public static XLColor NeonCarrot { get { return FromHtml("#FFFFA343"); } }

        public static XLColor NeonFuchsia { get { return FromHtml("#FFFE59C2"); } }

        public static XLColor NeonGreen { get { return FromHtml("#FF39FF14"); } }

        public static XLColor NonPhotoBlue { get { return FromHtml("#FFA4DDED"); } }

        public static XLColor OceanBoatBlue { get { return FromHtml("#FFCC7422"); } }

        public static XLColor Ochre { get { return FromHtml("#FFCC7722"); } }

        public static XLColor OldGold { get { return FromHtml("#FFCFB53B"); } }

        public static XLColor OldLavender { get { return FromHtml("#FF796878"); } }

        public static XLColor OldMauve { get { return FromHtml("#FF673147"); } }

        public static XLColor OldRose { get { return FromHtml("#FFC08081"); } }

        public static XLColor OliveDrab7 { get { return FromHtml("#FF3C341F"); } }

        public static XLColor Olivine { get { return FromHtml("#FF9AB973"); } }

        public static XLColor Onyx { get { return FromHtml("#FF0F0F0F"); } }

        public static XLColor OperaMauve { get { return FromHtml("#FFB784A7"); } }

        public static XLColor OrangeColorWheel { get { return FromHtml("#FFFF7F00"); } }

        public static XLColor OrangePeel { get { return FromHtml("#FFFF9F00"); } }

        public static XLColor OrangeRyb { get { return FromHtml("#FFFB9902"); } }

        public static XLColor OtterBrown { get { return FromHtml("#FF654321"); } }

        public static XLColor OuCrimsonRed { get { return FromHtml("#FF990000"); } }

        public static XLColor OuterSpace { get { return FromHtml("#FF414A4C"); } }

        public static XLColor OutrageousOrange { get { return FromHtml("#FFFF6E4A"); } }

        public static XLColor OxfordBlue { get { return FromHtml("#FF002147"); } }

        public static XLColor PakistanGreen { get { return FromHtml("#FF00421B"); } }

        public static XLColor PalatinateBlue { get { return FromHtml("#FF273BE2"); } }

        public static XLColor PalatinatePurple { get { return FromHtml("#FF682860"); } }

        public static XLColor PaleAqua { get { return FromHtml("#FFBCD4E6"); } }

        public static XLColor PaleBrown { get { return FromHtml("#FF987654"); } }

        public static XLColor PaleCarmine { get { return FromHtml("#FFAF4035"); } }

        public static XLColor PaleCerulean { get { return FromHtml("#FF9BC4E2"); } }

        public static XLColor PaleChestnut { get { return FromHtml("#FFDDADAF"); } }

        public static XLColor PaleCopper { get { return FromHtml("#FFDA8A67"); } }

        public static XLColor PaleCornflowerBlue { get { return FromHtml("#FFABCDEF"); } }

        public static XLColor PaleGold { get { return FromHtml("#FFE6BE8A"); } }

        public static XLColor PaleMagenta { get { return FromHtml("#FFF984E5"); } }

        public static XLColor PalePink { get { return FromHtml("#FFFADADD"); } }

        public static XLColor PaleRobinEggBlue { get { return FromHtml("#FF96DED1"); } }

        public static XLColor PaleSilver { get { return FromHtml("#FFC9C0BB"); } }

        public static XLColor PaleSpringBud { get { return FromHtml("#FFECEBBD"); } }

        public static XLColor PaleTaupe { get { return FromHtml("#FFBC987E"); } }

        public static XLColor PansyPurple { get { return FromHtml("#FF78184A"); } }

        public static XLColor ParisGreen { get { return FromHtml("#FF50C878"); } }

        public static XLColor PastelBlue { get { return FromHtml("#FFAEC6CF"); } }

        public static XLColor PastelBrown { get { return FromHtml("#FF836953"); } }

        public static XLColor PastelGray { get { return FromHtml("#FFCFCFC4"); } }

        public static XLColor PastelGreen { get { return FromHtml("#FF77DD77"); } }

        public static XLColor PastelMagenta { get { return FromHtml("#FFF49AC2"); } }

        public static XLColor PastelOrange { get { return FromHtml("#FFFFB347"); } }

        public static XLColor PastelPink { get { return FromHtml("#FFFFD1DC"); } }

        public static XLColor PastelPurple { get { return FromHtml("#FFB39EB5"); } }

        public static XLColor PastelRed { get { return FromHtml("#FFFF6961"); } }

        public static XLColor PastelViolet { get { return FromHtml("#FFCB99C9"); } }

        public static XLColor PastelYellow { get { return FromHtml("#FFFDFD96"); } }

        public static XLColor PaynesGrey { get { return FromHtml("#FF40404F"); } }

        public static XLColor Peach { get { return FromHtml("#FFFFE5B4"); } }

        public static XLColor PeachOrange { get { return FromHtml("#FFFFCC99"); } }

        public static XLColor PeachYellow { get { return FromHtml("#FFFADFAD"); } }

        public static XLColor Pear { get { return FromHtml("#FFD1E231"); } }

        public static XLColor Pearl { get { return FromHtml("#FFF0EAD6"); } }

        public static XLColor Peridot { get { return FromHtml("#FFE6E200"); } }

        public static XLColor Periwinkle { get { return FromHtml("#FFCCCCFF"); } }

        public static XLColor PersianBlue { get { return FromHtml("#FF1C39BB"); } }

        public static XLColor PersianGreen { get { return FromHtml("#FF00A693"); } }

        public static XLColor PersianIndigo { get { return FromHtml("#FF32127A"); } }

        public static XLColor PersianOrange { get { return FromHtml("#FFD99058"); } }

        public static XLColor PersianPink { get { return FromHtml("#FFF77FBE"); } }

        public static XLColor PersianPlum { get { return FromHtml("#FF701C1C"); } }

        public static XLColor PersianRed { get { return FromHtml("#FFCC3333"); } }

        public static XLColor PersianRose { get { return FromHtml("#FFFE28A2"); } }

        public static XLColor Persimmon { get { return FromHtml("#FFEC5800"); } }

        public static XLColor Phlox { get { return FromHtml("#FFDF00FF"); } }

        public static XLColor PhthaloBlue { get { return FromHtml("#FF000F89"); } }

        public static XLColor PhthaloGreen { get { return FromHtml("#FF123524"); } }

        public static XLColor PiggyPink { get { return FromHtml("#FFFDDDE6"); } }

        public static XLColor PineGreen { get { return FromHtml("#FF01796F"); } }

        public static XLColor PinkOrange { get { return FromHtml("#FFFF9966"); } }

        public static XLColor PinkPearl { get { return FromHtml("#FFE7ACCF"); } }

        public static XLColor PinkSherbet { get { return FromHtml("#FFF78FA7"); } }

        public static XLColor Pistachio { get { return FromHtml("#FF93C572"); } }

        public static XLColor Platinum { get { return FromHtml("#FFE5E4E2"); } }

        public static XLColor PlumTraditional { get { return FromHtml("#FF8E4585"); } }

        public static XLColor PortlandOrange { get { return FromHtml("#FFFF5A36"); } }

        public static XLColor PrincetonOrange { get { return FromHtml("#FFFF8F00"); } }

        public static XLColor Prune { get { return FromHtml("#FF701C1C"); } }

        public static XLColor PrussianBlue { get { return FromHtml("#FF003153"); } }

        public static XLColor PsychedelicPurple { get { return FromHtml("#FFDF00FF"); } }

        public static XLColor Puce { get { return FromHtml("#FFCC8899"); } }

        public static XLColor Pumpkin { get { return FromHtml("#FFFF7518"); } }

        public static XLColor PurpleHeart { get { return FromHtml("#FF69359C"); } }

        public static XLColor PurpleMountainMajesty { get { return FromHtml("#FF9678B6"); } }

        public static XLColor PurpleMunsell { get { return FromHtml("#FF9F00C5"); } }

        public static XLColor PurplePizzazz { get { return FromHtml("#FFFE4EDA"); } }

        public static XLColor PurpleTaupe { get { return FromHtml("#FF50404D"); } }

        public static XLColor PurpleX11 { get { return FromHtml("#FFA020F0"); } }

        public static XLColor RadicalRed { get { return FromHtml("#FFFF355E"); } }

        public static XLColor Raspberry { get { return FromHtml("#FFE30B5D"); } }

        public static XLColor RaspberryGlace { get { return FromHtml("#FF915F6D"); } }

        public static XLColor RaspberryPink { get { return FromHtml("#FFE25098"); } }

        public static XLColor RaspberryRose { get { return FromHtml("#FFB3446C"); } }

        public static XLColor RawUmber { get { return FromHtml("#FF826644"); } }

        public static XLColor RazzleDazzleRose { get { return FromHtml("#FFFF33CC"); } }

        public static XLColor Razzmatazz { get { return FromHtml("#FFE3256B"); } }

        public static XLColor RedMunsell { get { return FromHtml("#FFF2003C"); } }

        public static XLColor RedNcs { get { return FromHtml("#FFC40233"); } }

        public static XLColor RedPigment { get { return FromHtml("#FFED1C24"); } }

        public static XLColor RedRyb { get { return FromHtml("#FFFE2712"); } }

        public static XLColor Redwood { get { return FromHtml("#FFAB4E52"); } }

        public static XLColor Regalia { get { return FromHtml("#FF522D80"); } }

        public static XLColor RichBlack { get { return FromHtml("#FF004040"); } }

        public static XLColor RichBrilliantLavender { get { return FromHtml("#FFF1A7FE"); } }

        public static XLColor RichCarmine { get { return FromHtml("#FFD70040"); } }

        public static XLColor RichElectricBlue { get { return FromHtml("#FF0892D0"); } }

        public static XLColor RichLavender { get { return FromHtml("#FFA76BCF"); } }

        public static XLColor RichLilac { get { return FromHtml("#FFB666D2"); } }

        public static XLColor RichMaroon { get { return FromHtml("#FFB03060"); } }

        public static XLColor RifleGreen { get { return FromHtml("#FF414833"); } }

        public static XLColor RobinEggBlue { get { return FromHtml("#FF00CCCC"); } }

        public static XLColor Rose { get { return FromHtml("#FFFF007F"); } }

        public static XLColor RoseBonbon { get { return FromHtml("#FFF9429E"); } }

        public static XLColor RoseEbony { get { return FromHtml("#FF674846"); } }

        public static XLColor RoseGold { get { return FromHtml("#FFB76E79"); } }

        public static XLColor RoseMadder { get { return FromHtml("#FFE32636"); } }

        public static XLColor RosePink { get { return FromHtml("#FFFF66CC"); } }

        public static XLColor RoseQuartz { get { return FromHtml("#FFAA98A9"); } }

        public static XLColor RoseTaupe { get { return FromHtml("#FF905D5D"); } }

        public static XLColor RoseVale { get { return FromHtml("#FFAB4E52"); } }

        public static XLColor Rosewood { get { return FromHtml("#FF65000B"); } }

        public static XLColor RossoCorsa { get { return FromHtml("#FFD40000"); } }

        public static XLColor RoyalAzure { get { return FromHtml("#FF0038A8"); } }

        public static XLColor RoyalBlueTraditional { get { return FromHtml("#FF002366"); } }

        public static XLColor RoyalFuchsia { get { return FromHtml("#FFCA2C92"); } }

        public static XLColor RoyalPurple { get { return FromHtml("#FF7851A9"); } }

        public static XLColor Ruby { get { return FromHtml("#FFE0115F"); } }

        public static XLColor Ruddy { get { return FromHtml("#FFFF0028"); } }

        public static XLColor RuddyBrown { get { return FromHtml("#FFBB6528"); } }

        public static XLColor RuddyPink { get { return FromHtml("#FFE18E96"); } }

        public static XLColor Rufous { get { return FromHtml("#FFA81C07"); } }

        public static XLColor Russet { get { return FromHtml("#FF80461B"); } }

        public static XLColor Rust { get { return FromHtml("#FFB7410E"); } }

        public static XLColor SacramentoStateGreen { get { return FromHtml("#FF00563F"); } }

        public static XLColor SafetyOrangeBlazeOrange { get { return FromHtml("#FFFF6700"); } }

        public static XLColor Saffron { get { return FromHtml("#FFF4C430"); } }

        public static XLColor Salmon1 { get { return FromHtml("#FFFF8C69"); } }

        public static XLColor SalmonPink { get { return FromHtml("#FFFF91A4"); } }

        public static XLColor Sand { get { return FromHtml("#FFC2B280"); } }

        public static XLColor SandDune { get { return FromHtml("#FF967117"); } }

        public static XLColor Sandstorm { get { return FromHtml("#FFECD540"); } }

        public static XLColor SandyTaupe { get { return FromHtml("#FF967117"); } }

        public static XLColor Sangria { get { return FromHtml("#FF92000A"); } }

        public static XLColor SapGreen { get { return FromHtml("#FF507D2A"); } }

        public static XLColor Sapphire { get { return FromHtml("#FF082567"); } }

        public static XLColor SatinSheenGold { get { return FromHtml("#FFCBA135"); } }

        public static XLColor Scarlet { get { return FromHtml("#FFFF2000"); } }

        public static XLColor SchoolBusYellow { get { return FromHtml("#FFFFD800"); } }

        public static XLColor ScreaminGreen { get { return FromHtml("#FF76FF7A"); } }

        public static XLColor SealBrown { get { return FromHtml("#FF321414"); } }

        public static XLColor SelectiveYellow { get { return FromHtml("#FFFFBA00"); } }

        public static XLColor Sepia { get { return FromHtml("#FF704214"); } }

        public static XLColor Shadow { get { return FromHtml("#FF8A795D"); } }

        public static XLColor ShamrockGreen { get { return FromHtml("#FF009E60"); } }

        public static XLColor ShockingPink { get { return FromHtml("#FFFC0FC0"); } }

        public static XLColor Sienna1 { get { return FromHtml("#FF882D17"); } }

        public static XLColor Sinopia { get { return FromHtml("#FFCB410B"); } }

        public static XLColor Skobeloff { get { return FromHtml("#FF007474"); } }

        public static XLColor SkyMagenta { get { return FromHtml("#FFCF71AF"); } }

        public static XLColor SmaltDarkPowderBlue { get { return FromHtml("#FF003399"); } }

        public static XLColor SmokeyTopaz { get { return FromHtml("#FF933D41"); } }

        public static XLColor SmokyBlack { get { return FromHtml("#FF100C08"); } }

        public static XLColor SpiroDiscoBall { get { return FromHtml("#FF0FC0FC"); } }

        public static XLColor SplashedWhite { get { return FromHtml("#FFFEFDFF"); } }

        public static XLColor SpringBud { get { return FromHtml("#FFA7FC00"); } }

        public static XLColor StPatricksBlue { get { return FromHtml("#FF23297A"); } }

        public static XLColor StilDeGrainYellow { get { return FromHtml("#FFFADA5E"); } }

        public static XLColor Straw { get { return FromHtml("#FFE4D96F"); } }

        public static XLColor Sunglow { get { return FromHtml("#FFFFCC33"); } }

        public static XLColor Sunset { get { return FromHtml("#FFFAD6A5"); } }

        public static XLColor Tangelo { get { return FromHtml("#FFF94D00"); } }

        public static XLColor Tangerine { get { return FromHtml("#FFF28500"); } }

        public static XLColor TangerineYellow { get { return FromHtml("#FFFFCC00"); } }

        public static XLColor Taupe { get { return FromHtml("#FF483C32"); } }

        public static XLColor TaupeGray { get { return FromHtml("#FF8B8589"); } }

        public static XLColor TeaGreen { get { return FromHtml("#FFD0F0C0"); } }

        public static XLColor TealBlue { get { return FromHtml("#FF367588"); } }

        public static XLColor TealGreen { get { return FromHtml("#FF006D5B"); } }

        public static XLColor TeaRoseOrange { get { return FromHtml("#FFF88379"); } }

        public static XLColor TeaRoseRose { get { return FromHtml("#FFF4C2C2"); } }

        public static XLColor TennéTawny { get { return FromHtml("#FFCD5700"); } }

        public static XLColor TerraCotta { get { return FromHtml("#FFE2725B"); } }

        public static XLColor ThulianPink { get { return FromHtml("#FFDE6FA1"); } }

        public static XLColor TickleMePink { get { return FromHtml("#FFFC89AC"); } }

        public static XLColor TiffanyBlue { get { return FromHtml("#FF0ABAB5"); } }

        public static XLColor TigersEye { get { return FromHtml("#FFE08D3C"); } }

        public static XLColor Timberwolf { get { return FromHtml("#FFDBD7D2"); } }

        public static XLColor TitaniumYellow { get { return FromHtml("#FFEEE600"); } }

        public static XLColor Toolbox { get { return FromHtml("#FF746CC0"); } }

        public static XLColor TractorRed { get { return FromHtml("#FFFD0E35"); } }

        public static XLColor TropicalRainForest { get { return FromHtml("#FF00755E"); } }

        public static XLColor TuftsBlue { get { return FromHtml("#FF417DC1"); } }

        public static XLColor Tumbleweed { get { return FromHtml("#FFDEAA88"); } }

        public static XLColor TurkishRose { get { return FromHtml("#FFB57281"); } }

        public static XLColor Turquoise1 { get { return FromHtml("#FF30D5C8"); } }

        public static XLColor TurquoiseBlue { get { return FromHtml("#FF00FFEF"); } }

        public static XLColor TurquoiseGreen { get { return FromHtml("#FFA0D6B4"); } }

        public static XLColor TuscanRed { get { return FromHtml("#FF823535"); } }

        public static XLColor TwilightLavender { get { return FromHtml("#FF8A496B"); } }

        public static XLColor TyrianPurple { get { return FromHtml("#FF66023C"); } }

        public static XLColor UaBlue { get { return FromHtml("#FF0033AA"); } }

        public static XLColor UaRed { get { return FromHtml("#FFD9004C"); } }

        public static XLColor Ube { get { return FromHtml("#FF8878C3"); } }

        public static XLColor UclaBlue { get { return FromHtml("#FF536895"); } }

        public static XLColor UclaGold { get { return FromHtml("#FFFFB300"); } }

        public static XLColor UfoGreen { get { return FromHtml("#FF3CD070"); } }

        public static XLColor Ultramarine { get { return FromHtml("#FF120A8F"); } }

        public static XLColor UltramarineBlue { get { return FromHtml("#FF4166F5"); } }

        public static XLColor UltraPink { get { return FromHtml("#FFFF6FFF"); } }

        public static XLColor Umber { get { return FromHtml("#FF635147"); } }

        public static XLColor UnitedNationsBlue { get { return FromHtml("#FF5B92E5"); } }

        public static XLColor UnmellowYellow { get { return FromHtml("#FFFFFF66"); } }

        public static XLColor UpForestGreen { get { return FromHtml("#FF014421"); } }

        public static XLColor UpMaroon { get { return FromHtml("#FF7B1113"); } }

        public static XLColor UpsdellRed { get { return FromHtml("#FFAE2029"); } }

        public static XLColor Urobilin { get { return FromHtml("#FFE1AD21"); } }

        public static XLColor UscCardinal { get { return FromHtml("#FF990000"); } }

        public static XLColor UscGold { get { return FromHtml("#FFFFCC00"); } }

        public static XLColor UtahCrimson { get { return FromHtml("#FFD3003F"); } }

        public static XLColor Vanilla { get { return FromHtml("#FFF3E5AB"); } }

        public static XLColor VegasGold { get { return FromHtml("#FFC5B358"); } }

        public static XLColor VenetianRed { get { return FromHtml("#FFC80815"); } }

        public static XLColor Verdigris { get { return FromHtml("#FF43B3AE"); } }

        public static XLColor Vermilion { get { return FromHtml("#FFE34234"); } }

        public static XLColor Veronica { get { return FromHtml("#FFA020F0"); } }

        public static XLColor Violet1 { get { return FromHtml("#FF8F00FF"); } }

        public static XLColor VioletColorWheel { get { return FromHtml("#FF7F00FF"); } }

        public static XLColor VioletRyb { get { return FromHtml("#FF8601AF"); } }

        public static XLColor Viridian { get { return FromHtml("#FF40826D"); } }

        public static XLColor VividAuburn { get { return FromHtml("#FF922724"); } }

        public static XLColor VividBurgundy { get { return FromHtml("#FF9F1D35"); } }

        public static XLColor VividCerise { get { return FromHtml("#FFDA1D81"); } }

        public static XLColor VividTangerine { get { return FromHtml("#FFFFA089"); } }

        public static XLColor VividViolet { get { return FromHtml("#FF9F00FF"); } }

        public static XLColor WarmBlack { get { return FromHtml("#FF004242"); } }

        public static XLColor Wenge { get { return FromHtml("#FF645452"); } }

        public static XLColor WildBlueYonder { get { return FromHtml("#FFA2ADD0"); } }

        public static XLColor WildStrawberry { get { return FromHtml("#FFFF43A4"); } }

        public static XLColor WildWatermelon { get { return FromHtml("#FFFC6C85"); } }

        public static XLColor Wisteria { get { return FromHtml("#FFC9A0DC"); } }

        public static XLColor Xanadu { get { return FromHtml("#FF738678"); } }

        public static XLColor YaleBlue { get { return FromHtml("#FF0F4D92"); } }

        public static XLColor YellowMunsell { get { return FromHtml("#FFEFCC00"); } }

        public static XLColor YellowNcs { get { return FromHtml("#FFFFD300"); } }

        public static XLColor YellowProcess { get { return FromHtml("#FFFFEF00"); } }

        public static XLColor YellowRyb { get { return FromHtml("#FFFEFE33"); } }

        public static XLColor Zaffre { get { return FromHtml("#FF0014A8"); } }

        public static XLColor ZinnwalditeBrown { get { return FromHtml("#FF2C1608"); } }

        public static XLColor Transparent { get { return FromColor(Color.Transparent); } }
    }
}
