using ClosedXML.Excel.Caching;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using SkiaSharp;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        private static readonly XLColorRepository Repository = new XLColorRepository(key => new XLColor(key));

        private static readonly Dictionary<SKColors, XLColor> ByColor = new Dictionary<SKColors, XLColor>();
        private static readonly object ByColorLock = new object();

        internal static XLColor FromKey(ref XLColorKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        public static XLColor FromColor(SKColor color)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Color,
                Color = color
            };
            return FromKey(ref key);
        }

        public static XLColor FromColor(Color color)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Color,
                Color = SKColor.Parse(color.Rgb.Value)
            };
            return FromKey(ref key);
        }

        public static XLColor FromArgb(int argb)
        {
            return FromColor(ColorStringParser.FromArgb(argb));
        }

        public static XLColor FromArgb(int r, int g, int b)
        {
            return FromColor(ColorStringParser.FromArgb(r, g, b));
        }

        public static XLColor FromArgb(int a, int r, int g, int b)
        {
            return FromColor(ColorStringParser.FromArgb(a, r, g, b));
        }

        public static XLColor FromName(string name)
        {
            return FromColor(ColorStringParser.FromName(name));
        }

        public static XLColor FromHtml(string htmlColor)
        {
            return FromColor(ColorStringParser.ParseFromHtml(htmlColor));
        }

        public static XLColor FromIndex(int index)
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

        public static XLColor FromTheme(XLThemeColor themeColor, double themeTint)
        {
            var key = new XLColorKey
            {
                ColorType = XLColorType.Theme,
                ThemeColor = themeColor,
                ThemeTint = themeTint
            };
            return FromKey(ref key);
        }

        private static Dictionary<int, XLColor> _indexedColors;

        public static Dictionary<int, XLColor> IndexedColors
        {
            get
            {
                if (_indexedColors == null)
                {
                    var retVal = new Dictionary<int, XLColor>
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
                            {64, FromColor(SKColors.Transparent)}
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

        public static XLColor AliceBlue => FromColor(SKColors.AliceBlue);

        public static XLColor AntiqueWhite => FromColor(SKColors.AntiqueWhite);

        public static XLColor Aqua => FromColor(SKColors.Aqua);

        public static XLColor Aquamarine => FromColor(SKColors.Aquamarine);

        public static XLColor Azure => FromColor(SKColors.Azure);

        public static XLColor Beige => FromColor(SKColors.Beige);

        public static XLColor Bisque => FromColor(SKColors.Bisque);

        public static XLColor Black => FromColor(SKColors.Black);

        public static XLColor BlanchedAlmond => FromColor(SKColors.BlanchedAlmond);

        public static XLColor Blue => FromColor(SKColors.Blue);

        public static XLColor BlueViolet => FromColor(SKColors.BlueViolet);

        public static XLColor Brown => FromColor(SKColors.Brown);

        public static XLColor BurlyWood => FromColor(SKColors.BurlyWood);

        public static XLColor CadetBlue => FromColor(SKColors.CadetBlue);

        public static XLColor Chartreuse => FromColor(SKColors.Chartreuse);

        public static XLColor Chocolate => FromColor(SKColors.Chocolate);

        public static XLColor Coral => FromColor(SKColors.Coral);

        public static XLColor CornflowerBlue => FromColor(SKColors.CornflowerBlue);

        public static XLColor Cornsilk => FromColor(SKColors.Cornsilk);

        public static XLColor Crimson => FromColor(SKColors.Crimson);

        public static XLColor Cyan => FromColor(SKColors.Cyan);

        public static XLColor DarkBlue => FromColor(SKColors.DarkBlue);

        public static XLColor DarkCyan => FromColor(SKColors.DarkCyan);

        public static XLColor DarkGoldenrod => FromColor(SKColors.DarkGoldenrod);

        public static XLColor DarkGray => FromColor(SKColors.DarkGray);

        public static XLColor DarkGreen => FromColor(SKColors.DarkGreen);

        public static XLColor DarkKhaki => FromColor(SKColors.DarkKhaki);

        public static XLColor DarkMagenta => FromColor(SKColors.DarkMagenta);

        public static XLColor DarkOliveGreen => FromColor(SKColors.DarkOliveGreen);

        public static XLColor DarkOrange => FromColor(SKColors.DarkOrange);

        public static XLColor DarkOrchid => FromColor(SKColors.DarkOrchid);

        public static XLColor DarkRed => FromColor(SKColors.DarkRed);

        public static XLColor DarkSalmon => FromColor(SKColors.DarkSalmon);

        public static XLColor DarkSeaGreen => FromColor(SKColors.DarkSeaGreen);

        public static XLColor DarkSlateBlue => FromColor(SKColors.DarkSlateBlue);

        public static XLColor DarkSlateGray => FromColor(SKColors.DarkSlateGray);

        public static XLColor DarkTurquoise => FromColor(SKColors.DarkTurquoise);

        public static XLColor DarkViolet => FromColor(SKColors.DarkViolet);

        public static XLColor DeepPink => FromColor(SKColors.DeepPink);

        public static XLColor DeepSkyBlue => FromColor(SKColors.DeepSkyBlue);

        public static XLColor DimGray => FromColor(SKColors.DimGray);

        public static XLColor DodgerBlue => FromColor(SKColors.DodgerBlue);

        public static XLColor Firebrick => FromColor(SKColors.Firebrick);

        public static XLColor FloralWhite => FromColor(SKColors.FloralWhite);

        public static XLColor ForestGreen => FromColor(SKColors.ForestGreen);

        public static XLColor Fuchsia => FromColor(SKColors.Fuchsia);

        public static XLColor Gainsboro => FromColor(SKColors.Gainsboro);

        public static XLColor GhostWhite => FromColor(SKColors.GhostWhite);

        public static XLColor Gold => FromColor(SKColors.Gold);

        public static XLColor Goldenrod => FromColor(SKColors.Goldenrod);

        public static XLColor Gray => FromColor(SKColors.Gray);

        public static XLColor Green => FromColor(SKColors.Green);

        public static XLColor GreenYellow => FromColor(SKColors.GreenYellow);

        public static XLColor Honeydew => FromColor(SKColors.Honeydew);

        public static XLColor HotPink => FromColor(SKColors.HotPink);

        public static XLColor IndianRed => FromColor(SKColors.IndianRed);

        public static XLColor Indigo => FromColor(SKColors.Indigo);

        public static XLColor Ivory => FromColor(SKColors.Ivory);

        public static XLColor Khaki => FromColor(SKColors.Khaki);

        public static XLColor Lavender => FromColor(SKColors.Lavender);

        public static XLColor LavenderBlush => FromColor(SKColors.LavenderBlush);

        public static XLColor LawnGreen => FromColor(SKColors.LawnGreen);

        public static XLColor LemonChiffon => FromColor(SKColors.LemonChiffon);

        public static XLColor LightBlue => FromColor(SKColors.LightBlue);

        public static XLColor LightCoral => FromColor(SKColors.LightCoral);

        public static XLColor LightCyan => FromColor(SKColors.LightCyan);

        public static XLColor LightGoldenrodYellow => FromColor(SKColors.LightGoldenrodYellow);

        public static XLColor LightGray => FromColor(SKColors.LightGray);

        public static XLColor LightGreen => FromColor(SKColors.LightGreen);

        public static XLColor LightPink => FromColor(SKColors.LightPink);

        public static XLColor LightSalmon => FromColor(SKColors.LightSalmon);

        public static XLColor LightSeaGreen => FromColor(SKColors.LightSeaGreen);

        public static XLColor LightSkyBlue => FromColor(SKColors.LightSkyBlue);

        public static XLColor LightSlateGray => FromColor(SKColors.LightSlateGray);

        public static XLColor LightSteelBlue => FromColor(SKColors.LightSteelBlue);

        public static XLColor LightYellow => FromColor(SKColors.LightYellow);

        public static XLColor Lime => FromColor(SKColors.Lime);

        public static XLColor LimeGreen => FromColor(SKColors.LimeGreen);

        public static XLColor Linen => FromColor(SKColors.Linen);

        public static XLColor Magenta => FromColor(SKColors.Magenta);

        public static XLColor Maroon => FromColor(SKColors.Maroon);

        public static XLColor MediumAquamarine => FromColor(SKColors.MediumAquamarine);

        public static XLColor MediumBlue => FromColor(SKColors.MediumBlue);

        public static XLColor MediumOrchid => FromColor(SKColors.MediumOrchid);

        public static XLColor MediumPurple => FromColor(SKColors.MediumPurple);

        public static XLColor MediumSeaGreen => FromColor(SKColors.MediumSeaGreen);

        public static XLColor MediumSlateBlue => FromColor(SKColors.MediumSlateBlue);

        public static XLColor MediumSpringGreen => FromColor(SKColors.MediumSpringGreen);

        public static XLColor MediumTurquoise => FromColor(SKColors.MediumTurquoise);

        public static XLColor MediumVioletRed => FromColor(SKColors.MediumVioletRed);

        public static XLColor MidnightBlue => FromColor(SKColors.MidnightBlue);

        public static XLColor MintCream => FromColor(SKColors.MintCream);

        public static XLColor MistyRose => FromColor(SKColors.MistyRose);

        public static XLColor Moccasin => FromColor(SKColors.Moccasin);

        public static XLColor NavajoWhite => FromColor(SKColors.NavajoWhite);

        public static XLColor Navy => FromColor(SKColors.Navy);

        public static XLColor OldLace => FromColor(SKColors.OldLace);

        public static XLColor Olive => FromColor(SKColors.Olive);

        public static XLColor OliveDrab => FromColor(SKColors.OliveDrab);

        public static XLColor Orange => FromColor(SKColors.Orange);

        public static XLColor OrangeRed => FromColor(SKColors.OrangeRed);

        public static XLColor Orchid => FromColor(SKColors.Orchid);

        public static XLColor PaleGoldenrod => FromColor(SKColors.PaleGoldenrod);

        public static XLColor PaleGreen => FromColor(SKColors.PaleGreen);

        public static XLColor PaleTurquoise => FromColor(SKColors.PaleTurquoise);

        public static XLColor PaleVioletRed => FromColor(SKColors.PaleVioletRed);

        public static XLColor PapayaWhip => FromColor(SKColors.PapayaWhip);

        public static XLColor PeachPuff => FromColor(SKColors.PeachPuff);

        public static XLColor Peru => FromColor(SKColors.Peru);

        public static XLColor Pink => FromColor(SKColors.Pink);

        public static XLColor Plum => FromColor(SKColors.Plum);

        public static XLColor PowderBlue => FromColor(SKColors.PowderBlue);

        public static XLColor Purple => FromColor(SKColors.Purple);

        public static XLColor Red => FromColor(SKColors.Red);

        public static XLColor RosyBrown => FromColor(SKColors.RosyBrown);

        public static XLColor RoyalBlue => FromColor(SKColors.RoyalBlue);

        public static XLColor SaddleBrown => FromColor(SKColors.SaddleBrown);

        public static XLColor Salmon => FromColor(SKColors.Salmon);

        public static XLColor SandyBrown => FromColor(SKColors.SandyBrown);

        public static XLColor SeaGreen => FromColor(SKColors.SeaGreen);

        public static XLColor SeaShell => FromColor(SKColors.SeaShell);

        public static XLColor Sienna => FromColor(SKColors.Sienna);

        public static XLColor Silver => FromColor(SKColors.Silver);

        public static XLColor SkyBlue => FromColor(SKColors.SkyBlue);

        public static XLColor SlateBlue => FromColor(SKColors.SlateBlue);

        public static XLColor SlateGray => FromColor(SKColors.SlateGray);

        public static XLColor Snow => FromColor(SKColors.Snow);

        public static XLColor SpringGreen => FromColor(SKColors.SpringGreen);

        public static XLColor SteelBlue => FromColor(SKColors.SteelBlue);

        public static XLColor Tan => FromColor(SKColors.Tan);

        public static XLColor Teal => FromColor(SKColors.Teal);

        public static XLColor Thistle => FromColor(SKColors.Thistle);

        public static XLColor Tomato => FromColor(SKColors.Tomato);

        public static XLColor Turquoise => FromColor(SKColors.Turquoise);

        public static XLColor Violet => FromColor(SKColors.Violet);

        public static XLColor Wheat => FromColor(SKColors.Wheat);

        public static XLColor White => FromColor(SKColors.White);

        public static XLColor WhiteSmoke => FromColor(SKColors.WhiteSmoke);

        public static XLColor Yellow => FromColor(SKColors.Yellow);

        public static XLColor YellowGreen => FromColor(SKColors.YellowGreen);

        public static XLColor AirForceBlue => FromHtml("#FF5D8AA8");

        public static XLColor Alizarin => FromHtml("#FFE32636");

        public static XLColor Almond => FromHtml("#FFEFDECD");

        public static XLColor Amaranth => FromHtml("#FFE52B50");

        public static XLColor Amber => FromHtml("#FFFFBF00");

        public static XLColor AmberSaeEce => FromHtml("#FFFF7E00");

        public static XLColor AmericanRose => FromHtml("#FFFF033E");

        public static XLColor Amethyst => FromHtml("#FF9966CC");

        public static XLColor AntiFlashWhite => FromHtml("#FFF2F3F4");

        public static XLColor AntiqueBrass => FromHtml("#FFCD9575");

        public static XLColor AntiqueFuchsia => FromHtml("#FF915C83");

        public static XLColor AppleGreen => FromHtml("#FF8DB600");

        public static XLColor Apricot => FromHtml("#FFFBCEB1");

        public static XLColor Aquamarine1 => FromHtml("#FF7FFFD0");

        public static XLColor ArmyGreen => FromHtml("#FF4B5320");

        public static XLColor Arsenic => FromHtml("#FF3B444B");

        public static XLColor ArylideYellow => FromHtml("#FFE9D66B");

        public static XLColor AshGrey => FromHtml("#FFB2BEB5");

        public static XLColor Asparagus => FromHtml("#FF87A96B");

        public static XLColor AtomicTangerine => FromHtml("#FFFF9966");

        public static XLColor Auburn => FromHtml("#FF6D351A");

        public static XLColor Aureolin => FromHtml("#FFFDEE00");

        public static XLColor Aurometalsaurus => FromHtml("#FF6E7F80");

        public static XLColor Awesome => FromHtml("#FFFF2052");

        public static XLColor AzureColorWheel => FromHtml("#FF007FFF");

        public static XLColor BabyBlue => FromHtml("#FF89CFF0");

        public static XLColor BabyBlueEyes => FromHtml("#FFA1CAF1");

        public static XLColor BabyPink => FromHtml("#FFF4C2C2");

        public static XLColor BallBlue => FromHtml("#FF21ABCD");

        public static XLColor BananaMania => FromHtml("#FFFAE7B5");

        public static XLColor BattleshipGrey => FromHtml("#FF848482");

        public static XLColor Bazaar => FromHtml("#FF98777B");

        public static XLColor BeauBlue => FromHtml("#FFBCD4E6");

        public static XLColor Beaver => FromHtml("#FF9F8170");

        public static XLColor Bistre => FromHtml("#FF3D2B1F");

        public static XLColor Bittersweet => FromHtml("#FFFE6F5E");

        public static XLColor BleuDeFrance => FromHtml("#FF318CE7");

        public static XLColor BlizzardBlue => FromHtml("#FFACE5EE");

        public static XLColor Blond => FromHtml("#FFFAF0BE");

        public static XLColor BlueBell => FromHtml("#FFA2A2D0");

        public static XLColor BlueGray => FromHtml("#FF6699CC");

        public static XLColor BlueGreen => FromHtml("#FF00DDDD");

        public static XLColor BluePigment => FromHtml("#FF333399");

        public static XLColor BlueRyb => FromHtml("#FF0247FE");

        public static XLColor Blush => FromHtml("#FFDE5D83");

        public static XLColor Bole => FromHtml("#FF79443B");

        public static XLColor BondiBlue => FromHtml("#FF0095B6");

        public static XLColor BostonUniversityRed => FromHtml("#FFCC0000");

        public static XLColor BrandeisBlue => FromHtml("#FF0070FF");

        public static XLColor Brass => FromHtml("#FFB5A642");

        public static XLColor BrickRed => FromHtml("#FFCB4154");

        public static XLColor BrightCerulean => FromHtml("#FF1DACD6");

        public static XLColor BrightGreen => FromHtml("#FF66FF00");

        public static XLColor BrightLavender => FromHtml("#FFBF94E4");

        public static XLColor BrightMaroon => FromHtml("#FFC32148");

        public static XLColor BrightPink => FromHtml("#FFFF007F");

        public static XLColor BrightTurquoise => FromHtml("#FF08E8DE");

        public static XLColor BrightUbe => FromHtml("#FFD19FE8");

        public static XLColor BrilliantLavender => FromHtml("#FFF4BBFF");

        public static XLColor BrilliantRose => FromHtml("#FFFF55A3");

        public static XLColor BrinkPink => FromHtml("#FFFB607F");

        public static XLColor BritishRacingGreen => FromHtml("#FF004225");

        public static XLColor Bronze => FromHtml("#FFCD7F32");

        public static XLColor BrownTraditional => FromHtml("#FF964B00");

        public static XLColor BubbleGum => FromHtml("#FFFFC1CC");

        public static XLColor Bubbles => FromHtml("#FFE7FEFF");

        public static XLColor Buff => FromHtml("#FFF0DC82");

        public static XLColor BulgarianRose => FromHtml("#FF480607");

        public static XLColor Burgundy => FromHtml("#FF800020");

        public static XLColor BurntOrange => FromHtml("#FFCC5500");

        public static XLColor BurntSienna => FromHtml("#FFE97451");

        public static XLColor BurntUmber => FromHtml("#FF8A3324");

        public static XLColor Byzantine => FromHtml("#FFBD33A4");

        public static XLColor Byzantium => FromHtml("#FF702963");

        public static XLColor Cadet => FromHtml("#FF536872");

        public static XLColor CadetGrey => FromHtml("#FF91A3B0");

        public static XLColor CadmiumGreen => FromHtml("#FF006B3C");

        public static XLColor CadmiumOrange => FromHtml("#FFED872D");

        public static XLColor CadmiumRed => FromHtml("#FFE30022");

        public static XLColor CadmiumYellow => FromHtml("#FFFFF600");

        public static XLColor CalPolyPomonaGreen => FromHtml("#FF1E4D2B");

        public static XLColor CambridgeBlue => FromHtml("#FFA3C1AD");

        public static XLColor Camel => FromHtml("#FFC19A6B");

        public static XLColor CamouflageGreen => FromHtml("#FF78866B");

        public static XLColor CanaryYellow => FromHtml("#FFFFEF00");

        public static XLColor CandyAppleRed => FromHtml("#FFFF0800");

        public static XLColor CandyPink => FromHtml("#FFE4717A");

        public static XLColor CaputMortuum => FromHtml("#FF592720");

        public static XLColor Cardinal => FromHtml("#FFC41E3A");

        public static XLColor CaribbeanGreen => FromHtml("#FF00CC99");

        public static XLColor Carmine => FromHtml("#FF960018");

        public static XLColor CarminePink => FromHtml("#FFEB4C42");

        public static XLColor CarmineRed => FromHtml("#FFFF0038");

        public static XLColor CarnationPink => FromHtml("#FFFFA6C9");

        public static XLColor Carnelian => FromHtml("#FFB31B1B");

        public static XLColor CarolinaBlue => FromHtml("#FF99BADD");

        public static XLColor CarrotOrange => FromHtml("#FFED9121");

        public static XLColor Ceil => FromHtml("#FF92A1CF");

        public static XLColor Celadon => FromHtml("#FFACE1AF");

        public static XLColor CelestialBlue => FromHtml("#FF4997D0");

        public static XLColor Cerise => FromHtml("#FFDE3163");

        public static XLColor CerisePink => FromHtml("#FFEC3B83");

        public static XLColor Cerulean => FromHtml("#FF007BA7");

        public static XLColor CeruleanBlue => FromHtml("#FF2A52BE");

        public static XLColor Chamoisee => FromHtml("#FFA0785A");

        public static XLColor Champagne => FromHtml("#FFF7E7CE");

        public static XLColor Charcoal => FromHtml("#FF36454F");

        public static XLColor ChartreuseTraditional => FromHtml("#FFDFFF00");

        public static XLColor CherryBlossomPink => FromHtml("#FFFFB7C5");

        public static XLColor Chocolate1 => FromHtml("#FF7B3F00");

        public static XLColor ChromeYellow => FromHtml("#FFFFA700");

        public static XLColor Cinereous => FromHtml("#FF98817B");

        public static XLColor Cinnabar => FromHtml("#FFE34234");

        public static XLColor Citrine => FromHtml("#FFE4D00A");

        public static XLColor ClassicRose => FromHtml("#FFFBCCE7");

        public static XLColor Cobalt => FromHtml("#FF0047AB");

        public static XLColor ColumbiaBlue => FromHtml("#FF9BDDFF");

        public static XLColor CoolBlack => FromHtml("#FF002E63");

        public static XLColor CoolGrey => FromHtml("#FF8C92AC");

        public static XLColor Copper => FromHtml("#FFB87333");

        public static XLColor CopperRose => FromHtml("#FF996666");

        public static XLColor Coquelicot => FromHtml("#FFFF3800");

        public static XLColor CoralPink => FromHtml("#FFF88379");

        public static XLColor CoralRed => FromHtml("#FFFF4040");

        public static XLColor Cordovan => FromHtml("#FF893F45");

        public static XLColor Corn => FromHtml("#FFFBEC5D");

        public static XLColor CornellRed => FromHtml("#FFB31B1B");

        public static XLColor CosmicLatte => FromHtml("#FFFFF8E7");

        public static XLColor CottonCandy => FromHtml("#FFFFBCD9");

        public static XLColor Cream => FromHtml("#FFFFFDD0");

        public static XLColor CrimsonGlory => FromHtml("#FFBE0032");

        public static XLColor CyanProcess => FromHtml("#FF00B7EB");

        public static XLColor Daffodil => FromHtml("#FFFFFF31");

        public static XLColor Dandelion => FromHtml("#FFF0E130");

        public static XLColor DarkBrown => FromHtml("#FF654321");

        public static XLColor DarkByzantium => FromHtml("#FF5D3954");

        public static XLColor DarkCandyAppleRed => FromHtml("#FFA40000");

        public static XLColor DarkCerulean => FromHtml("#FF08457E");

        public static XLColor DarkChampagne => FromHtml("#FFC2B280");

        public static XLColor DarkChestnut => FromHtml("#FF986960");

        public static XLColor DarkCoral => FromHtml("#FFCD5B45");

        public static XLColor DarkElectricBlue => FromHtml("#FF536878");

        public static XLColor DarkGreen1 => FromHtml("#FF013220");

        public static XLColor DarkJungleGreen => FromHtml("#FF1A2421");

        public static XLColor DarkLava => FromHtml("#FF483C32");

        public static XLColor DarkLavender => FromHtml("#FF734F96");

        public static XLColor DarkMidnightBlue => FromHtml("#FF003366");

        public static XLColor DarkPastelBlue => FromHtml("#FF779ECB");

        public static XLColor DarkPastelGreen => FromHtml("#FF03C03C");

        public static XLColor DarkPastelPurple => FromHtml("#FF966FD6");

        public static XLColor DarkPastelRed => FromHtml("#FFC23B22");

        public static XLColor DarkPink => FromHtml("#FFE75480");

        public static XLColor DarkPowderBlue => FromHtml("#FF003399");

        public static XLColor DarkRaspberry => FromHtml("#FF872657");

        public static XLColor DarkScarlet => FromHtml("#FF560319");

        public static XLColor DarkSienna => FromHtml("#FF3C1414");

        public static XLColor DarkSpringGreen => FromHtml("#FF177245");

        public static XLColor DarkTan => FromHtml("#FF918151");

        public static XLColor DarkTangerine => FromHtml("#FFFFA812");

        public static XLColor DarkTaupe => FromHtml("#FF483C32");

        public static XLColor DarkTerraCotta => FromHtml("#FFCC4E5C");

        public static XLColor DartmouthGreen => FromHtml("#FF00693E");

        public static XLColor DavysGrey => FromHtml("#FF555555");

        public static XLColor DebianRed => FromHtml("#FFD70A53");

        public static XLColor DeepCarmine => FromHtml("#FFA9203E");

        public static XLColor DeepCarminePink => FromHtml("#FFEF3038");

        public static XLColor DeepCarrotOrange => FromHtml("#FFE9692C");

        public static XLColor DeepCerise => FromHtml("#FFDA3287");

        public static XLColor DeepChampagne => FromHtml("#FFFAD6A5");

        public static XLColor DeepChestnut => FromHtml("#FFB94E48");

        public static XLColor DeepFuchsia => FromHtml("#FFC154C1");

        public static XLColor DeepJungleGreen => FromHtml("#FF004B49");

        public static XLColor DeepLilac => FromHtml("#FF9955BB");

        public static XLColor DeepMagenta => FromHtml("#FFCC00CC");

        public static XLColor DeepPeach => FromHtml("#FFFFCBA4");

        public static XLColor DeepSaffron => FromHtml("#FFFF9933");

        public static XLColor Denim => FromHtml("#FF1560BD");

        public static XLColor Desert => FromHtml("#FFC19A6B");

        public static XLColor DesertSand => FromHtml("#FFEDC9AF");

        public static XLColor DogwoodRose => FromHtml("#FFD71868");

        public static XLColor DollarBill => FromHtml("#FF85BB65");

        public static XLColor Drab => FromHtml("#FF967117");

        public static XLColor DukeBlue => FromHtml("#FF00009C");

        public static XLColor EarthYellow => FromHtml("#FFE1A95F");

        public static XLColor Ecru => FromHtml("#FFC2B280");

        public static XLColor Eggplant => FromHtml("#FF614051");

        public static XLColor Eggshell => FromHtml("#FFF0EAD6");

        public static XLColor EgyptianBlue => FromHtml("#FF1034A6");

        public static XLColor ElectricBlue => FromHtml("#FF7DF9FF");

        public static XLColor ElectricCrimson => FromHtml("#FFFF003F");

        public static XLColor ElectricIndigo => FromHtml("#FF6F00FF");

        public static XLColor ElectricLavender => FromHtml("#FFF4BBFF");

        public static XLColor ElectricLime => FromHtml("#FFCCFF00");

        public static XLColor ElectricPurple => FromHtml("#FFBF00FF");

        public static XLColor ElectricUltramarine => FromHtml("#FF3F00FF");

        public static XLColor ElectricViolet => FromHtml("#FF8F00FF");

        public static XLColor Emerald => FromHtml("#FF50C878");

        public static XLColor EtonBlue => FromHtml("#FF96C8A2");

        public static XLColor Fallow => FromHtml("#FFC19A6B");

        public static XLColor FaluRed => FromHtml("#FF801818");

        public static XLColor Fandango => FromHtml("#FFB53389");

        public static XLColor FashionFuchsia => FromHtml("#FFF400A1");

        public static XLColor Fawn => FromHtml("#FFE5AA70");

        public static XLColor Feldgrau => FromHtml("#FF4D5D53");

        public static XLColor FernGreen => FromHtml("#FF4F7942");

        public static XLColor FerrariRed => FromHtml("#FFFF2800");

        public static XLColor FieldDrab => FromHtml("#FF6C541E");

        public static XLColor FireEngineRed => FromHtml("#FFCE2029");

        public static XLColor Flame => FromHtml("#FFE25822");

        public static XLColor FlamingoPink => FromHtml("#FFFC8EAC");

        public static XLColor Flavescent => FromHtml("#FFF7E98E");

        public static XLColor Flax => FromHtml("#FFEEDC82");

        public static XLColor FluorescentOrange => FromHtml("#FFFFBF00");

        public static XLColor FluorescentYellow => FromHtml("#FFCCFF00");

        public static XLColor Folly => FromHtml("#FFFF004F");

        public static XLColor ForestGreenTraditional => FromHtml("#FF014421");

        public static XLColor FrenchBeige => FromHtml("#FFA67B5B");

        public static XLColor FrenchBlue => FromHtml("#FF0072BB");

        public static XLColor FrenchLilac => FromHtml("#FF86608E");

        public static XLColor FrenchRose => FromHtml("#FFF64A8A");

        public static XLColor FuchsiaPink => FromHtml("#FFFF77FF");

        public static XLColor Fulvous => FromHtml("#FFE48400");

        public static XLColor FuzzyWuzzy => FromHtml("#FFCC6666");

        public static XLColor Gamboge => FromHtml("#FFE49B0F");

        public static XLColor Ginger => FromHtml("#FFF9F9FF");

        public static XLColor Glaucous => FromHtml("#FF6082B6");

        public static XLColor GoldenBrown => FromHtml("#FF996515");

        public static XLColor GoldenPoppy => FromHtml("#FFFCC200");

        public static XLColor GoldenYellow => FromHtml("#FFFFDF00");

        public static XLColor GoldMetallic => FromHtml("#FFD4AF37");

        public static XLColor GrannySmithApple => FromHtml("#FFA8E4A0");

        public static XLColor GrayAsparagus => FromHtml("#FF465945");

        public static XLColor GreenPigment => FromHtml("#FF00A550");

        public static XLColor GreenRyb => FromHtml("#FF66B032");

        public static XLColor Grullo => FromHtml("#FFA99A86");

        public static XLColor HalayaUbe => FromHtml("#FF663854");

        public static XLColor HanBlue => FromHtml("#FF446CCF");

        public static XLColor HanPurple => FromHtml("#FF5218FA");

        public static XLColor HansaYellow => FromHtml("#FFE9D66B");

        public static XLColor Harlequin => FromHtml("#FF3FFF00");

        public static XLColor HarvardCrimson => FromHtml("#FFC90016");

        public static XLColor HarvestGold => FromHtml("#FFDA9100");

        public static XLColor Heliotrope => FromHtml("#FFDF73FF");

        public static XLColor HollywoodCerise => FromHtml("#FFF400A1");

        public static XLColor HookersGreen => FromHtml("#FF007000");

        public static XLColor HotMagenta => FromHtml("#FFFF1DCE");

        public static XLColor HunterGreen => FromHtml("#FF355E3B");

        public static XLColor Iceberg => FromHtml("#FF71A6D2");

        public static XLColor Icterine => FromHtml("#FFFCF75E");

        public static XLColor Inchworm => FromHtml("#FFB2EC5D");

        public static XLColor IndiaGreen => FromHtml("#FF138808");

        public static XLColor IndianYellow => FromHtml("#FFE3A857");

        public static XLColor IndigoDye => FromHtml("#FF00416A");

        public static XLColor InternationalKleinBlue => FromHtml("#FF002FA7");

        public static XLColor InternationalOrange => FromHtml("#FFFF4F00");

        public static XLColor Iris => FromHtml("#FF5A4FCF");

        public static XLColor Isabelline => FromHtml("#FFF4F0EC");

        public static XLColor IslamicGreen => FromHtml("#FF009000");

        public static XLColor Jade => FromHtml("#FF00A86B");

        public static XLColor Jasper => FromHtml("#FFD73B3E");

        public static XLColor JazzberryJam => FromHtml("#FFA50B5E");

        public static XLColor Jonquil => FromHtml("#FFFADA5E");

        public static XLColor JuneBud => FromHtml("#FFBDDA57");

        public static XLColor JungleGreen => FromHtml("#FF29AB87");

        public static XLColor KellyGreen => FromHtml("#FF4CBB17");

        public static XLColor KhakiHtmlCssKhaki => FromHtml("#FFC3B091");

        public static XLColor LanguidLavender => FromHtml("#FFD6CADD");

        public static XLColor LapisLazuli => FromHtml("#FF26619C");

        public static XLColor LaSalleGreen => FromHtml("#FF087830");

        public static XLColor LaserLemon => FromHtml("#FFFEFE22");

        public static XLColor Lava => FromHtml("#FFCF1020");

        public static XLColor LavenderBlue => FromHtml("#FFCCCCFF");

        public static XLColor LavenderFloral => FromHtml("#FFB57EDC");

        public static XLColor LavenderGray => FromHtml("#FFC4C3D0");

        public static XLColor LavenderIndigo => FromHtml("#FF9457EB");

        public static XLColor LavenderPink => FromHtml("#FFFBAED2");

        public static XLColor LavenderPurple => FromHtml("#FF967BB6");

        public static XLColor LavenderRose => FromHtml("#FFFBA0E3");

        public static XLColor Lemon => FromHtml("#FFFFF700");

        public static XLColor LightApricot => FromHtml("#FFFDD5B1");

        public static XLColor LightBrown => FromHtml("#FFB5651D");

        public static XLColor LightCarminePink => FromHtml("#FFE66771");

        public static XLColor LightCornflowerBlue => FromHtml("#FF93CCEA");

        public static XLColor LightFuchsiaPink => FromHtml("#FFF984EF");

        public static XLColor LightMauve => FromHtml("#FFDCD0FF");

        public static XLColor LightPastelPurple => FromHtml("#FFB19CD9");

        public static XLColor LightSalmonPink => FromHtml("#FFFF9999");

        public static XLColor LightTaupe => FromHtml("#FFB38B6D");

        public static XLColor LightThulianPink => FromHtml("#FFE68FAC");

        public static XLColor LightYellow1 => FromHtml("#FFFFFFED");

        public static XLColor Lilac => FromHtml("#FFC8A2C8");

        public static XLColor LimeColorWheel => FromHtml("#FFBFFF00");

        public static XLColor LincolnGreen => FromHtml("#FF195905");

        public static XLColor Liver => FromHtml("#FF534B4F");

        public static XLColor Lust => FromHtml("#FFE62020");

        public static XLColor MacaroniAndCheese => FromHtml("#FFFFBD88");

        public static XLColor MagentaDye => FromHtml("#FFCA1F7B");

        public static XLColor MagentaProcess => FromHtml("#FFFF0090");

        public static XLColor MagicMint => FromHtml("#FFAAF0D1");

        public static XLColor Magnolia => FromHtml("#FFF8F4FF");

        public static XLColor Mahogany => FromHtml("#FFC04000");

        public static XLColor Maize => FromHtml("#FFFBEC5D");

        public static XLColor MajorelleBlue => FromHtml("#FF6050DC");

        public static XLColor Malachite => FromHtml("#FF0BDA51");

        public static XLColor Manatee => FromHtml("#FF979AAA");

        public static XLColor MangoTango => FromHtml("#FFFF8243");

        public static XLColor MaroonX11 => FromHtml("#FFB03060");

        public static XLColor Mauve => FromHtml("#FFE0B0FF");

        public static XLColor Mauvelous => FromHtml("#FFEF98AA");

        public static XLColor MauveTaupe => FromHtml("#FF915F6D");

        public static XLColor MayaBlue => FromHtml("#FF73C2FB");

        public static XLColor MeatBrown => FromHtml("#FFE5B73B");

        public static XLColor MediumAquamarine1 => FromHtml("#FF66DDAA");

        public static XLColor MediumCandyAppleRed => FromHtml("#FFE2062C");

        public static XLColor MediumCarmine => FromHtml("#FFAF4035");

        public static XLColor MediumChampagne => FromHtml("#FFF3E5AB");

        public static XLColor MediumElectricBlue => FromHtml("#FF035096");

        public static XLColor MediumJungleGreen => FromHtml("#FF1C352D");

        public static XLColor MediumPersianBlue => FromHtml("#FF0067A5");

        public static XLColor MediumRedViolet => FromHtml("#FFBB3385");

        public static XLColor MediumSpringBud => FromHtml("#FFC9DC87");

        public static XLColor MediumTaupe => FromHtml("#FF674C47");

        public static XLColor Melon => FromHtml("#FFFDBCB4");

        public static XLColor MidnightGreenEagleGreen => FromHtml("#FF004953");

        public static XLColor MikadoYellow => FromHtml("#FFFFC40C");

        public static XLColor Mint => FromHtml("#FF3EB489");

        public static XLColor MintGreen => FromHtml("#FF98FF98");

        public static XLColor ModeBeige => FromHtml("#FF967117");

        public static XLColor MoonstoneBlue => FromHtml("#FF73A9C2");

        public static XLColor MordantRed19 => FromHtml("#FFAE0C00");

        public static XLColor MossGreen => FromHtml("#FFADDFAD");

        public static XLColor MountainMeadow => FromHtml("#FF30BA8F");

        public static XLColor MountbattenPink => FromHtml("#FF997A8D");

        public static XLColor MsuGreen => FromHtml("#FF18453B");

        public static XLColor Mulberry => FromHtml("#FFC54B8C");

        public static XLColor Mustard => FromHtml("#FFFFDB58");

        public static XLColor Myrtle => FromHtml("#FF21421E");

        public static XLColor NadeshikoPink => FromHtml("#FFF6ADC6");

        public static XLColor NapierGreen => FromHtml("#FF2A8000");

        public static XLColor NaplesYellow => FromHtml("#FFFADA5E");

        public static XLColor NeonCarrot => FromHtml("#FFFFA343");

        public static XLColor NeonFuchsia => FromHtml("#FFFE59C2");

        public static XLColor NeonGreen => FromHtml("#FF39FF14");

        public static XLColor NonPhotoBlue => FromHtml("#FFA4DDED");

        public static XLColor OceanBoatBlue => FromHtml("#FFCC7422");

        public static XLColor Ochre => FromHtml("#FFCC7722");

        public static XLColor OldGold => FromHtml("#FFCFB53B");

        public static XLColor OldLavender => FromHtml("#FF796878");

        public static XLColor OldMauve => FromHtml("#FF673147");

        public static XLColor OldRose => FromHtml("#FFC08081");

        public static XLColor OliveDrab7 => FromHtml("#FF3C341F");

        public static XLColor Olivine => FromHtml("#FF9AB973");

        public static XLColor Onyx => FromHtml("#FF0F0F0F");

        public static XLColor OperaMauve => FromHtml("#FFB784A7");

        public static XLColor OrangeColorWheel => FromHtml("#FFFF7F00");

        public static XLColor OrangePeel => FromHtml("#FFFF9F00");

        public static XLColor OrangeRyb => FromHtml("#FFFB9902");

        public static XLColor OtterBrown => FromHtml("#FF654321");

        public static XLColor OuCrimsonRed => FromHtml("#FF990000");

        public static XLColor OuterSpace => FromHtml("#FF414A4C");

        public static XLColor OutrageousOrange => FromHtml("#FFFF6E4A");

        public static XLColor OxfordBlue => FromHtml("#FF002147");

        public static XLColor PakistanGreen => FromHtml("#FF00421B");

        public static XLColor PalatinateBlue => FromHtml("#FF273BE2");

        public static XLColor PalatinatePurple => FromHtml("#FF682860");

        public static XLColor PaleAqua => FromHtml("#FFBCD4E6");

        public static XLColor PaleBrown => FromHtml("#FF987654");

        public static XLColor PaleCarmine => FromHtml("#FFAF4035");

        public static XLColor PaleCerulean => FromHtml("#FF9BC4E2");

        public static XLColor PaleChestnut => FromHtml("#FFDDADAF");

        public static XLColor PaleCopper => FromHtml("#FFDA8A67");

        public static XLColor PaleCornflowerBlue => FromHtml("#FFABCDEF");

        public static XLColor PaleGold => FromHtml("#FFE6BE8A");

        public static XLColor PaleMagenta => FromHtml("#FFF984E5");

        public static XLColor PalePink => FromHtml("#FFFADADD");

        public static XLColor PaleRobinEggBlue => FromHtml("#FF96DED1");

        public static XLColor PaleSilver => FromHtml("#FFC9C0BB");

        public static XLColor PaleSpringBud => FromHtml("#FFECEBBD");

        public static XLColor PaleTaupe => FromHtml("#FFBC987E");

        public static XLColor PansyPurple => FromHtml("#FF78184A");

        public static XLColor ParisGreen => FromHtml("#FF50C878");

        public static XLColor PastelBlue => FromHtml("#FFAEC6CF");

        public static XLColor PastelBrown => FromHtml("#FF836953");

        public static XLColor PastelGray => FromHtml("#FFCFCFC4");

        public static XLColor PastelGreen => FromHtml("#FF77DD77");

        public static XLColor PastelMagenta => FromHtml("#FFF49AC2");

        public static XLColor PastelOrange => FromHtml("#FFFFB347");

        public static XLColor PastelPink => FromHtml("#FFFFD1DC");

        public static XLColor PastelPurple => FromHtml("#FFB39EB5");

        public static XLColor PastelRed => FromHtml("#FFFF6961");

        public static XLColor PastelViolet => FromHtml("#FFCB99C9");

        public static XLColor PastelYellow => FromHtml("#FFFDFD96");

        public static XLColor PaynesGrey => FromHtml("#FF40404F");

        public static XLColor Peach => FromHtml("#FFFFE5B4");

        public static XLColor PeachOrange => FromHtml("#FFFFCC99");

        public static XLColor PeachYellow => FromHtml("#FFFADFAD");

        public static XLColor Pear => FromHtml("#FFD1E231");

        public static XLColor Pearl => FromHtml("#FFF0EAD6");

        public static XLColor Peridot => FromHtml("#FFE6E200");

        public static XLColor Periwinkle => FromHtml("#FFCCCCFF");

        public static XLColor PersianBlue => FromHtml("#FF1C39BB");

        public static XLColor PersianGreen => FromHtml("#FF00A693");

        public static XLColor PersianIndigo => FromHtml("#FF32127A");

        public static XLColor PersianOrange => FromHtml("#FFD99058");

        public static XLColor PersianPink => FromHtml("#FFF77FBE");

        public static XLColor PersianPlum => FromHtml("#FF701C1C");

        public static XLColor PersianRed => FromHtml("#FFCC3333");

        public static XLColor PersianRose => FromHtml("#FFFE28A2");

        public static XLColor Persimmon => FromHtml("#FFEC5800");

        public static XLColor Phlox => FromHtml("#FFDF00FF");

        public static XLColor PhthaloBlue => FromHtml("#FF000F89");

        public static XLColor PhthaloGreen => FromHtml("#FF123524");

        public static XLColor PiggyPink => FromHtml("#FFFDDDE6");

        public static XLColor PineGreen => FromHtml("#FF01796F");

        public static XLColor PinkOrange => FromHtml("#FFFF9966");

        public static XLColor PinkPearl => FromHtml("#FFE7ACCF");

        public static XLColor PinkSherbet => FromHtml("#FFF78FA7");

        public static XLColor Pistachio => FromHtml("#FF93C572");

        public static XLColor Platinum => FromHtml("#FFE5E4E2");

        public static XLColor PlumTraditional => FromHtml("#FF8E4585");

        public static XLColor PortlandOrange => FromHtml("#FFFF5A36");

        public static XLColor PrincetonOrange => FromHtml("#FFFF8F00");

        public static XLColor Prune => FromHtml("#FF701C1C");

        public static XLColor PrussianBlue => FromHtml("#FF003153");

        public static XLColor PsychedelicPurple => FromHtml("#FFDF00FF");

        public static XLColor Puce => FromHtml("#FFCC8899");

        public static XLColor Pumpkin => FromHtml("#FFFF7518");

        public static XLColor PurpleHeart => FromHtml("#FF69359C");

        public static XLColor PurpleMountainMajesty => FromHtml("#FF9678B6");

        public static XLColor PurpleMunsell => FromHtml("#FF9F00C5");

        public static XLColor PurplePizzazz => FromHtml("#FFFE4EDA");

        public static XLColor PurpleTaupe => FromHtml("#FF50404D");

        public static XLColor PurpleX11 => FromHtml("#FFA020F0");

        public static XLColor RadicalRed => FromHtml("#FFFF355E");

        public static XLColor Raspberry => FromHtml("#FFE30B5D");

        public static XLColor RaspberryGlace => FromHtml("#FF915F6D");

        public static XLColor RaspberryPink => FromHtml("#FFE25098");

        public static XLColor RaspberryRose => FromHtml("#FFB3446C");

        public static XLColor RawUmber => FromHtml("#FF826644");

        public static XLColor RazzleDazzleRose => FromHtml("#FFFF33CC");

        public static XLColor Razzmatazz => FromHtml("#FFE3256B");

        public static XLColor RedMunsell => FromHtml("#FFF2003C");

        public static XLColor RedNcs => FromHtml("#FFC40233");

        public static XLColor RedPigment => FromHtml("#FFED1C24");

        public static XLColor RedRyb => FromHtml("#FFFE2712");

        public static XLColor Redwood => FromHtml("#FFAB4E52");

        public static XLColor Regalia => FromHtml("#FF522D80");

        public static XLColor RichBlack => FromHtml("#FF004040");

        public static XLColor RichBrilliantLavender => FromHtml("#FFF1A7FE");

        public static XLColor RichCarmine => FromHtml("#FFD70040");

        public static XLColor RichElectricBlue => FromHtml("#FF0892D0");

        public static XLColor RichLavender => FromHtml("#FFA76BCF");

        public static XLColor RichLilac => FromHtml("#FFB666D2");

        public static XLColor RichMaroon => FromHtml("#FFB03060");

        public static XLColor RifleGreen => FromHtml("#FF414833");

        public static XLColor RobinEggBlue => FromHtml("#FF00CCCC");

        public static XLColor Rose => FromHtml("#FFFF007F");

        public static XLColor RoseBonbon => FromHtml("#FFF9429E");

        public static XLColor RoseEbony => FromHtml("#FF674846");

        public static XLColor RoseGold => FromHtml("#FFB76E79");

        public static XLColor RoseMadder => FromHtml("#FFE32636");

        public static XLColor RosePink => FromHtml("#FFFF66CC");

        public static XLColor RoseQuartz => FromHtml("#FFAA98A9");

        public static XLColor RoseTaupe => FromHtml("#FF905D5D");

        public static XLColor RoseVale => FromHtml("#FFAB4E52");

        public static XLColor Rosewood => FromHtml("#FF65000B");

        public static XLColor RossoCorsa => FromHtml("#FFD40000");

        public static XLColor RoyalAzure => FromHtml("#FF0038A8");

        public static XLColor RoyalBlueTraditional => FromHtml("#FF002366");

        public static XLColor RoyalFuchsia => FromHtml("#FFCA2C92");

        public static XLColor RoyalPurple => FromHtml("#FF7851A9");

        public static XLColor Ruby => FromHtml("#FFE0115F");

        public static XLColor Ruddy => FromHtml("#FFFF0028");

        public static XLColor RuddyBrown => FromHtml("#FFBB6528");

        public static XLColor RuddyPink => FromHtml("#FFE18E96");

        public static XLColor Rufous => FromHtml("#FFA81C07");

        public static XLColor Russet => FromHtml("#FF80461B");

        public static XLColor Rust => FromHtml("#FFB7410E");

        public static XLColor SacramentoStateGreen => FromHtml("#FF00563F");

        public static XLColor SafetyOrangeBlazeOrange => FromHtml("#FFFF6700");

        public static XLColor Saffron => FromHtml("#FFF4C430");

        public static XLColor Salmon1 => FromHtml("#FFFF8C69");

        public static XLColor SalmonPink => FromHtml("#FFFF91A4");

        public static XLColor Sand => FromHtml("#FFC2B280");

        public static XLColor SandDune => FromHtml("#FF967117");

        public static XLColor Sandstorm => FromHtml("#FFECD540");

        public static XLColor SandyTaupe => FromHtml("#FF967117");

        public static XLColor Sangria => FromHtml("#FF92000A");

        public static XLColor SapGreen => FromHtml("#FF507D2A");

        public static XLColor Sapphire => FromHtml("#FF082567");

        public static XLColor SatinSheenGold => FromHtml("#FFCBA135");

        public static XLColor Scarlet => FromHtml("#FFFF2000");

        public static XLColor SchoolBusYellow => FromHtml("#FFFFD800");

        public static XLColor ScreaminGreen => FromHtml("#FF76FF7A");

        public static XLColor SealBrown => FromHtml("#FF321414");

        public static XLColor SelectiveYellow => FromHtml("#FFFFBA00");

        public static XLColor Sepia => FromHtml("#FF704214");

        public static XLColor Shadow => FromHtml("#FF8A795D");

        public static XLColor ShamrockGreen => FromHtml("#FF009E60");

        public static XLColor ShockingPink => FromHtml("#FFFC0FC0");

        public static XLColor Sienna1 => FromHtml("#FF882D17");

        public static XLColor Sinopia => FromHtml("#FFCB410B");

        public static XLColor Skobeloff => FromHtml("#FF007474");

        public static XLColor SkyMagenta => FromHtml("#FFCF71AF");

        public static XLColor SmaltDarkPowderBlue => FromHtml("#FF003399");

        public static XLColor SmokeyTopaz => FromHtml("#FF933D41");

        public static XLColor SmokyBlack => FromHtml("#FF100C08");

        public static XLColor SpiroDiscoBall => FromHtml("#FF0FC0FC");

        public static XLColor SplashedWhite => FromHtml("#FFFEFDFF");

        public static XLColor SpringBud => FromHtml("#FFA7FC00");

        public static XLColor StPatricksBlue => FromHtml("#FF23297A");

        public static XLColor StilDeGrainYellow => FromHtml("#FFFADA5E");

        public static XLColor Straw => FromHtml("#FFE4D96F");

        public static XLColor Sunglow => FromHtml("#FFFFCC33");

        public static XLColor Sunset => FromHtml("#FFFAD6A5");

        public static XLColor Tangelo => FromHtml("#FFF94D00");

        public static XLColor Tangerine => FromHtml("#FFF28500");

        public static XLColor TangerineYellow => FromHtml("#FFFFCC00");

        public static XLColor Taupe => FromHtml("#FF483C32");

        public static XLColor TaupeGray => FromHtml("#FF8B8589");

        public static XLColor TeaGreen => FromHtml("#FFD0F0C0");

        public static XLColor TealBlue => FromHtml("#FF367588");

        public static XLColor TealGreen => FromHtml("#FF006D5B");

        public static XLColor TeaRoseOrange => FromHtml("#FFF88379");

        public static XLColor TeaRoseRose => FromHtml("#FFF4C2C2");

        public static XLColor TennéTawny => FromHtml("#FFCD5700");

        public static XLColor TerraCotta => FromHtml("#FFE2725B");

        public static XLColor ThulianPink => FromHtml("#FFDE6FA1");

        public static XLColor TickleMePink => FromHtml("#FFFC89AC");

        public static XLColor TiffanyBlue => FromHtml("#FF0ABAB5");

        public static XLColor TigersEye => FromHtml("#FFE08D3C");

        public static XLColor Timberwolf => FromHtml("#FFDBD7D2");

        public static XLColor TitaniumYellow => FromHtml("#FFEEE600");

        public static XLColor Toolbox => FromHtml("#FF746CC0");

        public static XLColor TractorRed => FromHtml("#FFFD0E35");

        public static XLColor TropicalRainForest => FromHtml("#FF00755E");

        public static XLColor TuftsBlue => FromHtml("#FF417DC1");

        public static XLColor Tumbleweed => FromHtml("#FFDEAA88");

        public static XLColor TurkishRose => FromHtml("#FFB57281");

        public static XLColor Turquoise1 => FromHtml("#FF30D5C8");

        public static XLColor TurquoiseBlue => FromHtml("#FF00FFEF");

        public static XLColor TurquoiseGreen => FromHtml("#FFA0D6B4");

        public static XLColor TuscanRed => FromHtml("#FF823535");

        public static XLColor TwilightLavender => FromHtml("#FF8A496B");

        public static XLColor TyrianPurple => FromHtml("#FF66023C");

        public static XLColor UaBlue => FromHtml("#FF0033AA");

        public static XLColor UaRed => FromHtml("#FFD9004C");

        public static XLColor Ube => FromHtml("#FF8878C3");

        public static XLColor UclaBlue => FromHtml("#FF536895");

        public static XLColor UclaGold => FromHtml("#FFFFB300");

        public static XLColor UfoGreen => FromHtml("#FF3CD070");

        public static XLColor Ultramarine => FromHtml("#FF120A8F");

        public static XLColor UltramarineBlue => FromHtml("#FF4166F5");

        public static XLColor UltraPink => FromHtml("#FFFF6FFF");

        public static XLColor Umber => FromHtml("#FF635147");

        public static XLColor UnitedNationsBlue => FromHtml("#FF5B92E5");

        public static XLColor UnmellowYellow => FromHtml("#FFFFFF66");

        public static XLColor UpForestGreen => FromHtml("#FF014421");

        public static XLColor UpMaroon => FromHtml("#FF7B1113");

        public static XLColor UpsdellRed => FromHtml("#FFAE2029");

        public static XLColor Urobilin => FromHtml("#FFE1AD21");

        public static XLColor UscCardinal => FromHtml("#FF990000");

        public static XLColor UscGold => FromHtml("#FFFFCC00");

        public static XLColor UtahCrimson => FromHtml("#FFD3003F");

        public static XLColor Vanilla => FromHtml("#FFF3E5AB");

        public static XLColor VegasGold => FromHtml("#FFC5B358");

        public static XLColor VenetianRed => FromHtml("#FFC80815");

        public static XLColor Verdigris => FromHtml("#FF43B3AE");

        public static XLColor Vermilion => FromHtml("#FFE34234");

        public static XLColor Veronica => FromHtml("#FFA020F0");

        public static XLColor Violet1 => FromHtml("#FF8F00FF");

        public static XLColor VioletColorWheel => FromHtml("#FF7F00FF");

        public static XLColor VioletRyb => FromHtml("#FF8601AF");

        public static XLColor Viridian => FromHtml("#FF40826D");

        public static XLColor VividAuburn => FromHtml("#FF922724");

        public static XLColor VividBurgundy => FromHtml("#FF9F1D35");

        public static XLColor VividCerise => FromHtml("#FFDA1D81");

        public static XLColor VividTangerine => FromHtml("#FFFFA089");

        public static XLColor VividViolet => FromHtml("#FF9F00FF");

        public static XLColor WarmBlack => FromHtml("#FF004242");

        public static XLColor Wenge => FromHtml("#FF645452");

        public static XLColor WildBlueYonder => FromHtml("#FFA2ADD0");

        public static XLColor WildStrawberry => FromHtml("#FFFF43A4");

        public static XLColor WildWatermelon => FromHtml("#FFFC6C85");

        public static XLColor Wisteria => FromHtml("#FFC9A0DC");

        public static XLColor Xanadu => FromHtml("#FF738678");

        public static XLColor YaleBlue => FromHtml("#FF0F4D92");

        public static XLColor YellowMunsell => FromHtml("#FFEFCC00");

        public static XLColor YellowNcs => FromHtml("#FFFFD300");

        public static XLColor YellowProcess => FromHtml("#FFFFEF00");

        public static XLColor YellowRyb => FromHtml("#FFFEFE33");

        public static XLColor Zaffre => FromHtml("#FF0014A8");

        public static XLColor ZinnwalditeBrown => FromHtml("#FF2C1608");

        public static XLColor Transparent => FromColor(SKColors.Transparent);
    }
}
