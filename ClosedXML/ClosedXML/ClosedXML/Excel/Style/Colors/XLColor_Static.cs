using System;
using System.Collections.Generic;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        public static IXLColor FromColor(Color color)
        {
            return new XLColor(color);
        }
        public static IXLColor FromArgb(Int32 argb)
        {
            return new XLColor(Color.FromArgb(argb));
        }
        public static IXLColor FromArgb(Int32 r, Int32 g, Int32 b)
        {
            return new XLColor(Color.FromArgb(r, g, b));
        }
        public static IXLColor FromArgb(Int32 a, Int32 r, Int32 g, Int32 b)
        {
            return new XLColor(Color.FromArgb(a, r, g, b));
        }
        public static IXLColor FromKnownColor(KnownColor color)
        {
            return new XLColor(Color.FromKnownColor(color));
        }
        public static IXLColor FromName(String name)
        {
            return new XLColor(Color.FromName(name));
        }
        public static IXLColor FromHtml(String htmlColor)
        {
            return new XLColor(ColorTranslator.FromHtml(htmlColor));
        }
        public static IXLColor FromIndex(Int32 index)
        {
            return new XLColor(index);
        }
        public static IXLColor FromTheme(XLThemeColor themeColor)
        {
            return new XLColor(themeColor);
        }
        public static IXLColor FromTheme(XLThemeColor themeColor, Double themeTint)
        {
            return new XLColor(themeColor, themeTint);
        }

        private static Dictionary<Int32, IXLColor> indexedColors;
        public static Dictionary<Int32, IXLColor> IndexedColors
        {
            get
            {
                if (indexedColors == null)
                {
                    Dictionary<Int32, IXLColor> retVal = new Dictionary<Int32, IXLColor>();
                    retVal.Add(0, XLColor.FromHtml("#FF000000"));
                    retVal.Add(1, XLColor.FromHtml("#FFFFFFFF"));
                    retVal.Add(2, XLColor.FromHtml("#FFFF0000"));
                    retVal.Add(3, XLColor.FromHtml("#FF00FF00"));
                    retVal.Add(4, XLColor.FromHtml("#FF0000FF"));
                    retVal.Add(5, XLColor.FromHtml("#FFFFFF00"));
                    retVal.Add(6, XLColor.FromHtml("#FFFF00FF"));
                    retVal.Add(7, XLColor.FromHtml("#FF00FFFF"));
                    retVal.Add(8, XLColor.FromHtml("#FF000000"));
                    retVal.Add(9, XLColor.FromHtml("#FFFFFFFF"));
                    retVal.Add(10, XLColor.FromHtml("#FFFF0000"));
                    retVal.Add(11, XLColor.FromHtml("#FF00FF00"));
                    retVal.Add(12, XLColor.FromHtml("#FF0000FF"));
                    retVal.Add(13, XLColor.FromHtml("#FFFFFF00"));
                    retVal.Add(14, XLColor.FromHtml("#FFFF00FF"));
                    retVal.Add(15, XLColor.FromHtml("#FF00FFFF"));
                    retVal.Add(16, XLColor.FromHtml("#FF800000"));
                    retVal.Add(17, XLColor.FromHtml("#FF008000"));
                    retVal.Add(18, XLColor.FromHtml("#FF000080"));
                    retVal.Add(19, XLColor.FromHtml("#FF808000"));
                    retVal.Add(20, XLColor.FromHtml("#FF800080"));
                    retVal.Add(21, XLColor.FromHtml("#FF008080"));
                    retVal.Add(22, XLColor.FromHtml("#FFC0C0C0"));
                    retVal.Add(23, XLColor.FromHtml("#FF808080"));
                    retVal.Add(24, XLColor.FromHtml("#FF9999FF"));
                    retVal.Add(25, XLColor.FromHtml("#FF993366"));
                    retVal.Add(26, XLColor.FromHtml("#FFFFFFCC"));
                    retVal.Add(27, XLColor.FromHtml("#FFCCFFFF"));
                    retVal.Add(28, XLColor.FromHtml("#FF660066"));
                    retVal.Add(29, XLColor.FromHtml("#FFFF8080"));
                    retVal.Add(30, XLColor.FromHtml("#FF0066CC"));
                    retVal.Add(31, XLColor.FromHtml("#FFCCCCFF"));
                    retVal.Add(32, XLColor.FromHtml("#FF000080"));
                    retVal.Add(33, XLColor.FromHtml("#FFFF00FF"));
                    retVal.Add(34, XLColor.FromHtml("#FFFFFF00"));
                    retVal.Add(35, XLColor.FromHtml("#FF00FFFF"));
                    retVal.Add(36, XLColor.FromHtml("#FF800080"));
                    retVal.Add(37, XLColor.FromHtml("#FF800000"));
                    retVal.Add(38, XLColor.FromHtml("#FF008080"));
                    retVal.Add(39, XLColor.FromHtml("#FF0000FF"));
                    retVal.Add(40, XLColor.FromHtml("#FF00CCFF"));
                    retVal.Add(41, XLColor.FromHtml("#FFCCFFFF"));
                    retVal.Add(42, XLColor.FromHtml("#FFCCFFCC"));
                    retVal.Add(43, XLColor.FromHtml("#FFFFFF99"));
                    retVal.Add(44, XLColor.FromHtml("#FF99CCFF"));
                    retVal.Add(45, XLColor.FromHtml("#FFFF99CC"));
                    retVal.Add(46, XLColor.FromHtml("#FFCC99FF"));
                    retVal.Add(47, XLColor.FromHtml("#FFFFCC99"));
                    retVal.Add(48, XLColor.FromHtml("#FF3366FF"));
                    retVal.Add(49, XLColor.FromHtml("#FF33CCCC"));
                    retVal.Add(50, XLColor.FromHtml("#FF003300"));
                    retVal.Add(51, XLColor.FromHtml("#FF99CC00"));
                    retVal.Add(52, XLColor.FromHtml("#FFFFCC00"));
                    retVal.Add(53, XLColor.FromHtml("#FFFF9900"));
                    retVal.Add(54, XLColor.FromHtml("#FFFF6600"));
                    retVal.Add(55, XLColor.FromHtml("#FF666699"));
                    retVal.Add(56, XLColor.FromHtml("#FF969696"));
                    retVal.Add(57, XLColor.FromHtml("#FF003366"));
                    retVal.Add(58, XLColor.FromHtml("#FF339966"));
                    retVal.Add(59, XLColor.FromHtml("#FF333300"));
                    retVal.Add(60, XLColor.FromHtml("#FF993300"));
                    retVal.Add(61, XLColor.FromHtml("#FF993366"));
                    retVal.Add(62, XLColor.FromHtml("#FF333399"));
                    retVal.Add(63, XLColor.FromHtml("#FF333333"));
                    indexedColors = retVal;
                }
                return indexedColors;
            }
        }

        public static IXLColor NoColor { get { return new XLColor(); } }

        public static IXLColor AliceBlue { get { return FromColor(Color.AliceBlue); } }
        public static IXLColor AntiqueWhite { get { return FromColor(Color.AntiqueWhite); } }
        public static IXLColor Aqua { get { return FromColor(Color.Aqua); } }
        public static IXLColor Aquamarine { get { return FromColor(Color.Aquamarine); } }
        public static IXLColor Azure { get { return FromColor(Color.Azure); } }
        public static IXLColor Beige { get { return FromColor(Color.Beige); } }
        public static IXLColor Bisque { get { return FromColor(Color.Bisque); } }
        public static IXLColor Black { get { return FromColor(Color.Black); } }
        public static IXLColor BlanchedAlmond { get { return FromColor(Color.BlanchedAlmond); } }
        public static IXLColor Blue { get { return FromColor(Color.Blue); } }
        public static IXLColor BlueViolet { get { return FromColor(Color.BlueViolet); } }
        public static IXLColor Brown { get { return FromColor(Color.Brown); } }
        public static IXLColor BurlyWood { get { return FromColor(Color.BurlyWood); } }
        public static IXLColor CadetBlue { get { return FromColor(Color.CadetBlue); } }
        public static IXLColor Chartreuse { get { return FromColor(Color.Chartreuse); } }
        public static IXLColor Chocolate { get { return FromColor(Color.Chocolate); } }
        public static IXLColor Coral { get { return FromColor(Color.Coral); } }
        public static IXLColor CornflowerBlue { get { return FromColor(Color.CornflowerBlue); } }
        public static IXLColor Cornsilk { get { return FromColor(Color.Cornsilk); } }
        public static IXLColor Crimson { get { return FromColor(Color.Crimson); } }
        public static IXLColor Cyan { get { return FromColor(Color.Cyan); } }
        public static IXLColor DarkBlue { get { return FromColor(Color.DarkBlue); } }
        public static IXLColor DarkCyan { get { return FromColor(Color.DarkCyan); } }
        public static IXLColor DarkGoldenrod { get { return FromColor(Color.DarkGoldenrod); } }
        public static IXLColor DarkGray { get { return FromColor(Color.DarkGray); } }
        public static IXLColor DarkGreen { get { return FromColor(Color.DarkGreen); } }
        public static IXLColor DarkKhaki { get { return FromColor(Color.DarkKhaki); } }
        public static IXLColor DarkMagenta { get { return FromColor(Color.DarkMagenta); } }
        public static IXLColor DarkOliveGreen { get { return FromColor(Color.DarkOliveGreen); } }
        public static IXLColor DarkOrange { get { return FromColor(Color.DarkOrange); } }
        public static IXLColor DarkOrchid { get { return FromColor(Color.DarkOrchid); } }
        public static IXLColor DarkRed { get { return FromColor(Color.DarkRed); } }
        public static IXLColor DarkSalmon { get { return FromColor(Color.DarkSalmon); } }
        public static IXLColor DarkSeaGreen { get { return FromColor(Color.DarkSeaGreen); } }
        public static IXLColor DarkSlateBlue { get { return FromColor(Color.DarkSlateBlue); } }
        public static IXLColor DarkSlateGray { get { return FromColor(Color.DarkSlateGray); } }
        public static IXLColor DarkTurquoise { get { return FromColor(Color.DarkTurquoise); } }
        public static IXLColor DarkViolet { get { return FromColor(Color.DarkViolet); } }
        public static IXLColor DeepPink { get { return FromColor(Color.DeepPink); } }
        public static IXLColor DeepSkyBlue { get { return FromColor(Color.DeepSkyBlue); } }
        public static IXLColor DimGray { get { return FromColor(Color.DimGray); } }
        public static IXLColor DodgerBlue { get { return FromColor(Color.DodgerBlue); } }
        public static IXLColor Firebrick { get { return FromColor(Color.Firebrick); } }
        public static IXLColor FloralWhite { get { return FromColor(Color.FloralWhite); } }
        public static IXLColor ForestGreen { get { return FromColor(Color.ForestGreen); } }
        public static IXLColor Fuchsia { get { return FromColor(Color.Fuchsia); } }
        public static IXLColor Gainsboro { get { return FromColor(Color.Gainsboro); } }
        public static IXLColor GhostWhite { get { return FromColor(Color.GhostWhite); } }
        public static IXLColor Gold { get { return FromColor(Color.Gold); } }
        public static IXLColor Goldenrod { get { return FromColor(Color.Goldenrod); } }
        public static IXLColor Gray { get { return FromColor(Color.Gray); } }
        public static IXLColor Green { get { return FromColor(Color.Green); } }
        public static IXLColor GreenYellow { get { return FromColor(Color.GreenYellow); } }
        public static IXLColor Honeydew { get { return FromColor(Color.Honeydew); } }
        public static IXLColor HotPink { get { return FromColor(Color.HotPink); } }
        public static IXLColor IndianRed { get { return FromColor(Color.IndianRed); } }
        public static IXLColor Indigo { get { return FromColor(Color.Indigo); } }
        public static IXLColor Ivory { get { return FromColor(Color.Ivory); } }
        public static IXLColor Khaki { get { return FromColor(Color.Khaki); } }
        public static IXLColor Lavender { get { return FromColor(Color.Lavender); } }
        public static IXLColor LavenderBlush { get { return FromColor(Color.LavenderBlush); } }
        public static IXLColor LawnGreen { get { return FromColor(Color.LawnGreen); } }
        public static IXLColor LemonChiffon { get { return FromColor(Color.LemonChiffon); } }
        public static IXLColor LightBlue { get { return FromColor(Color.LightBlue); } }
        public static IXLColor LightCoral { get { return FromColor(Color.LightCoral); } }
        public static IXLColor LightCyan { get { return FromColor(Color.LightCyan); } }
        public static IXLColor LightGoldenrodYellow { get { return FromColor(Color.LightGoldenrodYellow); } }
        public static IXLColor LightGray { get { return FromColor(Color.LightGray); } }
        public static IXLColor LightGreen { get { return FromColor(Color.LightGreen); } }
        public static IXLColor LightPink { get { return FromColor(Color.LightPink); } }
        public static IXLColor LightSalmon { get { return FromColor(Color.LightSalmon); } }
        public static IXLColor LightSeaGreen { get { return FromColor(Color.LightSeaGreen); } }
        public static IXLColor LightSkyBlue { get { return FromColor(Color.LightSkyBlue); } }
        public static IXLColor LightSlateGray { get { return FromColor(Color.LightSlateGray); } }
        public static IXLColor LightSteelBlue { get { return FromColor(Color.LightSteelBlue); } }
        public static IXLColor LightYellow { get { return FromColor(Color.LightYellow); } }
        public static IXLColor Lime { get { return FromColor(Color.Lime); } }
        public static IXLColor LimeGreen { get { return FromColor(Color.LimeGreen); } }
        public static IXLColor Linen { get { return FromColor(Color.Linen); } }
        public static IXLColor Magenta { get { return FromColor(Color.Magenta); } }
        public static IXLColor Maroon { get { return FromColor(Color.Maroon); } }
        public static IXLColor MediumAquamarine { get { return FromColor(Color.MediumAquamarine); } }
        public static IXLColor MediumBlue { get { return FromColor(Color.MediumBlue); } }
        public static IXLColor MediumOrchid { get { return FromColor(Color.MediumOrchid); } }
        public static IXLColor MediumPurple { get { return FromColor(Color.MediumPurple); } }
        public static IXLColor MediumSeaGreen { get { return FromColor(Color.MediumSeaGreen); } }
        public static IXLColor MediumSlateBlue { get { return FromColor(Color.MediumSlateBlue); } }
        public static IXLColor MediumSpringGreen { get { return FromColor(Color.MediumSpringGreen); } }
        public static IXLColor MediumTurquoise { get { return FromColor(Color.MediumTurquoise); } }
        public static IXLColor MediumVioletRed { get { return FromColor(Color.MediumVioletRed); } }
        public static IXLColor MidnightBlue { get { return FromColor(Color.MidnightBlue); } }
        public static IXLColor MintCream { get { return FromColor(Color.MintCream); } }
        public static IXLColor MistyRose { get { return FromColor(Color.MistyRose); } }
        public static IXLColor Moccasin { get { return FromColor(Color.Moccasin); } }
        public static IXLColor NavajoWhite { get { return FromColor(Color.NavajoWhite); } }
        public static IXLColor Navy { get { return FromColor(Color.Navy); } }
        public static IXLColor OldLace { get { return FromColor(Color.OldLace); } }
        public static IXLColor Olive { get { return FromColor(Color.Olive); } }
        public static IXLColor OliveDrab { get { return FromColor(Color.OliveDrab); } }
        public static IXLColor Orange { get { return FromColor(Color.Orange); } }
        public static IXLColor OrangeRed { get { return FromColor(Color.OrangeRed); } }
        public static IXLColor Orchid { get { return FromColor(Color.Orchid); } }
        public static IXLColor PaleGoldenrod { get { return FromColor(Color.PaleGoldenrod); } }
        public static IXLColor PaleGreen { get { return FromColor(Color.PaleGreen); } }
        public static IXLColor PaleTurquoise { get { return FromColor(Color.PaleTurquoise); } }
        public static IXLColor PaleVioletRed { get { return FromColor(Color.PaleVioletRed); } }
        public static IXLColor PapayaWhip { get { return FromColor(Color.PapayaWhip); } }
        public static IXLColor PeachPuff { get { return FromColor(Color.PeachPuff); } }
        public static IXLColor Peru { get { return FromColor(Color.Peru); } }
        public static IXLColor Pink { get { return FromColor(Color.Pink); } }
        public static IXLColor Plum { get { return FromColor(Color.Plum); } }
        public static IXLColor PowderBlue { get { return FromColor(Color.PowderBlue); } }
        public static IXLColor Purple { get { return FromColor(Color.Purple); } }
        public static IXLColor Red { get { return FromColor(Color.Red); } }
        public static IXLColor RosyBrown { get { return FromColor(Color.RosyBrown); } }
        public static IXLColor RoyalBlue { get { return FromColor(Color.RoyalBlue); } }
        public static IXLColor SaddleBrown { get { return FromColor(Color.SaddleBrown); } }
        public static IXLColor Salmon { get { return FromColor(Color.Salmon); } }
        public static IXLColor SandyBrown { get { return FromColor(Color.SandyBrown); } }
        public static IXLColor SeaGreen { get { return FromColor(Color.SeaGreen); } }
        public static IXLColor SeaShell { get { return FromColor(Color.SeaShell); } }
        public static IXLColor Sienna { get { return FromColor(Color.Sienna); } }
        public static IXLColor Silver { get { return FromColor(Color.Silver); } }
        public static IXLColor SkyBlue { get { return FromColor(Color.SkyBlue); } }
        public static IXLColor SlateBlue { get { return FromColor(Color.SlateBlue); } }
        public static IXLColor SlateGray { get { return FromColor(Color.SlateGray); } }
        public static IXLColor Snow { get { return FromColor(Color.Snow); } }
        public static IXLColor SpringGreen { get { return FromColor(Color.SpringGreen); } }
        public static IXLColor SteelBlue { get { return FromColor(Color.SteelBlue); } }
        public static IXLColor Tan { get { return FromColor(Color.Tan); } }
        public static IXLColor Teal { get { return FromColor(Color.Teal); } }
        public static IXLColor Thistle { get { return FromColor(Color.Thistle); } }
        public static IXLColor Tomato { get { return FromColor(Color.Tomato); } }
        public static IXLColor Turquoise { get { return FromColor(Color.Turquoise); } }
        public static IXLColor Violet { get { return FromColor(Color.Violet); } }
        public static IXLColor Wheat { get { return FromColor(Color.Wheat); } }
        public static IXLColor White { get { return FromColor(Color.White); } }
        public static IXLColor WhiteSmoke { get { return FromColor(Color.WhiteSmoke); } }
        public static IXLColor Yellow { get { return FromColor(Color.Yellow); } }
        public static IXLColor YellowGreen { get { return FromColor(Color.YellowGreen); } }
        public static IXLColor AirForceBlue { get { return FromHtml("#FF5D8AA8"); } }
        public static IXLColor Alizarin { get { return FromHtml("#FFE32636"); } }
        public static IXLColor Almond { get { return FromHtml("#FFEFDECD"); } }
        public static IXLColor Amaranth { get { return FromHtml("#FFE52B50"); } }
        public static IXLColor Amber { get { return FromHtml("#FFFFBF00"); } }
        public static IXLColor AmberSaeEce { get { return FromHtml("#FFFF7E00"); } }
        public static IXLColor AmericanRose { get { return FromHtml("#FFFF033E"); } }
        public static IXLColor Amethyst { get { return FromHtml("#FF9966CC"); } }
        public static IXLColor AntiFlashWhite { get { return FromHtml("#FFF2F3F4"); } }
        public static IXLColor AntiqueBrass { get { return FromHtml("#FFCD9575"); } }
        public static IXLColor AntiqueFuchsia { get { return FromHtml("#FF915C83"); } }
        public static IXLColor AppleGreen { get { return FromHtml("#FF8DB600"); } }
        public static IXLColor Apricot { get { return FromHtml("#FFFBCEB1"); } }
        public static IXLColor Aquamarine1 { get { return FromHtml("#FF7FFFD0"); } }
        public static IXLColor ArmyGreen { get { return FromHtml("#FF4B5320"); } }
        public static IXLColor Arsenic { get { return FromHtml("#FF3B444B"); } }
        public static IXLColor ArylideYellow { get { return FromHtml("#FFE9D66B"); } }
        public static IXLColor AshGrey { get { return FromHtml("#FFB2BEB5"); } }
        public static IXLColor Asparagus { get { return FromHtml("#FF87A96B"); } }
        public static IXLColor AtomicTangerine { get { return FromHtml("#FFFF9966"); } }
        public static IXLColor Auburn { get { return FromHtml("#FF6D351A"); } }
        public static IXLColor Aureolin { get { return FromHtml("#FFFDEE00"); } }
        public static IXLColor Aurometalsaurus { get { return FromHtml("#FF6E7F80"); } }
        public static IXLColor Awesome { get { return FromHtml("#FFFF2052"); } }
        public static IXLColor AzureColorWheel { get { return FromHtml("#FF007FFF"); } }
        public static IXLColor BabyBlue { get { return FromHtml("#FF89CFF0"); } }
        public static IXLColor BabyBlueEyes { get { return FromHtml("#FFA1CAF1"); } }
        public static IXLColor BabyPink { get { return FromHtml("#FFF4C2C2"); } }
        public static IXLColor BallBlue { get { return FromHtml("#FF21ABCD"); } }
        public static IXLColor BananaMania { get { return FromHtml("#FFFAE7B5"); } }
        public static IXLColor BattleshipGrey { get { return FromHtml("#FF848482"); } }
        public static IXLColor Bazaar { get { return FromHtml("#FF98777B"); } }
        public static IXLColor BeauBlue { get { return FromHtml("#FFBCD4E6"); } }
        public static IXLColor Beaver { get { return FromHtml("#FF9F8170"); } }
        public static IXLColor Bistre { get { return FromHtml("#FF3D2B1F"); } }
        public static IXLColor Bittersweet { get { return FromHtml("#FFFE6F5E"); } }
        public static IXLColor BleuDeFrance { get { return FromHtml("#FF318CE7"); } }
        public static IXLColor BlizzardBlue { get { return FromHtml("#FFACE5EE"); } }
        public static IXLColor Blond { get { return FromHtml("#FFFAF0BE"); } }
        public static IXLColor BlueBell { get { return FromHtml("#FFA2A2D0"); } }
        public static IXLColor BlueGray { get { return FromHtml("#FF6699CC"); } }
        public static IXLColor BlueGreen { get { return FromHtml("#FF00DDDD"); } }
        public static IXLColor BluePigment { get { return FromHtml("#FF333399"); } }
        public static IXLColor BlueRyb { get { return FromHtml("#FF0247FE"); } }
        public static IXLColor Blush { get { return FromHtml("#FFDE5D83"); } }
        public static IXLColor Bole { get { return FromHtml("#FF79443B"); } }
        public static IXLColor BondiBlue { get { return FromHtml("#FF0095B6"); } }
        public static IXLColor BostonUniversityRed { get { return FromHtml("#FFCC0000"); } }
        public static IXLColor BrandeisBlue { get { return FromHtml("#FF0070FF"); } }
        public static IXLColor Brass { get { return FromHtml("#FFB5A642"); } }
        public static IXLColor BrickRed { get { return FromHtml("#FFCB4154"); } }
        public static IXLColor BrightCerulean { get { return FromHtml("#FF1DACD6"); } }
        public static IXLColor BrightGreen { get { return FromHtml("#FF66FF00"); } }
        public static IXLColor BrightLavender { get { return FromHtml("#FFBF94E4"); } }
        public static IXLColor BrightMaroon { get { return FromHtml("#FFC32148"); } }
        public static IXLColor BrightPink { get { return FromHtml("#FFFF007F"); } }
        public static IXLColor BrightTurquoise { get { return FromHtml("#FF08E8DE"); } }
        public static IXLColor BrightUbe { get { return FromHtml("#FFD19FE8"); } }
        public static IXLColor BrilliantLavender { get { return FromHtml("#FFF4BBFF"); } }
        public static IXLColor BrilliantRose { get { return FromHtml("#FFFF55A3"); } }
        public static IXLColor BrinkPink { get { return FromHtml("#FFFB607F"); } }
        public static IXLColor BritishRacingGreen { get { return FromHtml("#FF004225"); } }
        public static IXLColor Bronze { get { return FromHtml("#FFCD7F32"); } }
        public static IXLColor BrownTraditional { get { return FromHtml("#FF964B00"); } }
        public static IXLColor BubbleGum { get { return FromHtml("#FFFFC1CC"); } }
        public static IXLColor Bubbles { get { return FromHtml("#FFE7FEFF"); } }
        public static IXLColor Buff { get { return FromHtml("#FFF0DC82"); } }
        public static IXLColor BulgarianRose { get { return FromHtml("#FF480607"); } }
        public static IXLColor Burgundy { get { return FromHtml("#FF800020"); } }
        public static IXLColor BurntOrange { get { return FromHtml("#FFCC5500"); } }
        public static IXLColor BurntSienna { get { return FromHtml("#FFE97451"); } }
        public static IXLColor BurntUmber { get { return FromHtml("#FF8A3324"); } }
        public static IXLColor Byzantine { get { return FromHtml("#FFBD33A4"); } }
        public static IXLColor Byzantium { get { return FromHtml("#FF702963"); } }
        public static IXLColor Cadet { get { return FromHtml("#FF536872"); } }
        public static IXLColor CadetGrey { get { return FromHtml("#FF91A3B0"); } }
        public static IXLColor CadmiumGreen { get { return FromHtml("#FF006B3C"); } }
        public static IXLColor CadmiumOrange { get { return FromHtml("#FFED872D"); } }
        public static IXLColor CadmiumRed { get { return FromHtml("#FFE30022"); } }
        public static IXLColor CadmiumYellow { get { return FromHtml("#FFFFF600"); } }
        public static IXLColor CalPolyPomonaGreen { get { return FromHtml("#FF1E4D2B"); } }
        public static IXLColor CambridgeBlue { get { return FromHtml("#FFA3C1AD"); } }
        public static IXLColor Camel { get { return FromHtml("#FFC19A6B"); } }
        public static IXLColor CamouflageGreen { get { return FromHtml("#FF78866B"); } }
        public static IXLColor CanaryYellow { get { return FromHtml("#FFFFEF00"); } }
        public static IXLColor CandyAppleRed { get { return FromHtml("#FFFF0800"); } }
        public static IXLColor CandyPink { get { return FromHtml("#FFE4717A"); } }
        public static IXLColor CaputMortuum { get { return FromHtml("#FF592720"); } }
        public static IXLColor Cardinal { get { return FromHtml("#FFC41E3A"); } }
        public static IXLColor CaribbeanGreen { get { return FromHtml("#FF00CC99"); } }
        public static IXLColor Carmine { get { return FromHtml("#FF960018"); } }
        public static IXLColor CarminePink { get { return FromHtml("#FFEB4C42"); } }
        public static IXLColor CarmineRed { get { return FromHtml("#FFFF0038"); } }
        public static IXLColor CarnationPink { get { return FromHtml("#FFFFA6C9"); } }
        public static IXLColor Carnelian { get { return FromHtml("#FFB31B1B"); } }
        public static IXLColor CarolinaBlue { get { return FromHtml("#FF99BADD"); } }
        public static IXLColor CarrotOrange { get { return FromHtml("#FFED9121"); } }
        public static IXLColor Ceil { get { return FromHtml("#FF92A1CF"); } }
        public static IXLColor Celadon { get { return FromHtml("#FFACE1AF"); } }
        public static IXLColor CelestialBlue { get { return FromHtml("#FF4997D0"); } }
        public static IXLColor Cerise { get { return FromHtml("#FFDE3163"); } }
        public static IXLColor CerisePink { get { return FromHtml("#FFEC3B83"); } }
        public static IXLColor Cerulean { get { return FromHtml("#FF007BA7"); } }
        public static IXLColor CeruleanBlue { get { return FromHtml("#FF2A52BE"); } }
        public static IXLColor Chamoisee { get { return FromHtml("#FFA0785A"); } }
        public static IXLColor Champagne { get { return FromHtml("#FFF7E7CE"); } }
        public static IXLColor Charcoal { get { return FromHtml("#FF36454F"); } }
        public static IXLColor ChartreuseTraditional { get { return FromHtml("#FFDFFF00"); } }
        public static IXLColor CherryBlossomPink { get { return FromHtml("#FFFFB7C5"); } }
        public static IXLColor Chocolate1 { get { return FromHtml("#FF7B3F00"); } }
        public static IXLColor ChromeYellow { get { return FromHtml("#FFFFA700"); } }
        public static IXLColor Cinereous { get { return FromHtml("#FF98817B"); } }
        public static IXLColor Cinnabar { get { return FromHtml("#FFE34234"); } }
        public static IXLColor Citrine { get { return FromHtml("#FFE4D00A"); } }
        public static IXLColor ClassicRose { get { return FromHtml("#FFFBCCE7"); } }
        public static IXLColor Cobalt { get { return FromHtml("#FF0047AB"); } }
        public static IXLColor ColumbiaBlue { get { return FromHtml("#FF9BDDFF"); } }
        public static IXLColor CoolBlack { get { return FromHtml("#FF002E63"); } }
        public static IXLColor CoolGrey { get { return FromHtml("#FF8C92AC"); } }
        public static IXLColor Copper { get { return FromHtml("#FFB87333"); } }
        public static IXLColor CopperRose { get { return FromHtml("#FF996666"); } }
        public static IXLColor Coquelicot { get { return FromHtml("#FFFF3800"); } }
        public static IXLColor CoralPink { get { return FromHtml("#FFF88379"); } }
        public static IXLColor CoralRed { get { return FromHtml("#FFFF4040"); } }
        public static IXLColor Cordovan { get { return FromHtml("#FF893F45"); } }
        public static IXLColor Corn { get { return FromHtml("#FFFBEC5D"); } }
        public static IXLColor CornellRed { get { return FromHtml("#FFB31B1B"); } }
        public static IXLColor CosmicLatte { get { return FromHtml("#FFFFF8E7"); } }
        public static IXLColor CottonCandy { get { return FromHtml("#FFFFBCD9"); } }
        public static IXLColor Cream { get { return FromHtml("#FFFFFDD0"); } }
        public static IXLColor CrimsonGlory { get { return FromHtml("#FFBE0032"); } }
        public static IXLColor CyanProcess { get { return FromHtml("#FF00B7EB"); } }
        public static IXLColor Daffodil { get { return FromHtml("#FFFFFF31"); } }
        public static IXLColor Dandelion { get { return FromHtml("#FFF0E130"); } }
        public static IXLColor DarkBrown { get { return FromHtml("#FF654321"); } }
        public static IXLColor DarkByzantium { get { return FromHtml("#FF5D3954"); } }
        public static IXLColor DarkCandyAppleRed { get { return FromHtml("#FFA40000"); } }
        public static IXLColor DarkCerulean { get { return FromHtml("#FF08457E"); } }
        public static IXLColor DarkChampagne { get { return FromHtml("#FFC2B280"); } }
        public static IXLColor DarkChestnut { get { return FromHtml("#FF986960"); } }
        public static IXLColor DarkCoral { get { return FromHtml("#FFCD5B45"); } }
        public static IXLColor DarkElectricBlue { get { return FromHtml("#FF536878"); } }
        public static IXLColor DarkGreen1 { get { return FromHtml("#FF013220"); } }
        public static IXLColor DarkJungleGreen { get { return FromHtml("#FF1A2421"); } }
        public static IXLColor DarkLava { get { return FromHtml("#FF483C32"); } }
        public static IXLColor DarkLavender { get { return FromHtml("#FF734F96"); } }
        public static IXLColor DarkMidnightBlue { get { return FromHtml("#FF003366"); } }
        public static IXLColor DarkPastelBlue { get { return FromHtml("#FF779ECB"); } }
        public static IXLColor DarkPastelGreen { get { return FromHtml("#FF03C03C"); } }
        public static IXLColor DarkPastelPurple { get { return FromHtml("#FF966FD6"); } }
        public static IXLColor DarkPastelRed { get { return FromHtml("#FFC23B22"); } }
        public static IXLColor DarkPink { get { return FromHtml("#FFE75480"); } }
        public static IXLColor DarkPowderBlue { get { return FromHtml("#FF003399"); } }
        public static IXLColor DarkRaspberry { get { return FromHtml("#FF872657"); } }
        public static IXLColor DarkScarlet { get { return FromHtml("#FF560319"); } }
        public static IXLColor DarkSienna { get { return FromHtml("#FF3C1414"); } }
        public static IXLColor DarkSpringGreen { get { return FromHtml("#FF177245"); } }
        public static IXLColor DarkTan { get { return FromHtml("#FF918151"); } }
        public static IXLColor DarkTangerine { get { return FromHtml("#FFFFA812"); } }
        public static IXLColor DarkTaupe { get { return FromHtml("#FF483C32"); } }
        public static IXLColor DarkTerraCotta { get { return FromHtml("#FFCC4E5C"); } }
        public static IXLColor DartmouthGreen { get { return FromHtml("#FF00693E"); } }
        public static IXLColor DavysGrey { get { return FromHtml("#FF555555"); } }
        public static IXLColor DebianRed { get { return FromHtml("#FFD70A53"); } }
        public static IXLColor DeepCarmine { get { return FromHtml("#FFA9203E"); } }
        public static IXLColor DeepCarminePink { get { return FromHtml("#FFEF3038"); } }
        public static IXLColor DeepCarrotOrange { get { return FromHtml("#FFE9692C"); } }
        public static IXLColor DeepCerise { get { return FromHtml("#FFDA3287"); } }
        public static IXLColor DeepChampagne { get { return FromHtml("#FFFAD6A5"); } }
        public static IXLColor DeepChestnut { get { return FromHtml("#FFB94E48"); } }
        public static IXLColor DeepFuchsia { get { return FromHtml("#FFC154C1"); } }
        public static IXLColor DeepJungleGreen { get { return FromHtml("#FF004B49"); } }
        public static IXLColor DeepLilac { get { return FromHtml("#FF9955BB"); } }
        public static IXLColor DeepMagenta { get { return FromHtml("#FFCC00CC"); } }
        public static IXLColor DeepPeach { get { return FromHtml("#FFFFCBA4"); } }
        public static IXLColor DeepSaffron { get { return FromHtml("#FFFF9933"); } }
        public static IXLColor Denim { get { return FromHtml("#FF1560BD"); } }
        public static IXLColor Desert { get { return FromHtml("#FFC19A6B"); } }
        public static IXLColor DesertSand { get { return FromHtml("#FFEDC9AF"); } }
        public static IXLColor DogwoodRose { get { return FromHtml("#FFD71868"); } }
        public static IXLColor DollarBill { get { return FromHtml("#FF85BB65"); } }
        public static IXLColor Drab { get { return FromHtml("#FF967117"); } }
        public static IXLColor DukeBlue { get { return FromHtml("#FF00009C"); } }
        public static IXLColor EarthYellow { get { return FromHtml("#FFE1A95F"); } }
        public static IXLColor Ecru { get { return FromHtml("#FFC2B280"); } }
        public static IXLColor Eggplant { get { return FromHtml("#FF614051"); } }
        public static IXLColor Eggshell { get { return FromHtml("#FFF0EAD6"); } }
        public static IXLColor EgyptianBlue { get { return FromHtml("#FF1034A6"); } }
        public static IXLColor ElectricBlue { get { return FromHtml("#FF7DF9FF"); } }
        public static IXLColor ElectricCrimson { get { return FromHtml("#FFFF003F"); } }
        public static IXLColor ElectricIndigo { get { return FromHtml("#FF6F00FF"); } }
        public static IXLColor ElectricLavender { get { return FromHtml("#FFF4BBFF"); } }
        public static IXLColor ElectricLime { get { return FromHtml("#FFCCFF00"); } }
        public static IXLColor ElectricPurple { get { return FromHtml("#FFBF00FF"); } }
        public static IXLColor ElectricUltramarine { get { return FromHtml("#FF3F00FF"); } }
        public static IXLColor ElectricViolet { get { return FromHtml("#FF8F00FF"); } }
        public static IXLColor Emerald { get { return FromHtml("#FF50C878"); } }
        public static IXLColor EtonBlue { get { return FromHtml("#FF96C8A2"); } }
        public static IXLColor Fallow { get { return FromHtml("#FFC19A6B"); } }
        public static IXLColor FaluRed { get { return FromHtml("#FF801818"); } }
        public static IXLColor Fandango { get { return FromHtml("#FFB53389"); } }
        public static IXLColor FashionFuchsia { get { return FromHtml("#FFF400A1"); } }
        public static IXLColor Fawn { get { return FromHtml("#FFE5AA70"); } }
        public static IXLColor Feldgrau { get { return FromHtml("#FF4D5D53"); } }
        public static IXLColor FernGreen { get { return FromHtml("#FF4F7942"); } }
        public static IXLColor FerrariRed { get { return FromHtml("#FFFF2800"); } }
        public static IXLColor FieldDrab { get { return FromHtml("#FF6C541E"); } }
        public static IXLColor FireEngineRed { get { return FromHtml("#FFCE2029"); } }
        public static IXLColor Flame { get { return FromHtml("#FFE25822"); } }
        public static IXLColor FlamingoPink { get { return FromHtml("#FFFC8EAC"); } }
        public static IXLColor Flavescent { get { return FromHtml("#FFF7E98E"); } }
        public static IXLColor Flax { get { return FromHtml("#FFEEDC82"); } }
        public static IXLColor FluorescentOrange { get { return FromHtml("#FFFFBF00"); } }
        public static IXLColor FluorescentYellow { get { return FromHtml("#FFCCFF00"); } }
        public static IXLColor Folly { get { return FromHtml("#FFFF004F"); } }
        public static IXLColor ForestGreenTraditional { get { return FromHtml("#FF014421"); } }
        public static IXLColor FrenchBeige { get { return FromHtml("#FFA67B5B"); } }
        public static IXLColor FrenchBlue { get { return FromHtml("#FF0072BB"); } }
        public static IXLColor FrenchLilac { get { return FromHtml("#FF86608E"); } }
        public static IXLColor FrenchRose { get { return FromHtml("#FFF64A8A"); } }
        public static IXLColor FuchsiaPink { get { return FromHtml("#FFFF77FF"); } }
        public static IXLColor Fulvous { get { return FromHtml("#FFE48400"); } }
        public static IXLColor FuzzyWuzzy { get { return FromHtml("#FFCC6666"); } }
        public static IXLColor Gamboge { get { return FromHtml("#FFE49B0F"); } }
        public static IXLColor Ginger { get { return FromHtml("#FFF9F9FF"); } }
        public static IXLColor Glaucous { get { return FromHtml("#FF6082B6"); } }
        public static IXLColor GoldenBrown { get { return FromHtml("#FF996515"); } }
        public static IXLColor GoldenPoppy { get { return FromHtml("#FFFCC200"); } }
        public static IXLColor GoldenYellow { get { return FromHtml("#FFFFDF00"); } }
        public static IXLColor GoldMetallic { get { return FromHtml("#FFD4AF37"); } }
        public static IXLColor GrannySmithApple { get { return FromHtml("#FFA8E4A0"); } }
        public static IXLColor GrayAsparagus { get { return FromHtml("#FF465945"); } }
        public static IXLColor GreenPigment { get { return FromHtml("#FF00A550"); } }
        public static IXLColor GreenRyb { get { return FromHtml("#FF66B032"); } }
        public static IXLColor Grullo { get { return FromHtml("#FFA99A86"); } }
        public static IXLColor HalayaUbe { get { return FromHtml("#FF663854"); } }
        public static IXLColor HanBlue { get { return FromHtml("#FF446CCF"); } }
        public static IXLColor HanPurple { get { return FromHtml("#FF5218FA"); } }
        public static IXLColor HansaYellow { get { return FromHtml("#FFE9D66B"); } }
        public static IXLColor Harlequin { get { return FromHtml("#FF3FFF00"); } }
        public static IXLColor HarvardCrimson { get { return FromHtml("#FFC90016"); } }
        public static IXLColor HarvestGold { get { return FromHtml("#FFDA9100"); } }
        public static IXLColor Heliotrope { get { return FromHtml("#FFDF73FF"); } }
        public static IXLColor HollywoodCerise { get { return FromHtml("#FFF400A1"); } }
        public static IXLColor HookersGreen { get { return FromHtml("#FF007000"); } }
        public static IXLColor HotMagenta { get { return FromHtml("#FFFF1DCE"); } }
        public static IXLColor HunterGreen { get { return FromHtml("#FF355E3B"); } }
        public static IXLColor Iceberg { get { return FromHtml("#FF71A6D2"); } }
        public static IXLColor Icterine { get { return FromHtml("#FFFCF75E"); } }
        public static IXLColor Inchworm { get { return FromHtml("#FFB2EC5D"); } }
        public static IXLColor IndiaGreen { get { return FromHtml("#FF138808"); } }
        public static IXLColor IndianYellow { get { return FromHtml("#FFE3A857"); } }
        public static IXLColor IndigoDye { get { return FromHtml("#FF00416A"); } }
        public static IXLColor InternationalKleinBlue { get { return FromHtml("#FF002FA7"); } }
        public static IXLColor InternationalOrange { get { return FromHtml("#FFFF4F00"); } }
        public static IXLColor Iris { get { return FromHtml("#FF5A4FCF"); } }
        public static IXLColor Isabelline { get { return FromHtml("#FFF4F0EC"); } }
        public static IXLColor IslamicGreen { get { return FromHtml("#FF009000"); } }
        public static IXLColor Jade { get { return FromHtml("#FF00A86B"); } }
        public static IXLColor Jasper { get { return FromHtml("#FFD73B3E"); } }
        public static IXLColor JazzberryJam { get { return FromHtml("#FFA50B5E"); } }
        public static IXLColor Jonquil { get { return FromHtml("#FFFADA5E"); } }
        public static IXLColor JuneBud { get { return FromHtml("#FFBDDA57"); } }
        public static IXLColor JungleGreen { get { return FromHtml("#FF29AB87"); } }
        public static IXLColor KellyGreen { get { return FromHtml("#FF4CBB17"); } }
        public static IXLColor KhakiHtmlCssKhaki { get { return FromHtml("#FFC3B091"); } }
        public static IXLColor LanguidLavender { get { return FromHtml("#FFD6CADD"); } }
        public static IXLColor LapisLazuli { get { return FromHtml("#FF26619C"); } }
        public static IXLColor LaSalleGreen { get { return FromHtml("#FF087830"); } }
        public static IXLColor LaserLemon { get { return FromHtml("#FFFEFE22"); } }
        public static IXLColor Lava { get { return FromHtml("#FFCF1020"); } }
        public static IXLColor LavenderBlue { get { return FromHtml("#FFCCCCFF"); } }
        public static IXLColor LavenderFloral { get { return FromHtml("#FFB57EDC"); } }
        public static IXLColor LavenderGray { get { return FromHtml("#FFC4C3D0"); } }
        public static IXLColor LavenderIndigo { get { return FromHtml("#FF9457EB"); } }
        public static IXLColor LavenderPink { get { return FromHtml("#FFFBAED2"); } }
        public static IXLColor LavenderPurple { get { return FromHtml("#FF967BB6"); } }
        public static IXLColor LavenderRose { get { return FromHtml("#FFFBA0E3"); } }
        public static IXLColor Lemon { get { return FromHtml("#FFFFF700"); } }
        public static IXLColor LightApricot { get { return FromHtml("#FFFDD5B1"); } }
        public static IXLColor LightBrown { get { return FromHtml("#FFB5651D"); } }
        public static IXLColor LightCarminePink { get { return FromHtml("#FFE66771"); } }
        public static IXLColor LightCornflowerBlue { get { return FromHtml("#FF93CCEA"); } }
        public static IXLColor LightFuchsiaPink { get { return FromHtml("#FFF984EF"); } }
        public static IXLColor LightMauve { get { return FromHtml("#FFDCD0FF"); } }
        public static IXLColor LightPastelPurple { get { return FromHtml("#FFB19CD9"); } }
        public static IXLColor LightSalmonPink { get { return FromHtml("#FFFF9999"); } }
        public static IXLColor LightTaupe { get { return FromHtml("#FFB38B6D"); } }
        public static IXLColor LightThulianPink { get { return FromHtml("#FFE68FAC"); } }
        public static IXLColor LightYellow1 { get { return FromHtml("#FFFFFFED"); } }
        public static IXLColor Lilac { get { return FromHtml("#FFC8A2C8"); } }
        public static IXLColor LimeColorWheel { get { return FromHtml("#FFBFFF00"); } }
        public static IXLColor LincolnGreen { get { return FromHtml("#FF195905"); } }
        public static IXLColor Liver { get { return FromHtml("#FF534B4F"); } }
        public static IXLColor Lust { get { return FromHtml("#FFE62020"); } }
        public static IXLColor MacaroniAndCheese { get { return FromHtml("#FFFFBD88"); } }
        public static IXLColor MagentaDye { get { return FromHtml("#FFCA1F7B"); } }
        public static IXLColor MagentaProcess { get { return FromHtml("#FFFF0090"); } }
        public static IXLColor MagicMint { get { return FromHtml("#FFAAF0D1"); } }
        public static IXLColor Magnolia { get { return FromHtml("#FFF8F4FF"); } }
        public static IXLColor Mahogany { get { return FromHtml("#FFC04000"); } }
        public static IXLColor Maize { get { return FromHtml("#FFFBEC5D"); } }
        public static IXLColor MajorelleBlue { get { return FromHtml("#FF6050DC"); } }
        public static IXLColor Malachite { get { return FromHtml("#FF0BDA51"); } }
        public static IXLColor Manatee { get { return FromHtml("#FF979AAA"); } }
        public static IXLColor MangoTango { get { return FromHtml("#FFFF8243"); } }
        public static IXLColor MaroonX11 { get { return FromHtml("#FFB03060"); } }
        public static IXLColor Mauve { get { return FromHtml("#FFE0B0FF"); } }
        public static IXLColor Mauvelous { get { return FromHtml("#FFEF98AA"); } }
        public static IXLColor MauveTaupe { get { return FromHtml("#FF915F6D"); } }
        public static IXLColor MayaBlue { get { return FromHtml("#FF73C2FB"); } }
        public static IXLColor MeatBrown { get { return FromHtml("#FFE5B73B"); } }
        public static IXLColor MediumAquamarine1 { get { return FromHtml("#FF66DDAA"); } }
        public static IXLColor MediumCandyAppleRed { get { return FromHtml("#FFE2062C"); } }
        public static IXLColor MediumCarmine { get { return FromHtml("#FFAF4035"); } }
        public static IXLColor MediumChampagne { get { return FromHtml("#FFF3E5AB"); } }
        public static IXLColor MediumElectricBlue { get { return FromHtml("#FF035096"); } }
        public static IXLColor MediumJungleGreen { get { return FromHtml("#FF1C352D"); } }
        public static IXLColor MediumPersianBlue { get { return FromHtml("#FF0067A5"); } }
        public static IXLColor MediumRedViolet { get { return FromHtml("#FFBB3385"); } }
        public static IXLColor MediumSpringBud { get { return FromHtml("#FFC9DC87"); } }
        public static IXLColor MediumTaupe { get { return FromHtml("#FF674C47"); } }
        public static IXLColor Melon { get { return FromHtml("#FFFDBCB4"); } }
        public static IXLColor MidnightGreenEagleGreen { get { return FromHtml("#FF004953"); } }
        public static IXLColor MikadoYellow { get { return FromHtml("#FFFFC40C"); } }
        public static IXLColor Mint { get { return FromHtml("#FF3EB489"); } }
        public static IXLColor MintGreen { get { return FromHtml("#FF98FF98"); } }
        public static IXLColor ModeBeige { get { return FromHtml("#FF967117"); } }
        public static IXLColor MoonstoneBlue { get { return FromHtml("#FF73A9C2"); } }
        public static IXLColor MordantRed19 { get { return FromHtml("#FFAE0C00"); } }
        public static IXLColor MossGreen { get { return FromHtml("#FFADDFAD"); } }
        public static IXLColor MountainMeadow { get { return FromHtml("#FF30BA8F"); } }
        public static IXLColor MountbattenPink { get { return FromHtml("#FF997A8D"); } }
        public static IXLColor MsuGreen { get { return FromHtml("#FF18453B"); } }
        public static IXLColor Mulberry { get { return FromHtml("#FFC54B8C"); } }
        public static IXLColor Mustard { get { return FromHtml("#FFFFDB58"); } }
        public static IXLColor Myrtle { get { return FromHtml("#FF21421E"); } }
        public static IXLColor NadeshikoPink { get { return FromHtml("#FFF6ADC6"); } }
        public static IXLColor NapierGreen { get { return FromHtml("#FF2A8000"); } }
        public static IXLColor NaplesYellow { get { return FromHtml("#FFFADA5E"); } }
        public static IXLColor NeonCarrot { get { return FromHtml("#FFFFA343"); } }
        public static IXLColor NeonFuchsia { get { return FromHtml("#FFFE59C2"); } }
        public static IXLColor NeonGreen { get { return FromHtml("#FF39FF14"); } }
        public static IXLColor NonPhotoBlue { get { return FromHtml("#FFA4DDED"); } }
        public static IXLColor OceanBoatBlue { get { return FromHtml("#FFCC7422"); } }
        public static IXLColor Ochre { get { return FromHtml("#FFCC7722"); } }
        public static IXLColor OldGold { get { return FromHtml("#FFCFB53B"); } }
        public static IXLColor OldLavender { get { return FromHtml("#FF796878"); } }
        public static IXLColor OldMauve { get { return FromHtml("#FF673147"); } }
        public static IXLColor OldRose { get { return FromHtml("#FFC08081"); } }
        public static IXLColor OliveDrab7 { get { return FromHtml("#FF3C341F"); } }
        public static IXLColor Olivine { get { return FromHtml("#FF9AB973"); } }
        public static IXLColor Onyx { get { return FromHtml("#FF0F0F0F"); } }
        public static IXLColor OperaMauve { get { return FromHtml("#FFB784A7"); } }
        public static IXLColor OrangeColorWheel { get { return FromHtml("#FFFF7F00"); } }
        public static IXLColor OrangePeel { get { return FromHtml("#FFFF9F00"); } }
        public static IXLColor OrangeRyb { get { return FromHtml("#FFFB9902"); } }
        public static IXLColor OtterBrown { get { return FromHtml("#FF654321"); } }
        public static IXLColor OuCrimsonRed { get { return FromHtml("#FF990000"); } }
        public static IXLColor OuterSpace { get { return FromHtml("#FF414A4C"); } }
        public static IXLColor OutrageousOrange { get { return FromHtml("#FFFF6E4A"); } }
        public static IXLColor OxfordBlue { get { return FromHtml("#FF002147"); } }
        public static IXLColor PakistanGreen { get { return FromHtml("#FF00421B"); } }
        public static IXLColor PalatinateBlue { get { return FromHtml("#FF273BE2"); } }
        public static IXLColor PalatinatePurple { get { return FromHtml("#FF682860"); } }
        public static IXLColor PaleAqua { get { return FromHtml("#FFBCD4E6"); } }
        public static IXLColor PaleBrown { get { return FromHtml("#FF987654"); } }
        public static IXLColor PaleCarmine { get { return FromHtml("#FFAF4035"); } }
        public static IXLColor PaleCerulean { get { return FromHtml("#FF9BC4E2"); } }
        public static IXLColor PaleChestnut { get { return FromHtml("#FFDDADAF"); } }
        public static IXLColor PaleCopper { get { return FromHtml("#FFDA8A67"); } }
        public static IXLColor PaleCornflowerBlue { get { return FromHtml("#FFABCDEF"); } }
        public static IXLColor PaleGold { get { return FromHtml("#FFE6BE8A"); } }
        public static IXLColor PaleMagenta { get { return FromHtml("#FFF984E5"); } }
        public static IXLColor PalePink { get { return FromHtml("#FFFADADD"); } }
        public static IXLColor PaleRobinEggBlue { get { return FromHtml("#FF96DED1"); } }
        public static IXLColor PaleSilver { get { return FromHtml("#FFC9C0BB"); } }
        public static IXLColor PaleSpringBud { get { return FromHtml("#FFECEBBD"); } }
        public static IXLColor PaleTaupe { get { return FromHtml("#FFBC987E"); } }
        public static IXLColor PansyPurple { get { return FromHtml("#FF78184A"); } }
        public static IXLColor ParisGreen { get { return FromHtml("#FF50C878"); } }
        public static IXLColor PastelBlue { get { return FromHtml("#FFAEC6CF"); } }
        public static IXLColor PastelBrown { get { return FromHtml("#FF836953"); } }
        public static IXLColor PastelGray { get { return FromHtml("#FFCFCFC4"); } }
        public static IXLColor PastelGreen { get { return FromHtml("#FF77DD77"); } }
        public static IXLColor PastelMagenta { get { return FromHtml("#FFF49AC2"); } }
        public static IXLColor PastelOrange { get { return FromHtml("#FFFFB347"); } }
        public static IXLColor PastelPink { get { return FromHtml("#FFFFD1DC"); } }
        public static IXLColor PastelPurple { get { return FromHtml("#FFB39EB5"); } }
        public static IXLColor PastelRed { get { return FromHtml("#FFFF6961"); } }
        public static IXLColor PastelViolet { get { return FromHtml("#FFCB99C9"); } }
        public static IXLColor PastelYellow { get { return FromHtml("#FFFDFD96"); } }
        public static IXLColor PaynesGrey { get { return FromHtml("#FF40404F"); } }
        public static IXLColor Peach { get { return FromHtml("#FFFFE5B4"); } }
        public static IXLColor PeachOrange { get { return FromHtml("#FFFFCC99"); } }
        public static IXLColor PeachYellow { get { return FromHtml("#FFFADFAD"); } }
        public static IXLColor Pear { get { return FromHtml("#FFD1E231"); } }
        public static IXLColor Pearl { get { return FromHtml("#FFF0EAD6"); } }
        public static IXLColor Peridot { get { return FromHtml("#FFE6E200"); } }
        public static IXLColor Periwinkle { get { return FromHtml("#FFCCCCFF"); } }
        public static IXLColor PersianBlue { get { return FromHtml("#FF1C39BB"); } }
        public static IXLColor PersianGreen { get { return FromHtml("#FF00A693"); } }
        public static IXLColor PersianIndigo { get { return FromHtml("#FF32127A"); } }
        public static IXLColor PersianOrange { get { return FromHtml("#FFD99058"); } }
        public static IXLColor PersianPink { get { return FromHtml("#FFF77FBE"); } }
        public static IXLColor PersianPlum { get { return FromHtml("#FF701C1C"); } }
        public static IXLColor PersianRed { get { return FromHtml("#FFCC3333"); } }
        public static IXLColor PersianRose { get { return FromHtml("#FFFE28A2"); } }
        public static IXLColor Persimmon { get { return FromHtml("#FFEC5800"); } }
        public static IXLColor Phlox { get { return FromHtml("#FFDF00FF"); } }
        public static IXLColor PhthaloBlue { get { return FromHtml("#FF000F89"); } }
        public static IXLColor PhthaloGreen { get { return FromHtml("#FF123524"); } }
        public static IXLColor PiggyPink { get { return FromHtml("#FFFDDDE6"); } }
        public static IXLColor PineGreen { get { return FromHtml("#FF01796F"); } }
        public static IXLColor PinkOrange { get { return FromHtml("#FFFF9966"); } }
        public static IXLColor PinkPearl { get { return FromHtml("#FFE7ACCF"); } }
        public static IXLColor PinkSherbet { get { return FromHtml("#FFF78FA7"); } }
        public static IXLColor Pistachio { get { return FromHtml("#FF93C572"); } }
        public static IXLColor Platinum { get { return FromHtml("#FFE5E4E2"); } }
        public static IXLColor PlumTraditional { get { return FromHtml("#FF8E4585"); } }
        public static IXLColor PortlandOrange { get { return FromHtml("#FFFF5A36"); } }
        public static IXLColor PrincetonOrange { get { return FromHtml("#FFFF8F00"); } }
        public static IXLColor Prune { get { return FromHtml("#FF701C1C"); } }
        public static IXLColor PrussianBlue { get { return FromHtml("#FF003153"); } }
        public static IXLColor PsychedelicPurple { get { return FromHtml("#FFDF00FF"); } }
        public static IXLColor Puce { get { return FromHtml("#FFCC8899"); } }
        public static IXLColor Pumpkin { get { return FromHtml("#FFFF7518"); } }
        public static IXLColor PurpleHeart { get { return FromHtml("#FF69359C"); } }
        public static IXLColor PurpleMountainMajesty { get { return FromHtml("#FF9678B6"); } }
        public static IXLColor PurpleMunsell { get { return FromHtml("#FF9F00C5"); } }
        public static IXLColor PurplePizzazz { get { return FromHtml("#FFFE4EDA"); } }
        public static IXLColor PurpleTaupe { get { return FromHtml("#FF50404D"); } }
        public static IXLColor PurpleX11 { get { return FromHtml("#FFA020F0"); } }
        public static IXLColor RadicalRed { get { return FromHtml("#FFFF355E"); } }
        public static IXLColor Raspberry { get { return FromHtml("#FFE30B5D"); } }
        public static IXLColor RaspberryGlace { get { return FromHtml("#FF915F6D"); } }
        public static IXLColor RaspberryPink { get { return FromHtml("#FFE25098"); } }
        public static IXLColor RaspberryRose { get { return FromHtml("#FFB3446C"); } }
        public static IXLColor RawUmber { get { return FromHtml("#FF826644"); } }
        public static IXLColor RazzleDazzleRose { get { return FromHtml("#FFFF33CC"); } }
        public static IXLColor Razzmatazz { get { return FromHtml("#FFE3256B"); } }
        public static IXLColor RedMunsell { get { return FromHtml("#FFF2003C"); } }
        public static IXLColor RedNcs { get { return FromHtml("#FFC40233"); } }
        public static IXLColor RedPigment { get { return FromHtml("#FFED1C24"); } }
        public static IXLColor RedRyb { get { return FromHtml("#FFFE2712"); } }
        public static IXLColor Redwood { get { return FromHtml("#FFAB4E52"); } }
        public static IXLColor Regalia { get { return FromHtml("#FF522D80"); } }
        public static IXLColor RichBlack { get { return FromHtml("#FF004040"); } }
        public static IXLColor RichBrilliantLavender { get { return FromHtml("#FFF1A7FE"); } }
        public static IXLColor RichCarmine { get { return FromHtml("#FFD70040"); } }
        public static IXLColor RichElectricBlue { get { return FromHtml("#FF0892D0"); } }
        public static IXLColor RichLavender { get { return FromHtml("#FFA76BCF"); } }
        public static IXLColor RichLilac { get { return FromHtml("#FFB666D2"); } }
        public static IXLColor RichMaroon { get { return FromHtml("#FFB03060"); } }
        public static IXLColor RifleGreen { get { return FromHtml("#FF414833"); } }
        public static IXLColor RobinEggBlue { get { return FromHtml("#FF00CCCC"); } }
        public static IXLColor Rose { get { return FromHtml("#FFFF007F"); } }
        public static IXLColor RoseBonbon { get { return FromHtml("#FFF9429E"); } }
        public static IXLColor RoseEbony { get { return FromHtml("#FF674846"); } }
        public static IXLColor RoseGold { get { return FromHtml("#FFB76E79"); } }
        public static IXLColor RoseMadder { get { return FromHtml("#FFE32636"); } }
        public static IXLColor RosePink { get { return FromHtml("#FFFF66CC"); } }
        public static IXLColor RoseQuartz { get { return FromHtml("#FFAA98A9"); } }
        public static IXLColor RoseTaupe { get { return FromHtml("#FF905D5D"); } }
        public static IXLColor RoseVale { get { return FromHtml("#FFAB4E52"); } }
        public static IXLColor Rosewood { get { return FromHtml("#FF65000B"); } }
        public static IXLColor RossoCorsa { get { return FromHtml("#FFD40000"); } }
        public static IXLColor RoyalAzure { get { return FromHtml("#FF0038A8"); } }
        public static IXLColor RoyalBlueTraditional { get { return FromHtml("#FF002366"); } }
        public static IXLColor RoyalFuchsia { get { return FromHtml("#FFCA2C92"); } }
        public static IXLColor RoyalPurple { get { return FromHtml("#FF7851A9"); } }
        public static IXLColor Ruby { get { return FromHtml("#FFE0115F"); } }
        public static IXLColor Ruddy { get { return FromHtml("#FFFF0028"); } }
        public static IXLColor RuddyBrown { get { return FromHtml("#FFBB6528"); } }
        public static IXLColor RuddyPink { get { return FromHtml("#FFE18E96"); } }
        public static IXLColor Rufous { get { return FromHtml("#FFA81C07"); } }
        public static IXLColor Russet { get { return FromHtml("#FF80461B"); } }
        public static IXLColor Rust { get { return FromHtml("#FFB7410E"); } }
        public static IXLColor SacramentoStateGreen { get { return FromHtml("#FF00563F"); } }
        public static IXLColor SafetyOrangeBlazeOrange { get { return FromHtml("#FFFF6700"); } }
        public static IXLColor Saffron { get { return FromHtml("#FFF4C430"); } }
        public static IXLColor Salmon1 { get { return FromHtml("#FFFF8C69"); } }
        public static IXLColor SalmonPink { get { return FromHtml("#FFFF91A4"); } }
        public static IXLColor Sand { get { return FromHtml("#FFC2B280"); } }
        public static IXLColor SandDune { get { return FromHtml("#FF967117"); } }
        public static IXLColor Sandstorm { get { return FromHtml("#FFECD540"); } }
        public static IXLColor SandyTaupe { get { return FromHtml("#FF967117"); } }
        public static IXLColor Sangria { get { return FromHtml("#FF92000A"); } }
        public static IXLColor SapGreen { get { return FromHtml("#FF507D2A"); } }
        public static IXLColor Sapphire { get { return FromHtml("#FF082567"); } }
        public static IXLColor SatinSheenGold { get { return FromHtml("#FFCBA135"); } }
        public static IXLColor Scarlet { get { return FromHtml("#FFFF2000"); } }
        public static IXLColor SchoolBusYellow { get { return FromHtml("#FFFFD800"); } }
        public static IXLColor ScreaminGreen { get { return FromHtml("#FF76FF7A"); } }
        public static IXLColor SealBrown { get { return FromHtml("#FF321414"); } }
        public static IXLColor SelectiveYellow { get { return FromHtml("#FFFFBA00"); } }
        public static IXLColor Sepia { get { return FromHtml("#FF704214"); } }
        public static IXLColor Shadow { get { return FromHtml("#FF8A795D"); } }
        public static IXLColor ShamrockGreen { get { return FromHtml("#FF009E60"); } }
        public static IXLColor ShockingPink { get { return FromHtml("#FFFC0FC0"); } }
        public static IXLColor Sienna1 { get { return FromHtml("#FF882D17"); } }
        public static IXLColor Sinopia { get { return FromHtml("#FFCB410B"); } }
        public static IXLColor Skobeloff { get { return FromHtml("#FF007474"); } }
        public static IXLColor SkyMagenta { get { return FromHtml("#FFCF71AF"); } }
        public static IXLColor SmaltDarkPowderBlue { get { return FromHtml("#FF003399"); } }
        public static IXLColor SmokeyTopaz { get { return FromHtml("#FF933D41"); } }
        public static IXLColor SmokyBlack { get { return FromHtml("#FF100C08"); } }
        public static IXLColor SpiroDiscoBall { get { return FromHtml("#FF0FC0FC"); } }
        public static IXLColor SplashedWhite { get { return FromHtml("#FFFEFDFF"); } }
        public static IXLColor SpringBud { get { return FromHtml("#FFA7FC00"); } }
        public static IXLColor StPatricksBlue { get { return FromHtml("#FF23297A"); } }
        public static IXLColor StilDeGrainYellow { get { return FromHtml("#FFFADA5E"); } }
        public static IXLColor Straw { get { return FromHtml("#FFE4D96F"); } }
        public static IXLColor Sunglow { get { return FromHtml("#FFFFCC33"); } }
        public static IXLColor Sunset { get { return FromHtml("#FFFAD6A5"); } }
        public static IXLColor Tangelo { get { return FromHtml("#FFF94D00"); } }
        public static IXLColor Tangerine { get { return FromHtml("#FFF28500"); } }
        public static IXLColor TangerineYellow { get { return FromHtml("#FFFFCC00"); } }
        public static IXLColor Taupe { get { return FromHtml("#FF483C32"); } }
        public static IXLColor TaupeGray { get { return FromHtml("#FF8B8589"); } }
        public static IXLColor TeaGreen { get { return FromHtml("#FFD0F0C0"); } }
        public static IXLColor TealBlue { get { return FromHtml("#FF367588"); } }
        public static IXLColor TealGreen { get { return FromHtml("#FF006D5B"); } }
        public static IXLColor TeaRoseOrange { get { return FromHtml("#FFF88379"); } }
        public static IXLColor TeaRoseRose { get { return FromHtml("#FFF4C2C2"); } }
        public static IXLColor TennéTawny { get { return FromHtml("#FFCD5700"); } }
        public static IXLColor TerraCotta { get { return FromHtml("#FFE2725B"); } }
        public static IXLColor ThulianPink { get { return FromHtml("#FFDE6FA1"); } }
        public static IXLColor TickleMePink { get { return FromHtml("#FFFC89AC"); } }
        public static IXLColor TiffanyBlue { get { return FromHtml("#FF0ABAB5"); } }
        public static IXLColor TigersEye { get { return FromHtml("#FFE08D3C"); } }
        public static IXLColor Timberwolf { get { return FromHtml("#FFDBD7D2"); } }
        public static IXLColor TitaniumYellow { get { return FromHtml("#FFEEE600"); } }
        public static IXLColor Toolbox { get { return FromHtml("#FF746CC0"); } }
        public static IXLColor TractorRed { get { return FromHtml("#FFFD0E35"); } }
        public static IXLColor TropicalRainForest { get { return FromHtml("#FF00755E"); } }
        public static IXLColor TuftsBlue { get { return FromHtml("#FF417DC1"); } }
        public static IXLColor Tumbleweed { get { return FromHtml("#FFDEAA88"); } }
        public static IXLColor TurkishRose { get { return FromHtml("#FFB57281"); } }
        public static IXLColor Turquoise1 { get { return FromHtml("#FF30D5C8"); } }
        public static IXLColor TurquoiseBlue { get { return FromHtml("#FF00FFEF"); } }
        public static IXLColor TurquoiseGreen { get { return FromHtml("#FFA0D6B4"); } }
        public static IXLColor TuscanRed { get { return FromHtml("#FF823535"); } }
        public static IXLColor TwilightLavender { get { return FromHtml("#FF8A496B"); } }
        public static IXLColor TyrianPurple { get { return FromHtml("#FF66023C"); } }
        public static IXLColor UaBlue { get { return FromHtml("#FF0033AA"); } }
        public static IXLColor UaRed { get { return FromHtml("#FFD9004C"); } }
        public static IXLColor Ube { get { return FromHtml("#FF8878C3"); } }
        public static IXLColor UclaBlue { get { return FromHtml("#FF536895"); } }
        public static IXLColor UclaGold { get { return FromHtml("#FFFFB300"); } }
        public static IXLColor UfoGreen { get { return FromHtml("#FF3CD070"); } }
        public static IXLColor Ultramarine { get { return FromHtml("#FF120A8F"); } }
        public static IXLColor UltramarineBlue { get { return FromHtml("#FF4166F5"); } }
        public static IXLColor UltraPink { get { return FromHtml("#FFFF6FFF"); } }
        public static IXLColor Umber { get { return FromHtml("#FF635147"); } }
        public static IXLColor UnitedNationsBlue { get { return FromHtml("#FF5B92E5"); } }
        public static IXLColor UnmellowYellow { get { return FromHtml("#FFFFFF66"); } }
        public static IXLColor UpForestGreen { get { return FromHtml("#FF014421"); } }
        public static IXLColor UpMaroon { get { return FromHtml("#FF7B1113"); } }
        public static IXLColor UpsdellRed { get { return FromHtml("#FFAE2029"); } }
        public static IXLColor Urobilin { get { return FromHtml("#FFE1AD21"); } }
        public static IXLColor UscCardinal { get { return FromHtml("#FF990000"); } }
        public static IXLColor UscGold { get { return FromHtml("#FFFFCC00"); } }
        public static IXLColor UtahCrimson { get { return FromHtml("#FFD3003F"); } }
        public static IXLColor Vanilla { get { return FromHtml("#FFF3E5AB"); } }
        public static IXLColor VegasGold { get { return FromHtml("#FFC5B358"); } }
        public static IXLColor VenetianRed { get { return FromHtml("#FFC80815"); } }
        public static IXLColor Verdigris { get { return FromHtml("#FF43B3AE"); } }
        public static IXLColor Vermilion { get { return FromHtml("#FFE34234"); } }
        public static IXLColor Veronica { get { return FromHtml("#FFA020F0"); } }
        public static IXLColor Violet1 { get { return FromHtml("#FF8F00FF"); } }
        public static IXLColor VioletColorWheel { get { return FromHtml("#FF7F00FF"); } }
        public static IXLColor VioletRyb { get { return FromHtml("#FF8601AF"); } }
        public static IXLColor Viridian { get { return FromHtml("#FF40826D"); } }
        public static IXLColor VividAuburn { get { return FromHtml("#FF922724"); } }
        public static IXLColor VividBurgundy { get { return FromHtml("#FF9F1D35"); } }
        public static IXLColor VividCerise { get { return FromHtml("#FFDA1D81"); } }
        public static IXLColor VividTangerine { get { return FromHtml("#FFFFA089"); } }
        public static IXLColor VividViolet { get { return FromHtml("#FF9F00FF"); } }
        public static IXLColor WarmBlack { get { return FromHtml("#FF004242"); } }
        public static IXLColor Wenge { get { return FromHtml("#FF645452"); } }
        public static IXLColor WildBlueYonder { get { return FromHtml("#FFA2ADD0"); } }
        public static IXLColor WildStrawberry { get { return FromHtml("#FFFF43A4"); } }
        public static IXLColor WildWatermelon { get { return FromHtml("#FFFC6C85"); } }
        public static IXLColor Wisteria { get { return FromHtml("#FFC9A0DC"); } }
        public static IXLColor Xanadu { get { return FromHtml("#FF738678"); } }
        public static IXLColor YaleBlue { get { return FromHtml("#FF0F4D92"); } }
        public static IXLColor YellowMunsell { get { return FromHtml("#FFEFCC00"); } }
        public static IXLColor YellowNcs { get { return FromHtml("#FFFFD300"); } }
        public static IXLColor YellowProcess { get { return FromHtml("#FFFFEF00"); } }
        public static IXLColor YellowRyb { get { return FromHtml("#FFFEFE33"); } }
        public static IXLColor Zaffre { get { return FromHtml("#FF0014A8"); } }
        public static IXLColor ZinnwalditeBrown { get { return FromHtml("#FF2C1608"); } }

    }
}
