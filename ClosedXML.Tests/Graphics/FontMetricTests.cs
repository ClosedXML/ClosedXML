using ClosedXML.Graphics;
using NUnit.Framework;
using System;
using System.IO;

namespace ClosedXML.Tests.Graphics
{
    [TestFixture]
    public class FontMetricTests
    {
        private readonly FontMetric _font;

        public FontMetricTests()
        {
            var fontStream = File.OpenRead(Environment.ExpandEnvironmentVariables("%SystemRoot%/Fonts/calibri.ttf"));
            _font = FontMetric.LoadTrueType(fontStream);
        }

        [Test]
        public void ReadsFontMetrics()
        {
            Assert.That(_font.Ascent, Is.EqualTo(1950));
            Assert.That(_font.Descent, Is.EqualTo(550));
            Assert.That(_font.GetAdvanceWidth('0'), Is.EqualTo(1038d));
        }

        [Test]
        public void UndefinedCodepointsHaveEmSize()
        {
            var width = _font.GetAdvanceWidth((char)0x2);
            Assert.That(width, Is.EqualTo(2048));
        }
    }
}
