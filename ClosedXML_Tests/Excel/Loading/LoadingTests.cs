using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    // Tests in this fixture test only the successful loading of existing Excel files, 
    // i.e. we test that ClosedXML doesn't choke on a given input file
    // These tests DO NOT test that ClosedXML successfully recognises all the Excel parts or that it can successfully save those parts again.
    [TestFixture]
    public class LoadingTests
    {
        [Test]
        public void CanSuccessfullyLoadFiles()
        {
            var files = new List<string>()
            {
                @"Misc\TableWithCustomTheme.xlsx"
            };

            foreach (var file in files)
            {
                TestHelper.LoadFile(file);
            }
        }
    }
}