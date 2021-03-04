using ClosedXML.Examples.Delete;
using System.IO;

namespace ClosedXML.Examples
{
    public class ModifyFiles
    {
        public static void Run()
        {
            var path = Program.BaseModifiedDirectory;
            new DeleteRows().Create(Path.Combine(path, "DeleteRows.xlsx"));
        }
    }
}