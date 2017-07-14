using ClosedXML_Examples.Delete;
using System.IO;

namespace ClosedXML_Examples
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