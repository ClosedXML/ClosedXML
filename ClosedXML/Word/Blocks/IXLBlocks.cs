using System.Collections;

namespace ClosedXML.Word
{
    public interface IXLBlocks : IEnumerable
    {
        IXLDocument Document { get; set; }

        bool TryGetBlock(string blockId, out IXLBlock block);

        int Count { get; }
    }
}
