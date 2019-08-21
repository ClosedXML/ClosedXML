using System.Collections;

namespace ClosedXML.Word
{
    public interface IXLBlocks : IEnumerable
    {
        IXLDocument Document { get; set; }

        bool TryGetBlock(int blockId, out IXLBlock block);

        bool TryGetBlock(string blockName, out IXLBlock block);

        int GenerateBlockIds(bool fromLoadedDocument = false);
    }
}
