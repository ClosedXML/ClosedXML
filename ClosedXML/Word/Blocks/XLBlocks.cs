using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Word
{
    public class XLBlocks : IXLBlocks
    {
        IEnumerator IEnumerable.GetEnumerator( )
        {
            return GetEnumerator( );
        }

        public IEnumerator<IXLBlock> GetEnumerator( )
        {
            throw new System.NotImplementedException( );
        }
    }

    public static class XLBlocksExtensions
    {
        public static IEnumerable<IXLBlock> Add( this IEnumerable<IXLBlock> source, IXLBlock block )
        {
            foreach ( IXLBlock b in source )
            {
                yield return b;
            }
            yield return block;
        }

        public static IEnumerable<IXLBlock> AddBlocksToDocument( this IEnumerable<IXLBlock> source )
        {
            throw new System.NotImplementedException( );
            //TODO Add blocks to document
        }
    }
}
