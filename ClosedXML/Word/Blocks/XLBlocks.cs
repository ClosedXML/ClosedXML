using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Word
{
    public class XLBlocks : IXLBlocks
    {
        private readonly List<IXLBlock> _blocks = new List<IXLBlock>( );
        private readonly Dictionary<string, IXLBlock> _blockIds = new Dictionary<string, IXLBlock>( );
        public IXLDocument Document { get; set; }

        public XLBlocks( IXLDocument document )
        {
            Document = document;
        }

        IEnumerator IEnumerable.GetEnumerator( )
        {
            return new BlocksEnum( _blocks );
        }

        public BlocksEnum GetEnumerator( )
        {
            return new BlocksEnum( _blocks );
        }

        public bool TryGetBlock( string blockId, out IXLBlock block )
        {
            if ( _blockIds.Count != 0 )
            {
                if ( _blockIds.TryGetValue( blockId, out IXLBlock b ) )
                {
                    block = b;
                    return true;
                }
            }

            block = null;
            return false;
        }

        public int Count
        {
            get { return _blocks.Count; }
        }
    }

    public class BlocksEnum : IEnumerator
    {
        private int position = -1;
        private readonly List<IXLBlock> Blocks;

        public BlocksEnum( List<IXLBlock> blocks )
        {
            Blocks = blocks;
        }

        public bool MoveNext( )
        {
            position++;
            return position < Blocks.Count;
        }

        public void Reset( )
        {
            position = -1;
        }

        public object Current
        {
            get
            {
                try
                {
                    return Blocks[position];
                }
                catch ( IndexOutOfRangeException )
                {
                    throw new IndexOutOfRangeException( );
                }
            }
        }
    }

    public static class XLBlocksExtensions
    {
        internal static IXLBlock Add( this IXLBlocks blocks, IXLBlock block )
        {
            throw new NotImplementedException( );
        }
    }
}
