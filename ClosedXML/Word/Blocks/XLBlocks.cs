using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Word
{
    public class XLBlocks : IXLBlocks
    {
        private readonly List<IXLBlock> _blocks = new List<IXLBlock>();
        private readonly Dictionary<string, IXLBlock> _blockNames = new Dictionary<string, IXLBlock>();
        private readonly Dictionary<int, IXLBlock> _blockIds = new Dictionary<int, IXLBlock>();
        public IXLDocument Document { get; set; }

        public XLBlocks(IXLDocument document)
        {
            Document = document;
        }

        IEnumerator IEnumerable.GetEnumerator( )
        {
            return new BlocksEnum(_blocks);
        }

        public BlocksEnum GetEnumerator( )
        {
            return new BlocksEnum(_blocks);
        }

        public bool TryGetBlock(int blockId, out IXLBlock block)
        {
            if (_blockIds.Count != 0)
            {
                if (_blockIds.TryGetValue(blockId, out IXLBlock b))
                {
                    block = b;
                    return true;
                }
            }

            block = null;
            return false;
        }

        public bool TryGetBlock(string blockName, out IXLBlock block)
        {
            if (_blockNames.Count != 0)
            {
                if (_blockNames.TryGetValue(blockName, out IXLBlock b))
                {
                    block = b;
                    return true;
                }
            }

            block = null;
            return false;
        }

        public int GenerateBlockIds(bool fromLoadedDocument = false)
        {
            if (fromLoadedDocument)
            {
                throw new NotImplementedException();
            }

            if (_blocks.Count != 0)
            {
                return _blocks.Count + 1;
            }

            return 0;
        }
    }

    public class BlocksEnum : IEnumerator
    {
        private int position = -1;
        private readonly List<IXLBlock> Blocks;

        public BlocksEnum(List<IXLBlock> blocks)
        {
            this.Blocks = blocks;
        }

        public bool MoveNext()
        {
            position++;
            return (position < Blocks.Count);
        }

        public void Reset()
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
                catch (IndexOutOfRangeException)
                {
                    throw new IndexOutOfRangeException();
                }
            }
        }
    }

    public static class XLBlocksExtensions
    {
        public static IXLBlock Add( this IXLBlocks blocks, IXLBlock block)
        {
            blocks.Document.AddBlock(block);
            return block;
        }

        public static IXLBlock AddBlocksToDocument( this IXLBlocks blocks )
        {
            foreach (IXLBlock block in blocks.Document.Blocks())
            {
                block.BlockId = blocks.GenerateBlockIds();
                return block;
            }

            return null;

            //TODO Add blocks to document
        }
    }
}
