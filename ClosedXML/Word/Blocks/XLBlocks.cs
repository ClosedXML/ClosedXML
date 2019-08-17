using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Word
{
    public class XLBlocks : IXLBlocks
    {
        private readonly List<IXLBlock> _blocks = new List<IXLBlock>();
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
    }

    public class BlocksEnum : IEnumerator
    {
        private int position = -1;
        public readonly List<IXLBlock> Blocks;

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
            foreach (IXLBlock block in blocks)
            {
                Console.WriteLine(block.ToString());
                return block;
            }

            return null;

            //TODO Add blocks to document
        }
    }
}
