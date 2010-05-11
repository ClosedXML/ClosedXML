using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLSharedStrings: IEnumerable<String>
    {
        internal class SharedStringInfo
        {
            public UInt32 Position { get; set; }
            public UInt32 Count { get; set; }
        }

        private Dictionary<String, SharedStringInfo> sharedStrings = new Dictionary<String, SharedStringInfo>();

        private UInt32 lastPosition = 0;
        public UInt32 Add(String sharedString)
        {
            SharedStringInfo stringInfo;
            if(sharedStrings.ContainsKey(sharedString))
            {
                stringInfo = sharedStrings[sharedString];
                stringInfo.Count++;
            }
            else
            {
                stringInfo = new SharedStringInfo() { Position = lastPosition, Count = 1 };
                sharedStrings.Add(sharedString, stringInfo);
                lastPosition++;
            }
            return stringInfo.Position;
        }

        public String GetString(UInt32 position)
        {
            return sharedStrings.Where(s => s.Value.Position == position).Single().Key;
        }

        public UInt32 Count
        {
            get { return (UInt32)sharedStrings.Count; }
        }

        #region IEnumerable<string> Members

        public IEnumerator<String> GetEnumerator()
        {
            return sharedStrings.Keys.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion
    }
}
