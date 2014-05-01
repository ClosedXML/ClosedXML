using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace ClosedXML.Excel.Misc
{
	/// <summary>
	/// This is most definitely NOT thread-safe, but it is reentrant (e.g. you can insert and remove
	/// items during enumeration)
	/// </summary>
	/// <typeparam name="T"></typeparam>
	internal class XLReentrantEnumerableSet<T> : IEnumerable<T>
	{
		private HashSet<T> _hashSet;
		private List<T> _list;

		private int _activeEnumerators;

		public XLReentrantEnumerableSet()
		{
			_list = new List<T>();
			_hashSet = new HashSet<T>();
		}

		public void Add(T item)
		{
			if(!_hashSet.Contains(item))
			{
				// Add item to end of list
				_list.Add(item);

				// Store the item in the hashset too.
				_hashSet.Add(item);
			}
		}

		public void Remove(T item)
		{
			_hashSet.Remove(item);
			
			fixup();
		}

		private void fixup()
		{
			// Only fixup the list if there are no active enumerators
			if(_activeEnumerators > 0)
				return;

			// Only fixup the list if there are more than a certain number of items to deal with
			// This saves fixing up the list continually.
			if(_list.Count < 1000 || (_list.Count < _hashSet.Count * 1.5))
				return;

			// Rebuild the list skipping out omitted items
			_list = _list.Where(item => _hashSet.Contains(item)).ToList();
		}
		
		public IEnumerator<T> GetEnumerator()
		{
			// Mark that we are enumerating
			_activeEnumerators++;
			try
			{
				int idx = 0;

				// Important, store count here, as more items may be added while we are enumerating
				// and we only want to enumerate items that were already there.
				int count = _list.Count;
				while(idx < count)
				{
					var item = _list[idx];

					// Skip over items in the list which aren't in the hashset; they could have been
					// removed while we were enumerating or previously removed.
					if(_hashSet.Contains(item))
					{
						yield return item;
					}

					idx++;
				}

			}
			finally
			{
				// Finished enumerating, can now fixup
				_activeEnumerators--;
				fixup();
			}
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.GetEnumerator();
		}
	}
}
