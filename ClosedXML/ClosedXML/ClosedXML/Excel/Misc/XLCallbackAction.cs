using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Misc
{
	internal class XLCallbackAction
	{
		public XLCallbackAction(Action<XLRange, Int32> action)
		{
			this.Action = action;
		}

		public Action<XLRange, Int32> Action { get; set; }
	}
}
