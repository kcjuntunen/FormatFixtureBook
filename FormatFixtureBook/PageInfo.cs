using System.Collections.Generic;
using System.IO;

namespace FormatFixtureBook {
	public class PageInfo {
		private string itemGroup;
		private LinkedList<string> pageDescription;
		private LinkedList<string> sheetNo;
		public FileInfo fileInfo;

		public PageInfo(string itm_grp, string pg_descr, string shtNo) {
			itemGroup = itm_grp;
			pageDescription = new LinkedList<string>();
			pageDescription.AddLast(pg_descr);

			sheetNo = new LinkedList<string>();
			sheetNo.AddLast(shtNo);
		}

		public void Add(string pg_descr, string shtNo) {
			pageDescription.AddLast(pg_descr);
			sheetNo.AddLast(shtNo);
		}

		public LinkedListNode<string> FirstDescription() {
			return pageDescription.First;
		}

		public LinkedListNode<string> FirstSheetNo() {
			return sheetNo.First;
		}
	}
}
