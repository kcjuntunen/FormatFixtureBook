﻿using System.Collections.Generic;
using System.IO;

namespace FormatFixtureBook {
	public class PageInfo {
		private string itemGroup;
		private LinkedList<string> pageDescription;
		private LinkedList<string> sheetNo;
		public string Name;
		public bool VendorInfo;
		public FileInfo fileInfo;
		public string CellAddress = string.Empty;

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

		public LinkedListNode<string> LastDescription() {
			return pageDescription.Last;
		}

		public LinkedListNode<string> FirstSheetNo() {
			return sheetNo.First;
		}
		
		public LinkedListNode<string> LastSheetNo() {
			return sheetNo.Last;
		}
	}
}
