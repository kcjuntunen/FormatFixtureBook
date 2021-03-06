﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace FormatFixtureBook
{
	public class ExcelReader
	{
		private FileInfo xlsFileInfo;
		private LinkedList<PageInfo> subSections;
		private string initialDir = @"G:\ZALES\FIXTURE BOOK\SECTIONS";
		private string extension = @"SLDDRW";
		private ExcelReaderExtensionOptions extOpt;
		private ExcelReaderSearchOptions searchOpt;
		private bool foundFileFlag = false;

		public event EventHandler NewDir;
		protected virtual void OnNewDir(FileSystemEventArgs e) {
			NewDir?.Invoke(this, e);
		}

		public enum ExcelReaderExtensionOptions {
			SLDDRW = 0x02,
			PDF = 0x04
		}

		public enum ExcelReaderSearchOptions {
			PARENT_DIR = 0x02,
			TEMP_DIR = 0x04,
			THIS_DIR = 0x08,
			RECURSE = 0x10
		}

		public ExcelReader() {
			SelectFile();
			extOpt = ExcelReaderExtensionOptions.SLDDRW;
			searchOpt = ExcelReaderSearchOptions.PARENT_DIR;
			set_options();
		}

		public ExcelReader(string _fileName) {
			xlsFileInfo = new FileInfo(_fileName);
			initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
			validate(xlsFileInfo);
			set_options();
		}

		public ExcelReader(FileInfo fileInfo) {
			xlsFileInfo = fileInfo;
			initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
			validate(fileInfo);
			set_options();
		}
		
		public ExcelReader(string _fileName, ExcelReaderExtensionOptions _eo, ExcelReaderSearchOptions _so) {
			xlsFileInfo = new FileInfo(_fileName);
			extOpt = _eo;
			searchOpt = _so;
			validate(xlsFileInfo);
			set_options();
		}

		public ExcelReader(FileInfo fileInfo, ExcelReaderExtensionOptions _eo, ExcelReaderSearchOptions _so) {
			xlsFileInfo = fileInfo;
			extOpt = _eo;
			searchOpt = _so;
			validate(xlsFileInfo);
			set_options();
		}

		private void validate(FileInfo fileInfo) {
			if (fileInfo.Name.ToUpper().EndsWith(@"XLS") || fileInfo.Name.ToUpper().EndsWith(@"XLSX")) {
				set_options();
			} else {
				throw new ExcelReaderException(@"Unable to read non-Excel documents.");
			}
		}

		private void set_options() {
			if (extOpt == ExcelReaderExtensionOptions.SLDDRW) {
				extension = @"SLDDRW";
			}

			if (extOpt == ExcelReaderExtensionOptions.PDF) {
				extension = @"PDF";
			}

			if ((searchOpt & ExcelReaderSearchOptions.PARENT_DIR) == ExcelReaderSearchOptions.PARENT_DIR) {
				initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
			}


			if ((searchOpt & ExcelReaderSearchOptions.THIS_DIR) == ExcelReaderSearchOptions.THIS_DIR) {
				initialDir = string.Format(@"{0}\", xlsFileInfo.DirectoryName);
			}

			if ((searchOpt & ExcelReaderSearchOptions.TEMP_DIR) == ExcelReaderSearchOptions.TEMP_DIR) {
				initialDir = string.Format(@"{0}\", Path.GetTempPath());
			}
		}

		private void SelectFile() {
			OpenFileDialog ofd_ = new OpenFileDialog();
			ofd_.Filter = @"Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls";
			ofd_.FilterIndex = 0;
			ofd_.InitialDirectory = initialDir;

			if (ofd_.ShowDialog() == DialogResult.OK) {
				xlsFileInfo = new FileInfo(ofd_.FileName);
				initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
				validate(xlsFileInfo);
			}
		}

		private PageInfo find_file(string cell1_, string cell2_, string cell3_) {
			PageInfo currentPageInfo = new PageInfo(cell1_, cell2_, cell3_);
			string sDir_ = xlsFileInfo.DirectoryName;
			bool VendorInfo_ = false;
			if ((searchOpt & ExcelReaderSearchOptions.RECURSE) == ExcelReaderSearchOptions.RECURSE) {
				foreach (string dir_ in Directory.GetDirectories(initialDir)) {
					OnNewDir(new FileSystemEventArgs(WatcherChangeTypes.All, dir_, xlsFileInfo.FullName));
					currentPageInfo.fileInfo = search(cell1_, cell3_, dir_, ref VendorInfo_);
					if (currentPageInfo.fileInfo != null) {
						break;
					}
				}
			} else if ((searchOpt & ExcelReaderSearchOptions.THIS_DIR) == ExcelReaderSearchOptions.THIS_DIR) {
				OnNewDir(new FileSystemEventArgs(WatcherChangeTypes.All, sDir_, xlsFileInfo.Name));
				currentPageInfo.fileInfo = search(cell1_, cell3_, sDir_, ref VendorInfo_);
			} else if ((searchOpt & ExcelReaderSearchOptions.TEMP_DIR) == ExcelReaderSearchOptions.TEMP_DIR) {
				OnNewDir(new FileSystemEventArgs(WatcherChangeTypes.All, sDir_, xlsFileInfo.Name));
				sDir_ = string.Format(@"{0}\", Path.GetTempPath());
				currentPageInfo.fileInfo = search(cell1_, cell3_, sDir_, ref VendorInfo_);
			}
			currentPageInfo.VendorInfo = VendorInfo_;
			return currentPageInfo;
		}

		private FileInfo search(string termA, string termB, string dir, ref bool VendorInfo) {
			string[] f_ = Directory.GetFiles(dir, string.Format(@"FB.{0}.{1}", termA, extension));
			if (f_.Length > 0) {
				foundFileFlag = true;
				return new FileInfo(f_[0]);
			} else {
				string[] f3_ = Directory.GetFiles(dir, string.Format(@"{0}.{1}", termB, extension));
				if (f3_.Length > 0) {
					foundFileFlag = true;
					VendorInfo = true;
					return new FileInfo(f3_[0]);
				}
			}
			return null;
		}

		private void post_search() {
			var nd_ = subSections.First;
			while (nd_ != null) {
				if (nd_.Value.fileInfo == null) {
					string first = nd_.Value.FirstSheetNo().Value;
					string last = nd_.Value.LastSheetNo().Value;
					string f_ = string.Format(@"{0} - {1}", first, last.Replace(@"Z ", string.Empty));
					PageInfo pi_ = find_file(string.Empty, string.Empty, f_);
					if (pi_.fileInfo != null && pi_.fileInfo.Exists) {
						nd_.Value.fileInfo = pi_.fileInfo;
						foundFileFlag = true;
					}
				}
				nd_ = nd_.Next;
			}
		}

		public LinkedList<PageInfo> ReadFile() {
			try {
				using (ExcelPackage xlp_ = new ExcelPackage(xlsFileInfo)) {
					ExcelWorksheet wksht_ = xlp_.Workbook.Worksheets[1];
					int rows_ = wksht_.Dimension.End.Row;
					PageInfo currentPageInfo = new PageInfo(@"NULL", @"NULL", @"NULL");

					if (subSections == null) {
						subSections = new LinkedList<PageInfo>();
					}

					string tmp_ = string.Empty;
					string range_ = string.Empty;

					for (int i = 2; i <= rows_; i++) {
						string cell1_ = Convert.ToString(wksht_.Cells[i, 1].Value).Trim();
						string cell2_ = Convert.ToString(wksht_.Cells[i, 2].Value).Trim();
						string cell3_ = Convert.ToString(wksht_.Cells[i, 3].Value).Trim();

						if (cell1_ != string.Empty) {
							tmp_ = cell1_;
							range_ = wksht_.Cells[i, 1, i, 3].Address;
							currentPageInfo = find_file(cell1_, cell2_, cell3_);
							currentPageInfo.Name = tmp_;
							currentPageInfo.CellAddress = range_;
							subSections.AddLast(currentPageInfo);
						} else if (cell2_ != string.Empty && cell3_ != string.Empty) {
							currentPageInfo.Add(cell2_, cell3_);
							currentPageInfo.Name = tmp_;
							currentPageInfo.CellAddress = range_;
							//subSections.AddLast(currentPageInfo);
						}
					}
				}
			} catch (IOException _ioe) {
				throw new ExcelReaderException(@"Could not read file.\nIs it open in another window?", _ioe);
			}
			if (extension == @"PDF") {
				post_search();
			}
			if (!foundFileFlag) {
					throw new ExcelReaderFoundNoFilesException(
						"Couldn't find any files.\n" +
						"Try checking the \"Recursive\" box or move\n" +
						"the Excel file to the same directory as the drawings.");
				}
			return subSections;
		}
	}
}
