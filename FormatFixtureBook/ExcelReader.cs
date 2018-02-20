using System;
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

		public enum ExcelReaderExtensionOptions {
			SLDDRW = 0x02,
			PDF = 0x04
		}

		public enum ExcelReaderSearchOptions {
			PARENT_DIR = 0x02,
			TEMP_DIR = 0x04
		}

		public ExcelReader() {
			SelectFile();
			extOpt = ExcelReaderExtensionOptions.SLDDRW;
			searchOpt = ExcelReaderSearchOptions.PARENT_DIR;
		}

		public ExcelReader(string _fileName) {
			xlsFileInfo = new FileInfo(_fileName);
			initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
			if (_fileName.ToUpper().EndsWith(@"XLS") || _fileName.ToUpper().EndsWith(@"XLSX")) {
				set_options();
			} else {
				throw new ExcelReaderException(@"Unable to read non-Excel documents.");
			}
		}
		
		public ExcelReader(string _fileName, ExcelReaderExtensionOptions _eo, ExcelReaderSearchOptions _so) {
			xlsFileInfo = new FileInfo(_fileName);
			initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
			extOpt = _eo;
			searchOpt = _so;
			set_options();
		}

		private void set_options() {
			if (extOpt == ExcelReaderExtensionOptions.SLDDRW) {
				extension = @"SLDDRW";
			}

			if (extOpt == ExcelReaderExtensionOptions.PDF) {
				extension = @"PDF";
			}

			if ((searchOpt & ExcelReaderSearchOptions.PARENT_DIR) == ExcelReaderSearchOptions.PARENT_DIR) {

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
			}
		}

		private PageInfo find_file(string cell1_, string cell2_, string cell3_) {
			PageInfo currentPageInfo = new PageInfo(cell1_, cell2_, cell3_);
			string sDir_ = xlsFileInfo.DirectoryName;
			if ((searchOpt & ExcelReaderSearchOptions.PARENT_DIR) == ExcelReaderSearchOptions.PARENT_DIR) {
				sDir_ = string.Format(@"{0}\..", sDir_);
				foreach (string dir_ in Directory.GetDirectories(initialDir)) {
					currentPageInfo.fileInfo = search(cell1_, cell3_, dir_);
				}
			} else if ((searchOpt & ExcelReaderSearchOptions.TEMP_DIR) == ExcelReaderSearchOptions.TEMP_DIR) {
				currentPageInfo.fileInfo = search(cell1_, cell3_, sDir_);
			}
			return currentPageInfo;
		}

		private FileInfo search(string termA, string termB, string dir) {
			string[] f_ = Directory.GetFiles(dir, string.Format(@"FB.{0}.{1}", termA, extension));
			if (f_.Length > 0) {
				return new FileInfo(f_[0]);
			} else {
				string[] f3_ = Directory.GetFiles(dir, string.Format(@"{0}.{1}", termB, extension));
				if (f3_.Length > 0) {
					return new FileInfo(f3_[0]);
				}
			}
			return null;
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

					for (int i = 2; i <= rows_; i++) {
						string cell1_ = Convert.ToString(wksht_.Cells[i, 1].Value).Trim();
						string cell2_ = Convert.ToString(wksht_.Cells[i, 2].Value).Trim();
						string cell3_ = Convert.ToString(wksht_.Cells[i, 3].Value).Trim();

						if (cell1_ != string.Empty) {
							currentPageInfo = find_file(cell1_, cell2_, cell3_);
							subSections.AddLast(currentPageInfo);
						} else if (cell2_ != string.Empty && cell3_ != string.Empty) {
							currentPageInfo.Add(cell2_, cell3_);
						}


					}
				}
			} catch (IOException _ioe) {
				throw new ExcelReaderException(@"Could not read file.\nIs it open in another window?", _ioe);
			}
			return subSections;
		}
	}
}
