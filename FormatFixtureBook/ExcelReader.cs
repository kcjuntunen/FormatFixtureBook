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
		private string initialDir = @"G:\ZALES\FIXTURE BOOK\SECTIONS";
		private LinkedList<PageInfo> subSections;

		public string Extension = @"SLDDRW";

		public ExcelReader() {
			SelectFile();
		}

		public ExcelReader(string _fileName) {
			xlsFileInfo = new FileInfo(_fileName);
			initialDir = string.Format(@"{0}\..", xlsFileInfo.DirectoryName);
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
							currentPageInfo = new PageInfo(cell1_, cell2_, cell3_);
							foreach (string dir_ in Directory.GetDirectories(initialDir)) {
								string [] f_ = Directory.GetFiles(dir_, string.Format(@"FB.{0}.{1}", cell1_, Extension));
								if (f_.Length > 0) {
									currentPageInfo.fileInfo = new FileInfo(f_[0]);
									break;
								} else {
									string[] f3_ = Directory.GetFiles(dir_, string.Format(@"{0}.{1}", cell3_, Extension));
									if (f3_.Length > 0) {
										currentPageInfo.fileInfo = new FileInfo(f3_[0]);
										break;
									}
								}
							}
							subSections.AddLast(currentPageInfo);
						} else if (cell2_ != string.Empty && cell3_ != string.Empty) {
							currentPageInfo.Add(cell2_, cell3_);
						}
					}
				}
			} catch (IOException) {
				MessageBox.Show("Could not read file.\nIs it open in another window?", @"IO Exception",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error,
					MessageBoxDefaultButton.Button1,
					MessageBoxOptions.DefaultDesktopOnly);
			}
			return subSections;
		}
	}
}
