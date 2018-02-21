using Microsoft.VisualStudio.TestTools.UnitTesting;
using FormatFixtureBook;
using System.Diagnostics;

namespace FormatFixtureBookTests {
	[TestClass]
	public class ExcelReaderTest {
		[TestMethod]
		public void RootTest() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\All.xlsx",
				ExcelReader.ExcelReaderExtensionOptions.SLDDRW,
				ExcelReader.ExcelReaderSearchOptions.THIS_DIR | ExcelReader.ExcelReaderSearchOptions.RECURSE);

			var ll_ = er_.ReadFile();
			var nd_ = ll_.First;

			while (nd_ != null) {
				var descr_ = nd_.Value.FirstDescription();
				var sht_ = nd_.Value.FirstSheetNo();

				while (descr_ != null && sht_ != null) {
					Debug.WriteLine(string.Format(@"{0,-60} | {1,-70} | {2,10}", nd_.Value.fileInfo, descr_.Value, sht_.Value));
					descr_ = descr_.Next;
					sht_ = sht_.Next;
				}
				nd_ = nd_.Next;
			}
		}


		[TestMethod]
		public void SubTest() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 1\Section 1.xlsx",
				ExcelReader.ExcelReaderExtensionOptions.SLDDRW,
				ExcelReader.ExcelReaderSearchOptions.PARENT_DIR | ExcelReader.ExcelReaderSearchOptions.RECURSE);

			var ll_ = er_.ReadFile();
			var nd_ = ll_.First;

			while (nd_ != null) {
				var descr_ = nd_.Value.FirstDescription();
				var sht_ = nd_.Value.FirstSheetNo();

				while (descr_ != null && sht_ != null) {
					Debug.WriteLine(string.Format(@"{0,-60} | {1,-70} | {2,10}", nd_.Value.fileInfo, descr_.Value, sht_.Value));
					descr_ = descr_.Next;
					sht_ = sht_.Next;
				}
				nd_ = nd_.Next;
			}

		}

		[TestMethod]
		public void OtherSubTest() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 5\Section 5.xlsx",
				ExcelReader.ExcelReaderExtensionOptions.SLDDRW,
				ExcelReader.ExcelReaderSearchOptions.PARENT_DIR | ExcelReader.ExcelReaderSearchOptions.RECURSE);

			var ll_ = er_.ReadFile();
			var nd_ = ll_.First;

			while (nd_ != null) {
				var descr_ = nd_.Value.FirstDescription();
				var sht_ = nd_.Value.FirstSheetNo();

				while (descr_ != null && sht_ != null) {
					Debug.WriteLine(string.Format(@"{0,-60} | {1,-70} | {2,10}", nd_.Value.fileInfo, descr_.Value, sht_.Value));
					descr_ = descr_.Next;
					sht_ = sht_.Next;
				}
				nd_ = nd_.Next;
			}
		}

		public void ConstructBadReader() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 5\Section 5.csv");
		}

		[TestMethod]
		public void WrongExtensionExceptionTest() {
			System.Action badReader = ConstructBadReader;
			Assert.ThrowsException<ExcelReaderException>((System.Action)ConstructBadReader);
		}

		public void ReadOpenFile() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 5\Section 5.xlsx");
			var ll_ = er_.ReadFile();
		}

		[TestMethod]
		public void FileOpenExceptionTest() {
			Microsoft.Office.Interop.Excel.Application x_ = new Microsoft.Office.Interop.Excel.Application();
			x_.Workbooks.Open(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 5\Section 5.xlsx");
			Assert.ThrowsException<ExcelReaderException>((System.Action)ReadOpenFile);
			x_.Workbooks.Close();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(x_);
		}

		public void ReadBadFile() {
			ExcelReader er_ = new ExcelReader(@"C:\Users\juntunenkc\Desktop\pdf test\Bad Section 1.xlsx");
			var ll_ = er_.ReadFile();
		}

		[TestMethod]
		public void FoundNothingException() {
			Assert.ThrowsException<ExcelReaderFoundNoFilesException>((System.Action)ReadBadFile);
		}

		[TestMethod]
		public void TestEvent() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 5\Section 5.xlsx",
			 ExcelReader.ExcelReaderExtensionOptions.PDF,
			 ExcelReader.ExcelReaderSearchOptions.PARENT_DIR | ExcelReader.ExcelReaderSearchOptions.RECURSE);
			er_.NewDir += Er__NewDir;
			var ll_ = er_.ReadFile();
			Assert.IsTrue(testString.Length > 0);
		}

		private string testString = string.Empty;

		private void Er__NewDir(object sender, System.EventArgs e) {
			testString += (e as System.IO.FileSystemEventArgs).FullPath;
		}
	}
}
