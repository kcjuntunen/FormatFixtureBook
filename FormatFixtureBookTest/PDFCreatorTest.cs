using System.IO;
using FormatFixtureBook;
using Microsoft.VisualStudio.TestTools.UnitTesting;
#if LONG_TESTS
using SolidWorks.Interop.sldworks;
#endif

namespace FormatFixtureBookTests {
	[TestClass]
	public class PDFCreatorTest {
		[TestMethod]
		public void MergePDFsFromSection1Test() {
			ExcelReader er_ = new ExcelReader(@"C:\Users\juntunenkc\Desktop\pdf test\Section 1.xlsx",
				ExcelReader.ExcelReaderExtensionOptions.PDF, ExcelReader.ExcelReaderSearchOptions.THIS_DIR);
			var ll_ = er_.ReadFile();
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\pdf test\test.pdf"));
		}

		[TestMethod]
		public void FilesNotFoundTest() {
			ExcelReader er_ = new ExcelReader(new FileInfo(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 1\Section 1.xlsx"),
				ExcelReader.ExcelReaderExtensionOptions.PDF, ExcelReader.ExcelReaderSearchOptions.THIS_DIR);
			var ll_ = er_.ReadFile();
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\pdf test\test.pdf"));
		}

#if LONG_TESTS
		[TestMethod]
		public void CreateAndMergePDFsFromSection1Test() {
			ExcelReader er_ = new ExcelReader(@"C:\Users\juntunenkc\Desktop\pdf test\Section 1.xlsx");
			var ll_ = er_.ReadFile();
			SldWorks swApp_ = null;
			try {
				swApp_ = System.Runtime.InteropServices.Marshal.GetActiveObject(@"SldWorks.Application") as SldWorks;
			} catch (System.Runtime.InteropServices.COMException e) {
				System.Console.WriteLine(e.Message);
			}
			PDFCreator p = new PDFCreator();
			p.CreateDrawings(swApp_, ll_);

			var nd_ = ll_.First;
			while (nd_ != null) {
				nd_.Value.fileInfo = new FileInfo(nd_.Value.fileInfo.FullName.Replace(@".SLDDRW", @".PDF"));
				nd_ = nd_.Next;
			}
			System.Runtime.InteropServices.Marshal.ReleaseComObject(swApp_);
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\pdf test\test.pdf"));
		}
#endif
	}
}
