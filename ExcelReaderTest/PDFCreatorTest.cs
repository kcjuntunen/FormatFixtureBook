using System.IO;
using FormatFixtureBook;
using SolidWorks.Interop.sldworks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReaderTest {
	[TestClass]
	public class PDFCreatorTest {
		[TestMethod]
		public void CreatePDFsFromSection1Test() {
			ExcelReader er_ = new ExcelReader(@"G:\ZALES\FIXTURE BOOK\SECTIONS\SECTION 1\Section 1.xlsx");
			var ll_ = er_.ReadFile();
			SldWorks swApp_ = System.Runtime.InteropServices.Marshal.GetActiveObject(@"SldWorks.Application") as SldWorks;
			PDFCreator.CreateDrawings(swApp_, ll_);
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\test.pdf"));
		}
	}
}
