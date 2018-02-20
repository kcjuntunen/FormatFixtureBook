using System.IO;
using FormatFixtureBook;
using SolidWorks.Interop.sldworks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReaderTest {
	[TestClass]
	public class PDFCreatorTest {
		[TestMethod]
		public void MergePDFsFromSection1Test() {
			ExcelReader er_ = new ExcelReader(@"C:\Users\juntunenkc\Desktop\pdf test\Section 1.xlsx");
			er_.Extension = @"PDF";
			var ll_ = er_.ReadFile();
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\pdf test\test.pdf"));
		}

		[TestMethod]
		public void CreateAndMergePDFsFromSection1Test() {
			ExcelReader er_ = new ExcelReader(@"C:\Users\juntunenkc\Desktop\pdf test\Section 1.xlsx");
			var ll_ = er_.ReadFile();
			SldWorks swApp_ = System.Runtime.InteropServices.Marshal.GetActiveObject(@"SldWorks.Application") as SldWorks;
			PDFCreator.CreateDrawings(swApp_, ll_);
			var nd_ = ll_.First;
			while (nd_ != null) {
				nd_.Value.fileInfo = new FileInfo(nd_.Value.fileInfo.FullName.Replace(@".SLDDRW", @".PDF"));
				nd_ = nd_.Next;
			}
			PDFCreator.Merge(ll_, new FileInfo(@"C:\Users\juntunenkc\Desktop\pdf test\test.pdf"));
		}
	}
}
