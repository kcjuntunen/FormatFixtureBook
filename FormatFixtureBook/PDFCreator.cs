using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace FormatFixtureBook {
	public class PDFCreator {
		public static void CreateDrawings(SldWorks _swApp, LinkedList<PageInfo> _ll) {
			int dt = (int)swDocumentTypes_e.swDocDRAWING;
			int err = 0;
			int warn = 0;
			int saveVersion = (int)swSaveAsVersion_e.swSaveAsCurrentVersion;
			int saveOptions = (int)swSaveAsOptions_e.swSaveAsOptions_Silent;
			bool success;

			var nd_ = _ll.First;

			while (nd_ != null) {
				FileInfo slddrw_ = nd_.Value.fileInfo;
				string newName = slddrw_.Name.Replace(@".SLDDRW", @".PDF");
				FileInfo tmpFile = new FileInfo(string.Format(@"{0}\{1}", Path.GetTempPath(), newName));
				_swApp.OpenDocSilent(slddrw_.FullName, dt, ref err);
				_swApp.ActivateDoc3(slddrw_.FullName, true, 
					(int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref err);
				success = (_swApp.ActiveDoc as ModelDoc2).SaveAs4(tmpFile.FullName, saveVersion, saveOptions, ref err, ref warn);
				nd_.Value.fileInfo = tmpFile;
				_swApp.CloseDoc(slddrw_.FullName);
				nd_ = nd_.Next;
			}
		}

		public static void Merge(LinkedList<PageInfo> _ll, FileInfo _target) {
			byte[] ba = merge_files(_ll);
			using (FileStream fs = File.Create(_target.FullName)) {
				for (int i = 0; i < ba.Length; i++) {
					fs.WriteByte(ba[i]);
				}
				fs.Close();
			}
		}

		private static int count_pages(LinkedList<PageInfo> docs) {
			int total = 0;
			var nd_ = docs.First;
			while (docs != null) {
				PdfReader rdr_ = new PdfReader(nd_.Value.fileInfo.FullName);
				total += rdr_.NumberOfPages;
				nd_ = nd_.Next;
			}
			return total;
		}

		private static byte[] merge_files(LinkedList<PageInfo> ll) {
			using (Document document = new Document()) {
				using (MemoryStream ms = new MemoryStream()) {
					PdfCopy copy = new PdfCopy(document, ms);
					document.Open();
					int document_page_counter = 0;
					int total_pages = count_pages(ll);
					int font_size = 31;

					var nd_ = ll.First;
					while (nd_ != null) {
						using (PdfReader rdr_ = new PdfReader(nd_.Value.fileInfo.FullName)) {
							var sn_ = nd_.Value.FirstSheetNo();
							int pg_ = 0;
							while (sn_ != null) {
								document_page_counter++;
								PdfImportedPage ip_ = copy.GetImportedPage(rdr_, pg_++);
								PdfCopy.PageStamp ps_ = copy.CreatePageStamp(ip_);
								PdfContentByte cb_ = ps_.GetOverContent();
								Rectangle r_ = new Rectangle(720, 20, 47, 16);
								r_.BackgroundColor = BaseColor.WHITE;
								cb_.Rectangle(r_);
								Font f_ = FontFactory.GetFont(@"Arial", font_size);
								Chunk c_ = new Chunk(sn_.Value.ToString(), f_);
								c_.SetBackground(BaseColor.WHITE);
								ColumnText.ShowTextAligned(cb_,
									Element.ALIGN_CENTER,
									new Phrase(c_),
									20, 20,
									ip_.Width < ip_.Height ? 0 : 1);
								ps_.AlterContents();
								sn_ = sn_.Next;

								copy.AddPage(ip_);
							}
						}
						nd_ = nd_.Next;
						document.Close();
					}
					return ms.GetBuffer();
				}
			}
		}
	}
}
