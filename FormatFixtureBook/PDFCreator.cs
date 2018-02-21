using System;
using System.IO;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace FormatFixtureBook {
	public class PDFCreator {
		public PDFCreator() {

		}

		public event EventHandler Opening;
		protected virtual void OnOpening(FileSystemEventArgs e) {
			Opening?.Invoke(this, e);
		}

		public event EventHandler Closing;
		protected virtual void OnClosing(FileSystemEventArgs e) {
			Closing?.Invoke(this, e);
		}

		public void CreateDrawings(SldWorks _swApp, LinkedList<PageInfo> _ll) {
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
				FileSystemEventArgs fsea_ =
					new FileSystemEventArgs(WatcherChangeTypes.All, Path.GetDirectoryName(tmpFile.FullName), tmpFile.Name);
				OnOpening(fsea_);
				_swApp.OpenDocSilent(slddrw_.FullName, dt, ref err);
				_swApp.ActivateDoc3(slddrw_.FullName, true, 
					(int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref err);
				success = (_swApp.ActiveDoc as ModelDoc2).SaveAs4(tmpFile.FullName, saveVersion, saveOptions, ref err, ref warn);
				nd_.Value.fileInfo = tmpFile;
				OnClosing(fsea_);
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
			while (nd_ != null) {
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
					int total_pages = count_pages(ll);

					Rectangle sheet_number_blockhout_ = new Rectangle(1062, 17, 1180, 110);
					sheet_number_blockhout_.BackgroundColor = BaseColor.WHITE;
					Rectangle revisions_blockout_ = new Rectangle(817, 17, 1032, 110);
					revisions_blockout_.BackgroundColor = BaseColor.WHITE;
					Rectangle item_descr_blockout_ = new Rectangle(605, 17, 787, 110);
					item_descr_blockout_.BackgroundColor = BaseColor.WHITE;

					var nd_ = ll.First;
					while (nd_ != null) {
						using (PdfReader rdr_ = new PdfReader(nd_.Value.fileInfo.FullName)) {
							var sn_ = nd_.Value.FirstSheetNo();
							var descr_ = nd_.Value.FirstDescription();
							int pg_ = 1;
							int document_page_counter = 0;
							while (sn_ != null) {
								PdfImportedPage ip_ = copy.GetImportedPage(rdr_, pg_++);
								PdfCopy.PageStamp ps_ = copy.CreatePageStamp(ip_);
								PdfContentByte cb_ = ps_.GetOverContent();
								cb_.Rectangle(sheet_number_blockhout_);
								cb_.Rectangle(revisions_blockout_);
								cb_.Rectangle(item_descr_blockout_);

								Font sheet_number_font_ = FontFactory.GetFont(@"Arial", 31);
								Chunk sheet_number_ = new Chunk(sn_.Value.ToString(), sheet_number_font_);
								sheet_number_.SetBackground(BaseColor.WHITE);
								ColumnText.ShowTextAligned(cb_,
									Element.ALIGN_CENTER,
									new Phrase(sheet_number_),
									1127, 60,
									0);

								if (nd_.Value.VendorInfo) {
									Font item_descr_font_ = FontFactory.GetFont(@"Tw Cen MT", 17, Font.BOLD);
									string desc_ = @"VENDOR INFO";
									Chunk item_descr_ = new Chunk(desc_, item_descr_font_);
									sheet_number_.SetBackground(BaseColor.WHITE);
									ColumnText.ShowTextAligned(cb_,
										Element.ALIGN_CENTER,
										new Phrase(item_descr_),
										700, 60,
										0);

								} else {
									Font item_font_ = FontFactory.GetFont(@"Tw Cen MT", 17, Font.BOLD);
									string item_name_ = string.Format("{0}", nd_.Value.Name);
									Chunk item_ = new Chunk(item_name_, item_font_);
									sheet_number_.SetBackground(BaseColor.WHITE);
								 ColumnText.ShowTextAligned(cb_,
									 Element.ALIGN_CENTER,
									 new Phrase(item_),
									 700, 68,
									 0);

									Font item_descr_font_ = FontFactory.GetFont(@"Tw Cen MT", 12, Font.BOLD);
									string desc_ = string.Format(@"{0}", descr_.Value);
									if (document_page_counter++ > 0 && rdr_.NumberOfPages > 1) {
										desc_ = string.Format(@"SEE SHEET {0}", nd_.Value.FirstSheetNo().Value);
									}
									Chunk item_descr_ = new Chunk(desc_, item_descr_font_);
									sheet_number_.SetBackground(BaseColor.WHITE);
									ColumnText.ShowTextAligned(cb_,
										Element.ALIGN_CENTER,
										new Phrase(item_descr_),
										700, 50,
										0);
								}
								ps_.AlterContents();
								sn_ = sn_.Next;
								descr_ = descr_.Next;
								copy.AddPage(ip_);
							}
						}
						nd_ = nd_.Next;
					}
					document.Close();
					return ms.GetBuffer();
				}
			}
		}
	}
}
