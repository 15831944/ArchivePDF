using System.IO;

using iTextSharp.text.pdf;
using iTextSharp.text;

namespace ArchivePDF.csproj {
	public class UsefulInformationDestroyer {
		private string source;
		private string target;
		public UsefulInformationDestroyer(string sourcePath, string targetPath) {
			source = sourcePath;
			target = targetPath;
		}


		public void RedactUsefulDataFromDocument() {
			try {
				byte[] ba = _RedactUsefulDataFromDocument();
				using (FileStream fs = File.Create(target)) {
					for (int i = 0; i < ba.Length; i++) {
						fs.WriteByte(ba[i]);
					}
					fs.Close();
				}
			} catch (IOException ioe) {
				System.Windows.Forms.MessageBox.Show(ioe.Message, @"Couldn't create redacted version.");
			} catch (System.Exception e) {
				using (ErrMsg er_ = new ErrMsg(e)) {
					er_.ShowDialog();
				}
			}
		}

		private byte[] _RedactUsefulDataFromDocument() {
			using (Document document = new Document()) {
				using (MemoryStream memoryStream = new MemoryStream()) {
					try {
						PdfCopy copy = new PdfCopy(document, memoryStream);
						document.Open();
						using (PdfReader pdfReader = new PdfReader(source)) {
							for (int i = 1; i < pdfReader.NumberOfPages + 1; i++) {
								PdfImportedPage importedPage = copy.GetImportedPage(pdfReader, i);
								PdfCopy.PageStamp pageStamp = copy.CreatePageStamp(importedPage);
								PdfContentByte contentByte = pageStamp.GetOverContent();

								if (i == 1) {
									float y_offset = 137F;
									for (int j = 0; j < 5; j++, y_offset -= 13.5F) {
										Rectangle sourceRevRectangle = new Rectangle(21, y_offset, 44, y_offset + 10);
										sourceRevRectangle.BackgroundColor = BaseColor.WHITE;
										Rectangle sourceECRRectangle = new Rectangle(46, y_offset, 71, y_offset + 10);
										sourceECRRectangle.BackgroundColor = BaseColor.WHITE;
										Rectangle sourceDescrRectangle = new Rectangle(73, y_offset, 988, y_offset + 10);
										sourceDescrRectangle.BackgroundColor = BaseColor.WHITE;

										contentByte.Rectangle(sourceRevRectangle);
										contentByte.Rectangle(sourceECRRectangle);
										contentByte.Rectangle(sourceDescrRectangle);

										pageStamp.AlterContents();

									}
								}
								copy.AddPage(importedPage);
							}
						}
					} catch (System.Exception e) {
						throw e;
					} finally {
						document.Close();
					}
					return memoryStream.GetBuffer();
				}
			}
		}
	}
}
