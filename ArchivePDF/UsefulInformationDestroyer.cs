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
				using (ErrMsg em = new ErrMsg(ioe)) {
					em.ShowDialog();
				}
			}
		}

		private byte[] _RedactUsefulDataFromDocument() {
			using (Document document = new Document()) {
				using (MemoryStream memoryStream = new MemoryStream()) {
					PdfCopy copy = new PdfCopy(document, memoryStream);
					document.Open();
					using (PdfReader pdfReader = new PdfReader(source)) {
						for (int i = 1; i < pdfReader.NumberOfPages + 1; i++) {
							PdfImportedPage importedPage = copy.GetImportedPage(pdfReader, i);
								PdfCopy.PageStamp pageStamp = copy.CreatePageStamp(importedPage);
								PdfContentByte contentByte = pageStamp.GetOverContent();

							if (i == 1) {
								//Rectangle sourceRevRectangle = new Rectangle(21, 123, 44, 133);
								//sourceRevRectangle.BackgroundColor = BaseColor.BLACK;
								//Rectangle sourceECRRectangle = new Rectangle(46, 123, 71, 133);
								//sourceECRRectangle.BackgroundColor = BaseColor.BLACK;
								//Rectangle sourceDescrRectangle = new Rectangle(73, 123, 900, 133);
								//sourceDescrRectangle.BackgroundColor = BaseColor.BLACK;

								Rectangle sourceRevRectangle = new Rectangle(21, 137, 44, 147);
								sourceRevRectangle.BackgroundColor = BaseColor.BLACK;
								Rectangle sourceECRRectangle = new Rectangle(46, 137, 71, 147);
								sourceECRRectangle.BackgroundColor = BaseColor.BLACK;
								Rectangle sourceDescrRectangle = new Rectangle(73, 137, 900, 147);
								sourceDescrRectangle.BackgroundColor = BaseColor.BLACK;

								contentByte.Rectangle(sourceRevRectangle);
								contentByte.Rectangle(sourceECRRectangle);
								contentByte.Rectangle(sourceDescrRectangle);
								pageStamp.AlterContents();
							}
							copy.AddPage(importedPage);
						}
						document.Close();
						return memoryStream.GetBuffer();
					}
				}
			}
		}
	}
}
