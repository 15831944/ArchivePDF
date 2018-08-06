using System;
using System.Collections.Generic;
using System.IO;

using System.Runtime.InteropServices;

using System.Data.SqlClient;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swcommands;

namespace ArchivePDF.csproj {
	class PDFArchiver {
		private SldWorks swApp;
		private ModelDoc2 swModel;
		private Frame swFrame;
		private DrawingDoc swDraw;
		private View swView;
		private CustomPropertyManager swCustPropMgr;
		private swDocumentTypes_e modelType = swDocumentTypes_e.swDocNONE;
		private String sourcePath = String.Empty;
		private String drawingPath = String.Empty;
		private int drwKey = 0;
		public bool IsTarget = false;
		public string savedFile = string.Empty;
		private bool metal;
		//private Boolean shouldCheck = true;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="sw">Requires a <see cref="SolidWorks.Interop.sldworks.SldWorks"/> type.</param>
		/// <param name="ps">Requires a <see cref="ArchivePDF.csproj.PathSet"/></param>
		public PDFArchiver(ref SldWorks sw, ArchivePDF.csproj.PathSet ps) {
			swApp = sw;
			APathSet = ps;
			swModel = (ModelDoc2)swApp.ActiveDoc;
			ModelDocExtension ex = (ModelDocExtension)swModel.Extension;
			string lvl = GetRev(ex);
			if (DocCheck()) {
				if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING) {
					throw new ExportPDFException("You must have a Drawing Document open.");
				}
				swDraw = (DrawingDoc)swModel;
				swFrame = (Frame)swApp.Frame();

				swView = GetFirstView(swApp);
				metal = IsMetal(swView);
				sourcePath = swView.GetReferencedModelName().ToUpper().Trim();
				drawingPath = swModel.GetPathName().ToUpper().Trim();

				if (sourcePath.Contains("SLDASM")) {
					modelType = swDocumentTypes_e.swDocASSEMBLY;

				} else if (sourcePath.Contains("SLDPRT")) {
					modelType = swDocumentTypes_e.swDocPART;
				} else {
					modelType = swDocumentTypes_e.swDocNONE;
				}
			} else {
				MustSaveException e = new MustSaveException("The drawing has to be saved.");
				throw e;
			}
		}

		public View GetFirstView(SldWorks sw) {
			ModelDoc2 swModel = (ModelDoc2)sw.ActiveDoc;
			DrawingDoc d = (DrawingDoc)swModel;
			View v;
			string[] shtNames = (String[])swDraw.GetSheetNames();
			string message = string.Empty;

			//This should find the first page with something on it.
			IsTarget = IsTargetDrawing((sw.ActiveDoc as DrawingDoc).Sheet[shtNames[0]]);
			int x = 0;
			do {
				try {
					d.ActivateSheet(shtNames[x]);
				} catch (IndexOutOfRangeException e) {
					throw new IndexOutOfRangeException("Went beyond the number of sheets.", e);
				} catch (Exception e) {
					throw e;
				}
				v = (View)d.GetFirstView();
				v = (View)v.GetNextView();
				x++;
			} while ((v == null) && (x < d.GetSheetCount()));

			message = (string)v.GetName2() + ":\n";

			if (v == null) {
				throw new Exception("I couldn't find a model anywhere in this document.");
			}
			return v;
		}

		public void CorrectLayers(string rev, selectLayer f) {
			Sheet curSht = (Sheet)swDraw.GetCurrentSheet();
			string[] shts = (string[])swDraw.GetSheetNames();
			foreach (string s in shts) {
				swFrame.SetStatusBarText("Showing correct revs on " + s + "...");
				swDraw.ActivateSheet(s);
				f(rev);
			}
			swDraw.ActivateSheet(curSht.GetName());
		}

		public delegate void selectLayer(string rev);

		/// <summary>
		/// Selects the correct layer.
		/// </summary>
		public void chooseLayer(string rev) {
			char c = (char)rev[2];
			int revcount = (int)c - 64;

			LayerMgr lm = (LayerMgr)(swApp.ActiveDoc as ModelDoc2).GetLayerManager();
			string head = getLayerNameHeader(lm);

			for (int i = 0; i < Properties.Settings.Default.LayerTails.Count; i++) {
				string currentTail = Properties.Settings.Default.LayerTails[i];
				try {
					Layer l = (Layer)lm.GetLayer(string.Format("{0}{1}", head, currentTail));
					if (l != null) {
						l.Visible = false;
						if (Math.Floor((double)((revcount - 1) / 5)) == i) {
							l.Visible = true;
						}
					}
				} catch (Exception) {
					// Sometimes the layer doesn't exist.
				}
			}
		}

		private string getLayerNameHeader(LayerMgr lm) {
			foreach (string h in Properties.Settings.Default.LayerHeads) {
				foreach (string t in Properties.Settings.Default.LayerTails) {
					Layer l = (Layer)lm.GetLayer(string.Format("{0}{1}", h, t));
					if (l != null && l.Visible)
						return h;
				}
			}
			return "AMS";
		}

		private bool IsTargetDrawing(Sheet sheet) {
			if (sheet.GetTemplateName().ToUpper().Contains(@"TARGET_ASS")) {
				return true;
			}
			return false;
		}

		public bool IsMetal(View v) {
			ModelDoc2 md = (ModelDoc2)v.ReferencedDocument;
			ConfigurationManager cfMgr = md.ConfigurationManager;
			Configuration cf = cfMgr.ActiveConfiguration;

			CustomPropertyManager gcpm = md.Extension.get_CustomPropertyManager(string.Empty);
			CustomPropertyManager scpm;

			string _value = "WOOD";
			string _resValue = string.Empty;
			bool wasResolved;
			bool useCached = false;

			if (cf != null) {
				scpm = cf.CustomPropertyManager;
			} else {
				scpm = gcpm;
			}
			int res;

			res = gcpm.Get5("DEPARTMENT", useCached, out _value, out _resValue, out wasResolved);
			if (_value == string.Empty) {
				res = gcpm.Get5("DEPTID", useCached, out _value, out _resValue, out wasResolved);
				if (_value == string.Empty) {
					res = scpm.Get5("DEPARTMENT", useCached, out _value, out _resValue, out wasResolved);
				}
			}

			if (_value == "2" || _value.ToUpper() == "METAL") {
				return true;
			}

			return false;
		}

		/// <summary>
		/// Exports current drawing to PDFs.
		/// </summary>
		/// <returns>Returns a <see cref="System.Boolean"/>, should you need one.</returns>
		public Boolean ExportPdfs() {
			String pdfSourceName = Path.GetFileNameWithoutExtension(drawingPath);
			String pdfAltPath = Path.GetDirectoryName(drawingPath).Substring(3);

			String fFormat = String.Empty;
			ModelDocExtension swModExt = swModel.Extension;
			String Rev = GetRev(swModExt);
			if (Rev.Length > 2)
				CorrectLayers(Rev, chooseLayer);

			List<String> pdfTarget = new List<string>();

			if (!drawingPath.StartsWith(Properties.Settings.Default.SMapped.ToUpper())) {
				pdfTarget.Add(String.Format("{0}{1}\\{2}.PDF", APathSet.KPath, pdfAltPath, pdfSourceName));
				pdfTarget.Add(String.Format("{0}{1}\\{2}{3}.PDF", APathSet.GPath, pdfAltPath, pdfSourceName, Rev));

				Boolean success = SaveFiles(pdfTarget);

				if (metal && APathSet.WriteToDb && !_metalDrawing && find_pdf(pdfSourceName))
					AlertSomeone(pdfTarget[0]);

				return success;
			} else {
				Boolean success = ExportMetalPdfs();
				return success;
			}
		}

		private bool find_pdf(string doc) {
			string searchterm_ = string.Format(@"{0}.PDF", doc);
			using (SqlConnection conn = new SqlConnection(Properties.Settings.Default.connectionString)) {
				string sql_ = @"SELECT FileID, FName, FPath, DateCreated FROM GEN_DRAWINGS_MTL WHERE(FName = @fname)";
				using (SqlCommand comm = new SqlCommand(sql_, conn)) {
					comm.Parameters.AddWithValue(@"@fname", searchterm_);
					try {
						if (conn.State == System.Data.ConnectionState.Closed) {
							conn.Open();
						}

						using (SqlDataReader reader_ = comm.ExecuteReader()) {
							return reader_.HasRows;
						}
					} catch (InvalidOperationException ioe) {
						throw new ExportPDFException(@"I just didn't want to open the connection.");
					}
				}
			}
		}

		public Boolean ExportMetalPdfs() {
			String pdfSourceName = Path.GetFileNameWithoutExtension(drawingPath);
			String pdfAltPath = Path.GetDirectoryName(drawingPath).Substring(44);
			//String pdfAltPath = ".\\";

			String fFormat = String.Empty;
			ModelDocExtension swModExt = swModel.Extension;
			List<String> pdfTarget = new List<string>();

			pdfTarget.Add(String.Format("{0}{1}\\{2}.PDF", APathSet.MetalPath, pdfAltPath, pdfSourceName));
			Boolean success = SaveFiles(pdfTarget);
			return success;
		}

		/// <summary>
		/// Opens and saves models referenced in the open drawing.
		/// </summary>
		/// <returns>Returns a <see cref="System.Boolean"/>, should you need one.</returns>
		public Boolean ExportEDrawings() {
			ModelDoc2 currentDoc = swModel;
			string docName = currentDoc.GetPathName();
			String sourceName = Path.GetFileNameWithoutExtension(docName);
			String altPath = Path.GetDirectoryName(docName).Substring(3);
			ModelDocExtension swModExt = swModel.Extension;
			String Rev = GetRev(swModExt);
			String fFormat = String.Empty;
			List<string> ml = get_list_of_open_docs();
			List<String> target = new List<string>();
			Boolean measurable = true;
			Int32 options = 0;
			Int32 errors = 0;

			switch (modelType) {
				case swDocumentTypes_e.swDocASSEMBLY:
					fFormat = ".EASM";
					break;
				case swDocumentTypes_e.swDocPART:
					fFormat = ".EPRT";
					break;
				default:
					ExportPDFException e = new ExportPDFException("Document type error.");
					//e.Data.Add("who", System.Environment.UserName);
					//e.Data.Add("when", DateTime.Now);
					throw e;
			}

			swApp.ActivateDoc3(sourcePath, true, options, ref errors);
			swModel = (ModelDoc2)swApp.ActiveDoc;
			Configuration swConfig = (Configuration)swModel.GetActiveConfiguration();
			swFrame.SetStatusBarText("Positioning model.");
			swModel.ShowNamedView2("*Dimetric", 9);
			swModel.ViewZoomtofit2();

			if (!swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)) {
				swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, true);
				measurable = false;
			} else {
				measurable = true;
			}

			target.Add(String.Format("{0}{1}\\{2}{3}", APathSet.KPath, altPath, sourceName, fFormat));
			target.Add(String.Format("{0}{1}\\{2}{3}{4}", APathSet.GPath, altPath, sourceName, Rev, fFormat));

			Boolean success = SaveFiles(target);

			if (!measurable)
				swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, false);

			if (!ml.Contains(swModel.GetPathName())) {
				swFrame.SetStatusBarText("Closing model.");
				swApp.CloseDoc(sourcePath);
			} else if (APathSet.SaveFirst) {
				swModel.SaveSilent();
			}
			int err = 0;
			swApp.ActivateDoc3(currentDoc.GetTitle(), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref err);

			return success;
		}
		/// <summary>
		/// This thing does the saving.
		/// </summary>
		/// <param name="fl">A <see cref="System.Collections.Generic.List<>"/> of destination paths.</param>
		/// <returns>Returns a <see cref="System.Boolean"/>, indicating success.</returns>
		private Boolean SaveFiles(List<String> fl) {
			Int32 saveVersion = (Int32)swSaveAsVersion_e.swSaveAsCurrentVersion;
			Int32 saveOptions = (Int32)swSaveAsOptions_e.swSaveAsOptions_Silent;
			Int32 refErrors = 0;
			Int32 refWarnings = 0;
			Boolean success = true;
			string tmpPath = Path.GetTempPath();
			ModelDocExtension swModExt = default(ModelDocExtension);
			ExportPdfData swExportPDFData = default(ExportPdfData);

			foreach (String fileName in fl) {
				FileInfo fi = new FileInfo(fileName);
				string tmpFile = tmpPath + "\\" + fi.Name;
				if (drawingPath != String.Empty) {
					swFrame.SetStatusBarText(String.Format("Checking path: '{0}'", fileName));
					if (!CreatePath(fileName)) {
						ExportPDFException e = new ExportPDFException("Unable to save file, folder could not be created.");
						//e.Data.Add("who", System.Environment.UserName);
						//e.Data.Add("when", DateTime.Now);
						throw e;
					}

					String[] obj = (string[])swDraw.GetSheetNames();
					object[] objs = new object[obj.Length - 1];
					DispatchWrapper[] dr = new DispatchWrapper[obj.Length - 1];
					for (int i = 0; i < obj.Length - 1; i++) {
						swDraw.ActivateSheet(obj[i]);
						Sheet s = (Sheet)swDraw.GetCurrentSheet();
						objs[i] = s;
						dr[i] = new DispatchWrapper(objs[i]);
					}

					swFrame.SetStatusBarText(String.Format("Exporting '{0}'", fileName));
					bool layerPrint = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportIncludeLayersNotToPrint);
					swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportIncludeLayersNotToPrint, true);
					swExportPDFData = swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
					swModExt = swModel.Extension;
					success = swExportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, (dr));
					success = swModExt.SaveAs(tmpFile, saveVersion, saveOptions, swExportPDFData, ref refErrors, ref refWarnings);
					//success = swModel.SaveAs4(tmpFile, saveVersion, saveOptions, ref refErrors, ref refWarnings);

					swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportIncludeLayersNotToPrint, layerPrint);

					try {
						File.Copy(tmpFile, fileName, true);
					} catch (UnauthorizedAccessException uae) {
						throw new ExportPDFException(
								String.Format("You don't have the reqired permission to access '{0}'.", fileName),
								uae);
					} catch (ArgumentException ae) {
						throw new ExportPDFException(
								String.Format("Either '{0}' or '{1}' is not a proper file name.", tmpFile, fileName),
								ae);
					} catch (PathTooLongException ptle) {
						throw new ExportPDFException(
								String.Format("Source='{0}'; Dest='{1}' <= One of these is too long.", tmpFile, fileName),
								ptle);
					} catch (DirectoryNotFoundException dnfe) {
						throw new ExportPDFException(
								String.Format("Source='{0}'; Dest='{1}' <= One of these is invalid.", tmpFile, fileName),
								dnfe);
					} catch (FileNotFoundException fnfe) {
						throw new ExportPDFException(
								String.Format("Crap! I lost '{0}'!", tmpFile),
								fnfe);
					} catch (IOException) {
						System.Windows.Forms.MessageBox.Show(
								String.Format("If you have the file, '{0}', selected in an Explorer window, " +
								"you may have to close it.", fileName), "This file is open somewhere.",
								System.Windows.Forms.MessageBoxButtons.OK,
								System.Windows.Forms.MessageBoxIcon.Error);
						return false;
					} catch (NotSupportedException nse) {
						throw new ExportPDFException(
								String.Format("Source='{0}'; Dest='{1}' <= One of these is an invalid format.",
								tmpFile, fileName), nse);
					}


					if (!File.Exists(fileName))
						success = false;

					if (success) {
						swFrame.SetStatusBarText(String.Format("Exported '{0}'", fileName));

						if ((fileName.StartsWith(Properties.Settings.Default.KPath) && fileName.EndsWith("PDF")) && APathSet.WriteToDb) {
							savedFile = fileName;
							InsertIntoDb(fileName, Properties.Settings.Default.table);
							drwKey = GetKeyCol(fi.Name);
						}

						if ((fileName.StartsWith(Properties.Settings.Default.MetalPath) && fileName.EndsWith("PDF")) && APathSet.WriteToDb) {
							_metalDrawing = true;
							InsertIntoDb(fileName, Properties.Settings.Default.metalTable);
							drwKey = GetKeyCol(fi.Name);
						}

					} else {
						ExportPDFException e = new ExportPDFException(String.Format("Failed to save '{0}'", fileName));
						//e.Data.Add("who", System.Environment.UserName);
						//e.Data.Add("when", DateTime.Now);
						throw e;
					}
				}
			}
			if (success)
				return true;
			else
				return false;
		}

		private String GetRev(ModelDocExtension swModExt) {
			swCustPropMgr = (CustomPropertyManager)swModExt.get_CustomPropertyManager("");
			String[] propNames = (String[])swCustPropMgr.GetNames();
			String ValOut = String.Empty;
			String ResolvedValOut = String.Empty;
			Boolean WasResolved = false;

			String result = String.Empty;

			foreach (String name in propNames) {
				swCustPropMgr.Get5(name, false, out ValOut, out ResolvedValOut, out WasResolved);
				if (name.Contains("REVISION") && !name.Contains(@"LEVEL") && ValOut != String.Empty)
					result = "-" + ValOut;
			}

			if (result.Length != 3) {
				MustHaveRevException e = new MustHaveRevException("Check to make sure drawing is at least revision AA or later.");
				//e.Data.Add("who", System.Environment.UserName);
				//e.Data.Add("when", DateTime.Now);
				//e.Data.Add("result", result);
				throw e;
			}

			//System.Diagnostics.Debug.Print(result);
			return result;
		}

		private Boolean DocCheck() {
			if (swModel == null) {
				ExportPDFException e = new ExportPDFException("You must have a drawing document open.");
				return false;
			} else if ((Int32)swModel.GetType() != (Int32)swDocumentTypes_e.swDocDRAWING) {
				ExportPDFException e = new ExportPDFException("You must have a drawing document open.");
				return false;
			} else if (swModel.GetPathName() == String.Empty) {
				swModel.Extension.RunCommand((int)swCommands_e.swCommands_SaveAs, swModel.GetTitle());
				if (swModel.GetPathName() == string.Empty)
					return false;

				return true;
			} else
				return true;
		}

		private Boolean CreatePath(String pathCheck) {
			pathCheck = Path.GetDirectoryName(pathCheck);
			String targetpath = String.Empty;

			if (!Directory.Exists(pathCheck)) {
				String msg = String.Format("'{0}' does not exist.  Do you wish to create this folder?", pathCheck);
				System.Windows.Forms.MessageBoxButtons mbb = System.Windows.Forms.MessageBoxButtons.YesNo;
				System.Windows.Forms.MessageBoxIcon mbi = System.Windows.Forms.MessageBoxIcon.Question;
				System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show(msg, "New folder", mbb, mbi);
				if (dr == System.Windows.Forms.DialogResult.Yes) {
					DirectoryInfo di = Directory.CreateDirectory(pathCheck);
					return di.Exists;
				} else
					return false;
			}
			return true;
		}

		private ArchivePDF.csproj.PathSet _ps;

		public ArchivePDF.csproj.PathSet APathSet {
			get { return _ps; }
			set { _ps = value; }
		}

		public bool drawing_is_open(string m) {
			return get_list_of_open_docs().Contains(m);
		}

		public List<string> get_list_of_open_docs() {
			ModelDoc2 temp = (ModelDoc2)swApp.GetFirstDocument();
			List<string> ml = new List<string>();
			temp.GetNext();
			while (temp != null) {
				string temp_string = temp.GetPathName();
				if (temp.Visible == true && !ml.Contains(temp_string)) {
					ml.Add(temp.GetPathName());
				}
				temp = (ModelDoc2)temp.GetNext();
			}
			return ml;
		}

		public static void _sendErrMessage(Exception e, SldWorks sw) {
			String msg = String.Empty;
			msg = String.Format("{0} threw error:\n{1}", e.TargetSite, e.Message);
			msg += String.Format("\n\nStack trace:\n{0}", e.StackTrace);

			if (e.Data.Count > 0) {
				msg += "\n\nData:\n";

				foreach (System.Collections.DictionaryEntry de in e.Data) {
					msg += String.Format("{0}: {1}\n", de.Key, de.Value);
				}
			}

			sw.SendMsgToUser2(msg, (Int32)swMessageBoxIcon_e.swMbInformation, (Int32)swMessageBoxBtn_e.swMbOk);
		}

		protected virtual bool FileIsInUse(string fileName) {
			FileInfo file = new FileInfo(fileName);
			FileStream stream = null;

			try {
				stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
			} catch (IOException) {
				//the file is unavailable because it is:
				//still being written to
				//or being processed by another thread
				//or does not exist (has already been processed)
				return true;
			} finally {
				if (stream != null)
					stream.Close();
			}
			return false;
		}

		private bool DoInsert(string fileName, string table) {
			FileInfo fi = new FileInfo(fileName);
			SqlConnection connection = new SqlConnection(Properties.Settings.Default.connectionString);
			int insRes = 0;

			try {
				connection.Open();
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException(@"The connection is already open, or information is missing from the connection string: '" +
						Properties.Settings.Default.connectionString + "'.", ioe);
			} catch (SqlException se) {
				throw new ExportPDFException(@"A connection-level error occurred while opening the connection, whatever that means.", se);
			} catch (Exception e) {
				throw new ExportPDFException(@"Whoops. There's a problem", e);
			}

			try {
				string sql = String.Format(@"UPDATE {0} SET {1} = @fpath, DateCreated = @fdate WHERE ((({0}.{2}) = @filename))",
						table,
						Properties.Settings.Default.fullpath,
						Properties.Settings.Default.basename,
						fi.Name);

				swFrame.SetStatusBarText(sql);

				SqlCommand command = new SqlCommand(sql, connection);
				command.Parameters.AddWithValue("@fpath", fi.DirectoryName
									.Replace(Properties.Settings.Default.KPath, Properties.Settings.Default.KMapped)
									.Replace(Properties.Settings.Default.MetalPath, Properties.Settings.Default.SMapped) + @"\");
				command.Parameters.AddWithValue("@fdate", DateTime.Now);
				command.Parameters.AddWithValue("@filename", fi.Name);

				int aff = command.ExecuteNonQuery();
				if (aff > 0) {
					return true;
				} else {
					sql = String.Format(@"INSERT INTO {0} ({1}, {2}, {3}) VALUES (@fname, @pname, @date);",
							table,
							Properties.Settings.Default.basename,
							Properties.Settings.Default.fullpath,
							Properties.Settings.Default.datecreated);

					swFrame.SetStatusBarText(sql);
					command = new SqlCommand(sql, connection);
					command.Parameters.AddWithValue("@fname", fi.Name);
					command.Parameters.AddWithValue("@pname", fi.DirectoryName
									.Replace(Properties.Settings.Default.KPath, Properties.Settings.Default.KMapped)
									.Replace(Properties.Settings.Default.MetalPath, Properties.Settings.Default.SMapped) + @"\");
					command.Parameters.AddWithValue("@date", DateTime.Now);
					insRes = command.ExecuteNonQuery();
				}
			} catch (InvalidCastException ice) {
				throw new ExportPDFException("Invalid cast exception.", ice);
			} catch (SqlException se) {
				throw new ExportPDFException(String.Format("Couldn't execute query: {0}", se));
			} catch (IOException ie) {
				throw new ExportPDFException("IO exception", ie);
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException("The SqlConnection closed or dropped during a streaming operation.", ioe);
			} catch (Exception e) {
				throw new ExportPDFException("I looked for the connection, but it was gone.", e);
			} finally {
				connection.Close();
			}

			if (insRes > 0)
				return true;
			else
				return false;
		}

		private void InsertIntoDb(string fileName, string table) {
			FileInfo fi = new FileInfo(fileName);
			string oldPath = ExistingPath(fi.FullName, table);
			string newPath = fi.DirectoryName
									.Replace(Properties.Settings.Default.KPath, Properties.Settings.Default.KMapped)
									.Replace(Properties.Settings.Default.MetalPath, Properties.Settings.Default.SMapped) + @"\";

			if (oldPath != string.Empty && oldPath != newPath) {
				string message = string.Format("Original path was '{0}'\n\nUpdate to '{1}\\'?", oldPath, fi.DirectoryName
									.Replace(Properties.Settings.Default.KPath, Properties.Settings.Default.KMapped)
									.Replace(Properties.Settings.Default.MetalPath, Properties.Settings.Default.SMapped));
				swMessageBoxResult_e mbRes = (swMessageBoxResult_e)swApp.SendMsgToUser2(message, (int)swMessageBoxIcon_e.swMbQuestion, (int)swMessageBoxBtn_e.swMbOkCancel);
				if (mbRes == swMessageBoxResult_e.swMbHitOk) {
					DoInsert(fileName, table);
				}
			} else {
				DoInsert(fileName, table);
			}
		}

		private int GetKeyCol(string pathlessFilename) {
			SqlConnection connection = new SqlConnection(Properties.Settings.Default.connectionString);
			int alertKeyCol = 0;
			try {
				connection.Open();
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException(@"The connection is already open, or information is missing from the connection string: '" +
						Properties.Settings.Default.connectionString + "'.", ioe);
			} catch (SqlException se) {
				throw new ExportPDFException(@"A connection-level error occurred while opening the connection, whatever that means.", se);
			} catch (Exception e) {
				throw new ExportPDFException(@"Whoops. There's a problem", e);
			}

			try {
				SqlCommand comm = new SqlCommand("SELECT FileID " +
					string.Format("FROM {0} ", Properties.Settings.Default.table) +
					"WHERE FName COLLATE Latin1_General_CI_AI LIKE @fname " +
					"ORDER BY DateCreated DESC", connection);
				comm.Parameters.AddWithValue("@fname", pathlessFilename);
				SqlDataReader dr = comm.ExecuteReader(System.Data.CommandBehavior.SingleResult);
				if (dr.Read()) {
					alertKeyCol = dr.GetInt32(0);
				}
			} catch (Exception e) {
				throw e;
			} finally {
				connection.Close();
			}

			return alertKeyCol;
		}

		private string ExistingPath(string path, string table) {
			string res = string.Empty;
			FileInfo fi = new FileInfo(path);
			SqlConnection connection = new SqlConnection(Properties.Settings.Default.connectionString);

			try {
				connection.Open();
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException(@"The connection is already open, or information is missing from the connection string: '" +
						Properties.Settings.Default.connectionString + "'.", ioe);
			} catch (SqlException se) {
				throw new ExportPDFException(@"A connection-level error occurred while opening the connection, whatever that means.", se);
			} catch (Exception e) {
				throw new ExportPDFException(@"Whoops. There's a problem", e);
			}

			try {
				string sql = string.Format("SELECT * FROM {0} WHERE FName = @fname", table);
				SqlCommand command = new SqlCommand(sql, connection);
				command = new SqlCommand(sql, connection);
				command.Parameters.AddWithValue("@fname", new FileInfo(path).Name);
				using (SqlDataReader dr = command.ExecuteReader()) {
					if (dr.Read()) {
						res = dr.GetString(2);
					}
				}

			} catch (InvalidCastException ice) {
				throw new ExportPDFException("Invalid cast exception.", ice);
			} catch (SqlException se) {
				throw new ExportPDFException(String.Format("Couldn't execute query: {0}", se));
			} catch (IOException ie) {
				throw new ExportPDFException("IO exception", ie);
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException("The SqlConnection closed or dropped during a streaming operation.", ioe);
			} catch (Exception e) {
				throw new ExportPDFException("I looked for the connection, but it was gone.", e);
			} finally {
				connection.Close();
			}
			return res;
		}

		private bool UncheckedAlertExists(string path) {
			bool res = false;
			FileInfo fi = new FileInfo(path);
			SqlConnection connection = new SqlConnection(Properties.Settings.Default.connectionString);

			try {
				connection.Open();
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException(@"The connection is already open, or information is missing from the connection string: '" +
						Properties.Settings.Default.connectionString + "'.", ioe);
			} catch (SqlException se) {
				throw new ExportPDFException(@"A connection-level error occurred while opening the connection, whatever that means.", se);
			} catch (Exception e) {
				throw new ExportPDFException(@"Whoops. There's a problem", e);
			}

			try {
				string sql = "SELECT * FROM dbo.GEN_ALERTS WHERE ALERTDESC = @fname AND ALERTCHK = 0";
				SqlCommand command = new SqlCommand(sql, connection);
				command = new SqlCommand(sql, connection);
				command.Parameters.AddWithValue("@fname", new FileInfo(path).Name);
				using (SqlDataReader dr = command.ExecuteReader()) {
					if (dr.Read()) {
						res = true;
					}
				}

			} catch (InvalidCastException ice) {
				throw new ExportPDFException("Invalid cast exception.", ice);
			} catch (SqlException se) {
				throw new ExportPDFException(String.Format("Couldn't execute query: {0}", se));
			} catch (IOException ie) {
				throw new ExportPDFException("IO exception", ie);
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException("The SqlConnection closed or dropped during a streaming operation.", ioe);
			} catch (Exception e) {
				throw new ExportPDFException("I looked for the connection, but it was gone.", e);
			} finally {
				connection.Close();
			}
			return res;
		}

		private void AlertSomeone(string path) {
			FileInfo fi = new FileInfo(path);
			SqlConnection connection = new SqlConnection(Properties.Settings.Default.connectionString);
			int insRes;
			bool uncheckedExists = UncheckedAlertExists(path);

			if (drwKey == 0) return;

			try {
				connection.Open();
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException(@"The connection is already open, or information is missing from the connection string: '" +
						Properties.Settings.Default.connectionString + "'.", ioe);
			} catch (SqlException se) {
				throw new ExportPDFException(@"A connection-level error occurred while opening the connection, whatever that means.", se);
			} catch (Exception e) {
				throw new ExportPDFException(@"Whoops. There's a problem", e);
			}

			try {
				string sql = string.Empty;
				SqlCommand command = new SqlCommand();
				if (!uncheckedExists) {
					sql = @"INSERT INTO dbo.GEN_ALERTS (ALERTTYPE, ALERTDESC, ALERTMSG, ALERTKEYCOL) VALUES (@atype, @adesc, @amsg, @akc);";
					command = new SqlCommand(sql, connection);
					command.Parameters.AddWithValue("@atype", 1);
					command.Parameters.AddWithValue("@adesc", fi.Name);
					command.Parameters.AddWithValue("@amsg", string.Format("Updated by {0}.", System.Environment.UserName));
					command.Parameters.AddWithValue("@akc", drwKey);
					insRes = command.ExecuteNonQuery();
				} else {
					sql = @"UPDATE dbo.GEN_ALERTS SET ALERTDATE = @aDate, ALERTMSG = @amsg WHERE ALERTDESC = @fname;";
					command = new SqlCommand(sql, connection);
					command.Parameters.AddWithValue("@aDate", DateTime.Now);
					command.Parameters.AddWithValue("@amsg", string.Format("Updated by {0}.", System.Environment.UserName));
					command.Parameters.AddWithValue("@fname", fi.Name);
					insRes = command.ExecuteNonQuery();
				}

			} catch (InvalidCastException ice) {
				throw new ExportPDFException("Invalid cast exception.", ice);
			} catch (SqlException se) {
				throw new ExportPDFException(String.Format("Couldn't execute query: {0}", se));
			} catch (IOException ie) {
				throw new ExportPDFException("IO exception", ie);
			} catch (InvalidOperationException ioe) {
				throw new ExportPDFException("The SqlConnection closed or dropped during a streaming operation.", ioe);
			} catch (Exception e) {
				throw new ExportPDFException("I looked for the connection, but it was gone.", e);
			} finally {
				connection.Close();
			}
		}

		private bool _metalDrawing = false;
		/// <summary>
		/// Is this a metal drawing? This will be figured out after a pdf is archived.
		/// </summary>
		public bool MetalDrawing {
			get { return _metalDrawing; }
			set { _metalDrawing = value; }
		}

	}

}
