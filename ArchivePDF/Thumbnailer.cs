using System;
using System.IO;
using System.Text;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using ArchivePDF.csproj.Properties;
//using SolidworksBMPWrapper;
//using ImageMagickObject;

namespace Thumbnail {
	class Thumbnailer {
		private Frame swFrame;
		private DrawingDoc swDraw;
		private View swView;
		private String sourcePath = String.Empty;
		private String drawingPath = String.Empty;
		public bool assmbly = true;

		public Thumbnailer(SldWorks sw) {
			swApp = sw;
			APathSet.ShtFmtPath = @"\\AMSTORE-SVR-22\cad\Solid Works\AMSTORE_SHEET_FORMATS\zPostCard.slddrt";
			swFrame = (Frame)swApp.Frame();
			swModel = (ModelDoc2)swApp.ActiveDoc;
			swDraw = (DrawingDoc)swModel;

			if (swApp == null)
				throw new ThumbnailerException("I know you gave me the SW Application Object, but I dropped it somewhere.");

			if (swDraw == null)
				throw new ThumbnailerException("You must having a drawing document open.");

			swView = GetFirstView(swApp);

			sourcePath = swView.GetReferencedModelName().ToUpper().Trim();

			if (!sourcePath.Contains("SLDASM"))
				assmbly = false;

			drawingPath = swModel.GetPathName().ToUpper().Trim();
		}

		public Thumbnailer(SldWorks sw, ArchivePDF.csproj.PathSet ps) {
			swApp = sw;
			APathSet = ps;

			swFrame = (Frame)swApp.Frame();
			swModel = (ModelDoc2)swApp.ActiveDoc;
			swDraw = (DrawingDoc)swModel;

			if (swApp == null)
				throw new ThumbnailerException("I know you gave me the SW Application Object, but I dropped it somewhere.");

			if (swDraw == null)
				throw new ThumbnailerException("You must having a drawing document open.");

			swView = GetFirstView(swApp);

			sourcePath = swView.GetReferencedModelName().ToUpper().Trim();

			if (!sourcePath.Contains("SLDASM"))
				assmbly = false;

			drawingPath = swModel.GetPathName().ToUpper().Trim();
		}

		public void CreateThumbnail() {
			if (assmbly) {
				bool bRet;
				double xSize = 2 * .0254;
				double ySize = 2 * .0254;
				double xCenter = (xSize / 2) - Settings.Default.WeirdArbitraryFactor;
				double yCenter = (ySize / 2);

				swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPapersUserDefined;
				swDwgTemplates_e drwgTemplate = swDwgTemplates_e.swDwgTemplateNone;
				swDisplayMode_e dispMode = swDisplayMode_e.swFACETED_HIDDEN;

				//G:\\Solid Works\\AMSTORE_SHEET_FORMATS\\AM_PART.slddrt
				//swModel = (ModelDoc2)swApp.NewDocument(@"\\AMSTORE-SVR-22\cad\Solid Works\AMSTORE_SHEET_FORMATS\zPostCard.slddrt", (int)paperSize, xSize, ySize);
				swModel = (ModelDoc2)swApp.NewDocument(APathSet.ShtFmtPath, (int)paperSize, xSize, ySize);
				swDraw = (DrawingDoc)swModel;
				bRet = swDraw.SetupSheet5("AMS1", (int)paperSize, (int)drwgTemplate, 1, 1, false, "", xSize, ySize, "Default", false);

				View view = swDraw.DropDrawingViewFromPalette(76, xSize, ySize, 0);
				swDraw.ActivateView("Drawing View1");
				bRet = swModel.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", xSize, ySize, 0, false, 0, null, 0);

				view = swDraw.CreateDrawViewFromModelView3(sourcePath, "*Isometric", xCenter, yCenter, 0);
				bRet = view.SetDisplayMode3(false, (int)dispMode, false, true);
				System.Diagnostics.Debug.Print(view.ScaleDecimal.ToString());

				ScaleAppropriately(swDraw, view, 0.8, 2);

				System.Diagnostics.Debug.Print(view.ScaleDecimal.ToString());
				swDraw.ActivateView(view.Name);

				TurnOffSillyMarks();
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="formats">An array of paths to SolidWorks sheet templates.</param>
		/// <param name="monochrome">A bool array indicating whether we ought to have ImageMagick convert to 2-bit.</param>
		public void CreateThumbnails(string[] formats, bool[] monochrome) {
			bool bRet;
			double xSize = 2 * .0254;
			double ySize = 2 * .0254;
			double xCenter = (xSize / 2);
			double yCenter = (ySize / 2);
			string orientation = "*Isometric";
			string template = string.Empty;

			swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPapersUserDefined;
			swDwgTemplates_e drwgTemplate = swDwgTemplates_e.swDwgTemplateCustom;
			swDisplayMode_e dispMode = swDisplayMode_e.swHIDDEN;

			for (int i = 0; i < formats.Length; i++) {
				template = (new FileInfo(formats[i]).Name).Split(new char[] { '.' })[0];
				switch (formats[i].ToUpper().Contains("POSTCARD")) {
					case true:
						xSize = 7 * 0.0254;
						ySize = 5 * 0.0254;

						xCenter = (xSize / 2);
						yCenter = (ySize / 2);
						orientation = "*Isometric";
						break;
					case false:
						xSize = 2 * 0.0254;
						ySize = 2 * 0.0254;

						xCenter = (xSize / 2) - Settings.Default.WeirdArbitraryFactor;
						yCenter = (ySize / 2);
						orientation = "*Trimetric";
						break;
					default:
						break;
				}
				swModel = (ModelDoc2)swApp.NewDocument(formats[i], (int)paperSize, xSize, ySize);
				swDraw = (DrawingDoc)swModel;
				bRet = swDraw.SetupSheet5("Sheet1", (int)paperSize, (int)drwgTemplate, 1, 1, false, template, xSize, ySize, "Default", false);

				View view = swDraw.DropDrawingViewFromPalette(76, xSize, ySize, 0);
				swDraw.ActivateView("Drawing View1");
				bRet = swModel.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", xSize, ySize, 0, false, 0, null, 0);
				view = swDraw.CreateDrawViewFromModelView3(sourcePath, orientation, xCenter, yCenter, 0);
				bRet = view.SetDisplayMode3(false, (int)dispMode, false, true);

				//view.ScaleDecimal = GetAppropriateScalingFactor(swDraw, view, 0.8, 2);
				ScaleAppropriately(swDraw, view, 0.9, 2);

				swDraw.ActivateView(view.Name);

				TurnOffSillyMarks();

				switch (monochrome[i]) {
					case true:
						SaveAsBMP(Settings.Default.BMPPath);
						break;
					case false:
						SaveAsJPG(Settings.Default.JPGPath);
						break;
					default:
						break;
				}
				CloseThumbnail();
			}
		}

		public void ScaleAppropriately(DrawingDoc dd, View vv, double resize_factor, double starting_factor) {
			double[] sht_size = GetSheetSize(dd);
			double new_factor = starting_factor;

			vv.ScaleDecimal = starting_factor;
			double[] view_size = GetBoundingBoxSize(vv);

			while (view_size[0] > sht_size[0] || view_size[1] > sht_size[1]) {
				vv.ScaleDecimal *= resize_factor;
				view_size = GetBoundingBoxSize(vv);
			}
		}

		public double GetAppropriateScalingFactor(DrawingDoc dd, View vv, double resize_factor, double starting_factor) {
			double[] sht_size = GetSheetSize(dd);
			double new_factor = starting_factor;
			vv.ScaleDecimal = starting_factor;
			double[] view_size = GetBoundingBoxSize(vv);
			double[] tmp_view_size = view_size;

			do {
				tmp_view_size[0] *= resize_factor;
				tmp_view_size[1] *= resize_factor;
				new_factor *= resize_factor;
			} while (tmp_view_size[0] > sht_size[0] || tmp_view_size[1] > sht_size[1]);

			return new_factor;
		}

		public double[] GetSheetSize(DrawingDoc d) {
			Sheet sht = (Sheet)d.GetCurrentSheet();
			double[] sp = (double[])sht.GetProperties();
			double[] s_size = { sp[5], sp[6] };
			return s_size;
		}

		public double[] GetBoundingBoxSize(View v) {
			double[] bb = (double[])v.GetOutline();
			double[] bb_size = { (bb[2] - bb[0]), (bb[3] - bb[1]) };
			return bb_size;
		}

		public double[] GetPosition(View v) {
			return (double[])swView.Position;
		}

		public void TurnOffSillyMarks() {
			// For some reason, the model imports with a bunch of ugly pointless stuff displayed. So, here's where we turn them off.
			bool bRet = swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayOrigins, false);
			bRet &= swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayPlanes, false);
			bRet &= swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayRoutePoints, false);
			bRet &= swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplaySketches, false);
		}

		public View GetFirstView(SldWorks sw) {
			ModelDoc2 swModel = (ModelDoc2)sw.ActiveDoc;
			View v;
			DrawingDoc d = (DrawingDoc)swModel;
			string[] shtNames = (String[])swDraw.GetSheetNames();
			string message = string.Empty;

			//This should find the first page with something on it.
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

		public void SaveAsJPG(string targetPath) {
			int saveVersion = (int)swSaveAsVersion_e.swSaveAsCurrentVersion;
			int saveOptions = (int)swSaveAsOptions_e.swSaveAsOptions_Silent;
			int refErrors = 0;
			int refWarnings = 0;
			bool bRet;
			StringBuilder tp = new StringBuilder(targetPath);
			StringBuilder targetFileName = new StringBuilder();

			if (!targetPath.EndsWith("\\"))
				tp.Append("\\");

			targetFileName.AppendFormat("{0}{1}.jpg", tp.ToString(), Path.GetFileNameWithoutExtension(drawingPath));
			string tempFileName = "" + targetFileName.ToString();

			System.Diagnostics.Debug.Print("Saving " + tempFileName);
			swFrame.SetStatusBarText("Saving '" + tempFileName + "'...");

			bRet = swModel.SaveAs4(tempFileName, saveVersion, saveOptions, ref refErrors, ref refWarnings);
		}

		public void SaveAsBMP(string targetPath) {
			bool bRet;
			string tmpPath = Path.GetTempPath();
			StringBuilder tp = new StringBuilder(targetPath);
			StringBuilder targetFilename = new StringBuilder();

			if (!targetPath.EndsWith("\\"))
				tp.Append("\\");

			targetFilename.AppendFormat("{0}{1}.bmp", tmpPath, Path.GetFileNameWithoutExtension(drawingPath));
			string tempFileName = "" + targetFilename.ToString();

			System.Diagnostics.Debug.Print("Saving " + tempFileName);
			swFrame.SetStatusBarText("Saving '" + tempFileName + "'...");

			// Maybe this affects BMP. The "Export Options" screen says it'll affect TIF/PSD/JPG/PNG.
			swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, 600);
			bRet = swModel.SaveBMP(tempFileName.ToString(), 96 * 2, 96 * 2);

			string imPath = WheresImageMagick();
			if (imPath.Length > 0) {
				System.Diagnostics.Process p = new System.Diagnostics.Process();
				p.EnableRaisingEvents = false;
				p.StartInfo.FileName = string.Format(@"{0}\convert.exe ", imPath);
				// I'll just resize. ImageMagick can't mess this up, right?
				p.StartInfo.Arguments = string.Format("\"{0}\" -resize 192x192! -colors 2 \"{1}{2}.bmp\"",
					targetFilename.ToString(),
					tp.ToString(),
					Path.GetFileNameWithoutExtension(drawingPath));
				p.Start();
				p.WaitForExit();
			} else {
				swApp.SendMsgToUser2(ArchivePDF.csproj.Properties.Resources.WheresImageMagick,
					(int)swMessageBoxIcon_e.swMbWarning,
					(int)swMessageBoxBtn_e.swMbOk);
			}
		}

		public string WheresImageMagick() {
			Microsoft.Win32.RegistryKey rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\ImageMagick\Current\");
			if (rk != null) {
				return (string)rk.GetValue("BinPath".ToUpper());
			} else {
				return string.Empty;
			}
		}

		public void CloseThumbnail() {
			//if (assmbly) {
			swFrame.SetStatusBarText("Closing '" + swModel.GetPathName().ToUpper() + "'...");
			System.Diagnostics.Debug.Print("Closing '" + swModel.GetPathName().ToUpper() + "'...");
			swApp.CloseDoc(swModel.GetPathName().ToUpper());
			//}
		}

		public void CloseFiles() {
			if (assmbly) {
				swFrame.SetStatusBarText("Closing everything...");
				System.Diagnostics.Debug.Print("Closing everything...");
				swApp.CloseAllDocuments(true);
			}
		}

		private SldWorks swApp;

		public SldWorks swApplication {
			get { return swApp; }
			set { swApp = value; }
		}

		private ModelDoc2 swModel;

		public ModelDoc2 swDocument {
			get { return swModel; }
			set { swModel = value; }
		}

		private ArchivePDF.csproj.PathSet _ps;

		public ArchivePDF.csproj.PathSet APathSet {
			get { return _ps; }
			set { _ps = value; }
		}


	}
}
