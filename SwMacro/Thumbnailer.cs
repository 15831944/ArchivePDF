using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

//using SolidworksBMPWrapper;
//using ImageMagickObject;

namespace Thumbnail
{
    class Thumbnailer
    {
        private Frame swFrame;
        private DrawingDoc swDraw;
        private View swView;
        private String sourcePath = String.Empty;
        private String drawingPath = String.Empty;

        public Thumbnailer(SldWorks sw)
        {
            this.swApp = sw;
            this.APathSet.ShtFmtPath = @"\\AMSTORE-SVR-22\cad\Solid Works\AMSTORE_SHEET_FORMATS\zPostCard.slddrt";
            swFrame = (Frame)swApp.Frame();
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            if (swApp == null)
                throw new ThumbnailerException("I know you gave me the SW Application Object, but I dropped it somewhere.");

            if (swDraw == null)
                throw new ThumbnailerException("You must having a drawing document open.");

            //This should find the first page with something on it.
            String[] shtNames = (String[])swDraw.GetSheetNames();
            Int32 x = 0;
            while ((swView == null) && (x <= swDraw.GetSheetCount()))
            {
                try
                {
                    swDraw.ActivateSheet(shtNames[x]);
                }
                catch (ArgumentOutOfRangeException e)
                {
                    throw new ThumbnailerException("Went beyond the number of sheets.", e);
                }

                swView = (View)swDraw.GetFirstView();
                swView = (View)swView.GetNextView();
                x++;
            }

            swFrame.SetStatusBarText("Found " + swView.Name);

            if (swView == null)
                throw new ThumbnailerException("I couldn't find a view.");

            sourcePath = swView.GetReferencedModelName().ToUpper().Trim();
            drawingPath = swModel.GetPathName().ToUpper().Trim();
        }

        public Thumbnailer(SldWorks sw, ArchivePDF.csproj.PathSet ps)
        {
            this.swApp = sw;
            this.APathSet = ps;

            swFrame = (Frame)swApp.Frame();
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            if (swApp == null)
                throw new ThumbnailerException("I know you gave me the SW Application Object, but I dropped it somewhere.");

            if (swDraw == null)
                throw new ThumbnailerException("You must having a drawing document open.");

            //This should find the first page with something on it.
            String[] shtNames = (String[])swDraw.GetSheetNames();
            Int32 x = 0;
            while ((swView == null) && (x <= swDraw.GetSheetCount()))
            {
                try
                {
                    swDraw.ActivateSheet(shtNames[x]);
                }
                catch (ArgumentOutOfRangeException e)
                {
                    throw new ThumbnailerException("Went beyond the number of sheets.", e);
                }

                swView = (View)swDraw.GetFirstView();
                swView = (View)swView.GetNextView();
                x++;
            }

            swFrame.SetStatusBarText("Found " + swView.Name);

            if (swView == null)
                throw new ThumbnailerException("I couldn't find a view.");

            sourcePath = swView.GetReferencedModelName().ToUpper().Trim();
            drawingPath = swModel.GetPathName().ToUpper().Trim();
        }

        public void CreateThumbnail()
        {
            bool bRet;
            double xSize = 2 * .0254;
            double ySize = 2 * .0254;
            double xCenter = (xSize / 2) - (.3 * .0254);
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
            
            view = swDraw.CreateDrawViewFromModelView3(Path.GetFileNameWithoutExtension(sourcePath), "*Isometric", xCenter, yCenter, 0);
            bRet = view.SetDisplayMode3(false, (int)dispMode, false, true);
            System.Diagnostics.Debug.Print(view.ScaleDecimal.ToString());
            view.ScaleDecimal = view.ScaleDecimal * 2 ;
            System.Diagnostics.Debug.Print(view.ScaleDecimal.ToString());
            swDraw.ActivateView(view.Name);

            // For some reason, the model imports with a bunch of ugly pointless stuff displayed. So, here's where we turn them off.
            bRet = swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayOrigins, false);
            bRet = swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayPlanes, false);
            bRet = swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayRoutePoints, false);
            bRet = swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplaySketches, false);
        }

        public void SaveAsJPG(string targetPath)
        {
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

        public void SaveAsBMP(string targetPath)
        {
            bool bRet;
            StringBuilder tp = new StringBuilder(targetPath);
            StringBuilder targetFilename = new StringBuilder();
            
            if (!targetPath.EndsWith("\\"))
                tp.Append("\\");

            targetFilename.AppendFormat("{0}{1}.bmp", tp.ToString(), Path.GetFileNameWithoutExtension(drawingPath));
            string tempFileName = "" + targetFilename.ToString();

            System.Diagnostics.Debug.Print("Saving " + tempFileName);
            swFrame.SetStatusBarText("Saving '" + tempFileName + "'...");
            bRet = swModel.SaveBMP(tempFileName.ToString(), 96 * 2, 96 *2);
        }

        public void CloseThumbnail()
        {
            swFrame.SetStatusBarText("Closing '" + swModel.GetPathName().ToUpper() + "'...");
            System.Diagnostics.Debug.Print("Closing '" + swModel.GetPathName().ToUpper() + "'...");
            swApp.CloseDoc(swModel.GetPathName().ToUpper());
        }

        public void CloseFiles()
        {
            swFrame.SetStatusBarText("Closing everything...");
            System.Diagnostics.Debug.Print("Closing everything...");
            swApp.CloseAllDocuments(true);
        }

        private SldWorks swApp;

        public SldWorks swApplication
        {
            get { return swApp; }
            set { swApp = value; }
        }

        private ModelDoc2 swModel;

        public ModelDoc2 swDocument
        {
            get { return swModel; }
            set { swModel = value; }
        }

        private ArchivePDF.csproj.PathSet _ps;

        public ArchivePDF.csproj.PathSet APathSet
        {
            get { return _ps; }
            set { _ps = value; }
        }
	
	
    }
}
