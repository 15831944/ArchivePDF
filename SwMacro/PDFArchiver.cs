using System;
using System.Collections.Generic;
using System.Text;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ArchivePDF.csproj
{
    class PDFArchiver
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Frame swFrame;
        private DrawingDoc swDraw;
        private View swView;
        private CustomPropertyManager swCustPropMgr;
        private swDocumentTypes_e modelType = swDocumentTypes_e.swDocNONE;
        private String sourcePath = String.Empty;
        private String drawingPath = String.Empty;
        //private Boolean shouldCheck = true;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sw">Requires a <see cref="SolidWorks.Interop.sldworks.SldWorks"/> type.</param>
        /// <param name="ps">Requires a <see cref="ArchivePDF.csproj.PathSet"/></param>
        public PDFArchiver(ref SldWorks sw, ArchivePDF.csproj.PathSet ps)
        {
            swApp = sw;
            this.APathSet = ps;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            if (DocCheck())
            {
                swDraw = (DrawingDoc)swModel;

                if (swDraw == null)
                {
                    throw new ExportPDFException("You must have a Drawing Document open.");
                }


                String[] shtNames = (String[])swDraw.GetSheetNames();

                swFrame = (Frame)swApp.Frame();

                //This should find the first page with something on it.
                Int32 x = 0;
                do
                {
                    try
                    {
                        swDraw.ActivateSheet(shtNames[x]);
                    }
                    catch (IndexOutOfRangeException e)
                    {
                        throw new ExportPDFException("Went beyond the number of sheets.", e);
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }

                    swView = (View)swDraw.GetFirstView();
                    swView = (View)swView.GetNextView();
                    x++;
                } while ((swView == null) && (x < swDraw.GetSheetCount()));

                if (swView == null)
                {
                    throw new ExportPDFException("I couldn't find a model anywhere in this document.");
                }

                sourcePath = swView.GetReferencedModelName().ToUpper().Trim();
                drawingPath = swModel.GetPathName().ToUpper().Trim();

                if (sourcePath.Contains("SLDASM"))
                {
                    modelType = swDocumentTypes_e.swDocASSEMBLY;

                }
                else if (sourcePath.Contains("SLDPRT"))
                {
                    modelType = swDocumentTypes_e.swDocPART;
                }
                else
                {
                    modelType = swDocumentTypes_e.swDocNONE;
                }
            }
            else
            {
                ExportPDFException e = new ExportPDFException(String.Format("The drawing has to be saved, {0}.", System.Environment.UserName));
                throw e;
            }
        }

        /// <summary>
        /// Exports current drawing to PDFs.
        /// </summary>
        /// <returns>Returns a <see cref="System.Boolean"/>, should you need one.</returns>
        public Boolean ExportPdfs()
        {
            //System.Diagnostics.Debug.Print("---------PDF---------");

            String pdfSourceName = System.IO.Path.GetFileNameWithoutExtension(drawingPath);
            String pdfAltPath = System.IO.Path.GetDirectoryName(drawingPath).Substring(3);

            String fFormat = String.Empty;
            ModelDocExtension swModExt = swModel.Extension;
            String Rev = GetRev(swModExt);
            List<String> pdfTarget = new List<string>();

            pdfTarget.Add(String.Format("{0}{1}\\{2}.PDF", this.APathSet.KPath , pdfAltPath, pdfSourceName));
            pdfTarget.Add(String.Format("{0}{1}\\{2}{3}.PDF", this.APathSet.GPath, pdfAltPath, pdfSourceName, Rev));
            Boolean success = this.SaveFiles(pdfTarget);
            return success;
        }

        public Boolean ExportMetalPdfs()
        {
            System.Diagnostics.Debug.Print("---------MetalPDF---------");

            String pdfSourceName = System.IO.Path.GetFileNameWithoutExtension(drawingPath);
            String pdfAltPath = System.IO.Path.GetDirectoryName(drawingPath).Substring(45);
            //String pdfAltPath = ".\\";

            String fFormat = String.Empty;
            ModelDocExtension swModExt = swModel.Extension;
            List<String> pdfTarget = new List<string>();

            pdfTarget.Add(String.Format("{0}{1}\\{2}.PDF", this.APathSet.MetalPath, pdfAltPath, pdfSourceName));
            Boolean success = this.SaveFiles(pdfTarget);
            return success;
        }
        /// <summary>
        /// Opens and saves models referenced in the open drawing.
        /// </summary>
        /// <returns>Returns a <see cref="System.Boolean"/>, should you need one.</returns>
        public Boolean ExportEDrawings()
        {
            //System.Diagnostics.Debug.Print("---------EDrawing---------");
            String sourceName = System.IO.Path.GetFileNameWithoutExtension(sourcePath);
            String altPath = System.IO.Path.GetDirectoryName(sourcePath).Substring(3);
            //System.Diagnostics.Debug.Print("sourcePath = {0}, sourceName = {1}, altPath = {2}", sourcePath, sourceName, altPath);
            ModelDocExtension swModExt = swModel.Extension;
            String Rev = GetRev(swModExt);
            String fFormat = String.Empty;
            List<String> target = new List<string>();
            Boolean measurable = true;
            Int32 options = 0;
            Int32 errors = 0;

            switch (modelType)
            {
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

            if (swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure))
            {
                measurable = false;
            }
            else
            {
                measurable = true;
            }

            target.Add(String.Format("{0}{1}\\{2}{3}", this.APathSet.KPath, altPath, sourceName, fFormat));
            target.Add(String.Format("{0}{1}\\{2}{3}{4}", this.APathSet.GPath, altPath, sourceName, Rev, fFormat));


            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, true);
            Boolean success = this.SaveFiles(target);

            if (!measurable)
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, false);

            swFrame.SetStatusBarText("Closing model.");
            swApp.CloseDoc(sourcePath);
            
            return success;
        }
        /// <summary>
        /// This thing does the saving.
        /// </summary>
        /// <param name="fl">A <see cref="System.Collections.Generic.List<>"/> of destination paths.</param>
        /// <returns>Returns a <see cref="System.Boolean"/>, indicating success.</returns>
        private Boolean SaveFiles(List<String> fl)
        {
            Int32 tries = 9;
            Int32 saveVersion = (Int32)swSaveAsVersion_e.swSaveAsCurrentVersion;
            Int32 saveOptions = (Int32)swSaveAsOptions_e.swSaveAsOptions_Silent;
            Int32 refErrors = 0;
            Int32 refWarnings = 0;
            Boolean success = true;

            foreach (String fileName in fl)
            {
                //System.Diagnostics.Debug.Print(">>{0}<<", fileName);
                if (drawingPath != String.Empty)
                {
                    swFrame.SetStatusBarText(String.Format("Checking path: '{0}'", fileName));
                    if (!CreatePath(fileName))
                    {
                        ExportPDFException e = new ExportPDFException("Unable to save file, folder could not be created.");
                        //e.Data.Add("who", System.Environment.UserName);
                        //e.Data.Add("when", DateTime.Now);
                        throw e;
                    }

                    swFrame.SetStatusBarText(String.Format("Exporting '{0}'", fileName));

                    if (System.IO.File.Exists(fileName))
                        System.IO.File.Delete(fileName);

                    do
                    {
                        success = swModel.SaveAs4(fileName, saveVersion, saveOptions, ref refErrors, ref refWarnings);
                        swFrame.SetStatusBarText(String.Format("Exporting '{0}': try {1}", fileName, --tries));
                    } while (!System.IO.File.Exists(fileName) && tries > 0);

                    if (!System.IO.File.Exists(fileName))
                        success = false;

                    if (success)
                    {
                        swFrame.SetStatusBarText(String.Format("Exported '{0}'", fileName));
                    }
                    else
                    {
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
        private String GetRev(ModelDocExtension swModExt)
        {
            swCustPropMgr = (CustomPropertyManager)swModExt.get_CustomPropertyManager("");
            String[] propNames = (String[])swCustPropMgr.GetNames();
            String ValOut = String.Empty;
            String ResolvedValOut = String.Empty;
            Boolean WasResolved = false;

            String result = String.Empty;
            
            foreach (String name in propNames)
            {
                swCustPropMgr.Get5(name, false, out ValOut, out ResolvedValOut, out WasResolved);
                if (name.Contains("REVISION") && ValOut != String.Empty)
                    result = "-" + ValOut;
            }

            if (result.Length != 3)
            {
                ExportPDFException e = new ExportPDFException("Check to make sure drawing is at least revision AA or later.");
                //e.Data.Add("who", System.Environment.UserName);
                //e.Data.Add("when", DateTime.Now);
                //e.Data.Add("result", result);
                throw e;
            }

            //System.Diagnostics.Debug.Print(result);
            return result;
        }
        private Boolean DocCheck()
        {
            if (swModel == null)
            {
                ExportPDFException e = new ExportPDFException("You must have a drawing document open.");
                //e.Data.Add("who", System.Environment.UserName);
                //e.Data.Add("when", DateTime.Now);
                return false;
            }
            else if ((Int32)swModel.GetType() != (Int32)swDocumentTypes_e.swDocDRAWING)
            {
                ExportPDFException e = new ExportPDFException("You must have a drawing document open.");
                //e.Data.Add("who", System.Environment.UserName);
                //e.Data.Add("when", DateTime.Now);
                return false;
            }
            else if (swModel.GetPathName() == String.Empty)
            {
                ExportPDFException e = new ExportPDFException("You must first save drawing before attempting to archive PDF.");
                //e.Data.Add("who", System.Environment.UserName);
                //e.Data.Add("when", DateTime.Now);
                return false;
            }
            else
                return true;
        }
        private Boolean CreatePath(String pathCheck)
        {
            pathCheck = System.IO.Path.GetDirectoryName(pathCheck);
            String targetpath = String.Empty;

            if (!System.IO.Directory.Exists(pathCheck))
            {
                String msg = String.Format("'{0}' does not exist.  Do you wish to create this folder?", pathCheck);
                System.Windows.Forms.MessageBoxButtons mbb = System.Windows.Forms.MessageBoxButtons.YesNo;
                System.Windows.Forms.MessageBoxIcon mbi = System.Windows.Forms.MessageBoxIcon.Question;
                System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show(msg, "New folder", mbb, mbi);
                if (dr == System.Windows.Forms.DialogResult.Yes)
                {
                    System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(pathCheck);
                    return di.Exists;
                }
                else
                    return false;
            }
            return true;
        }

        private ArchivePDF.csproj.PathSet _ps;

        public ArchivePDF.csproj.PathSet APathSet
        {
            get { return _ps; }
            set { _ps = value; }
        }
	

    public static void _sendErrMessage(Exception e, SldWorks sw)
        {
            String msg = String.Empty;
            msg = String.Format("{0} threw error:\n{1}", e.TargetSite, e.Message);
            msg += String.Format("\n\nStack trace:\n{0}", e.StackTrace);

            if (e.Data.Count > 0)
            {
                msg += "\n\nData:\n";

                foreach (System.Collections.DictionaryEntry de in e.Data)
                {
                    msg += String.Format("{0}: {1}\n", de.Key, de.Value);
                }
            }

            sw.SendMsgToUser2(msg, (Int32)swMessageBoxIcon_e.swMbInformation, (Int32)swMessageBoxBtn_e.swMbOk);
        }
    }

}
