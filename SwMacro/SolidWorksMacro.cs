using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;

using Thumbnail;

namespace ArchivePDF.csproj
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            string jsonPath = @"\\AMSTORE-SVR-22\cad\Solid Works\Amstore_Macros\ArchivePDF.json";

            if (System.IO.File.Exists(@"C:\Optimize\ArchivePDF.json"))
                jsonPath = @"C:\Optimize\ArchivePDF.json";
            
            string json = string.Empty;

            if (System.IO.File.Exists(jsonPath))
            {
                try
                {
                    json = System.IO.File.ReadAllText(jsonPath);
                }
                catch (Exception e)
                {
                    ErrMsg em = new ErrMsg(e);
                    em.ShowDialog();
                }

                try
                {
                    APathSet = Newtonsoft.Json.JsonConvert.DeserializeObject<PathSet>(json);
                }
                catch (Exception e)
                {
                    ErrMsg em = new ErrMsg(e);
                    em.ShowDialog();
                }
            }

            // Saving first.
            if (APathSet.SaveFirst)
            {
                try
                {
                    Frame swFrams = (Frame)swApp.Frame();
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                    Int32 errors = 0;
                    Int32 warnings = 0;
                    swFrams.SetStatusBarText("Saving ... this is usually the most time consuming part ...");
                    swModel.Save3((Int32)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                }
                catch (ExportPDFException e)
                {
                    //PDFArchiver._sendErrMessage(e, swApp);
                    ErrMsg em = new ErrMsg(e);
                    em.ShowDialog();
                }
                catch (Exception ex)
                {
                    //PDFArchiver._sendErrMessage(ex, swApp);
                    ErrMsg em = new ErrMsg(ex);
                    em.ShowDialog();
                }
            }

            if (APathSet.ExportPDF)
                if (this.ExportPdfs())
                    this.ExportThumbnail();

        }

        private bool ExportPdfs()
        {
            bool res = false;
            // Checking gauge, and exporting.
            try
            {
                GaugeSetter gs = new GaugeSetter(swApp);
                gs.CheckAndUpdateGaugeNotes();
                gs.Close();

                PDFArchiver pda = new PDFArchiver(ref swApp, APathSet);
                res = (pda.ExportPdfs());
                
                if (APathSet.ExportEDrw)
                    res = res && pda.ExportEDrawings();
            }
            catch (GaugeSetterException gEx)
            {
                //PDFArchiver._sendErrMessage(gEx, swApp);
                ErrMsg em = new ErrMsg(gEx);
                em.ShowDialog();
            }
            catch (ExportPDFException e)
            {
                //PDFArchiver._sendErrMessage(e, swApp);
                ErrMsg em = new ErrMsg(e);
                em.ShowDialog();
            }
            catch (Exception ex)
            {
                //PDFArchiver._sendErrMessage(ex, swApp);

                ErrMsg em = new ErrMsg(ex);
                em.ShowDialog();
            }
            return res;
        }

        private void ExportThumbnail()
        {
            if (APathSet.ExportImg)
            {
                try
                {
                    Thumbnailer thN = new Thumbnailer(swApp, APathSet);
                    thN.CreateThumbnail();
                    thN.SaveAsJPG(APathSet.JPGPath);
                    thN.CloseThumbnail();
                }
                catch (ThumbnailerException thEx)
                {
                    ErrMsg em = new ErrMsg(thEx);
                    em.ShowDialog();
                }
                catch (Exception ex)
                {
                    ErrMsg em = new ErrMsg(ex);
                    em.ShowDialog();
                }
            }
        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;

        private PathSet _ps;

        private PathSet APathSet
        {
            get { return _ps; }
            set { _ps = value; }
        }
	
    }
}


