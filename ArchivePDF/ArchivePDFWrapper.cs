using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Thumbnail;

namespace ArchivePDF.csproj {
  public class ArchivePDFWrapper {
    public ArchivePDFWrapper(SldWorks sw) {
      swApp = sw;
    }

    public void Archive() {
      string jsonPath = Properties.Settings.Default.defaultJSON;

      if (System.IO.File.Exists(Properties.Settings.Default.localJSON))
        jsonPath = Properties.Settings.Default.localJSON;

      string json = string.Empty;
      Frame swFrams = (Frame)swApp.Frame();
      swFrams.SetStatusBarText("Starting PDF Archiver...");
      ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
      Int32 errors = 0;
      Int32 warnings = 0;
      swDocumentTypes_e docType;
      docType = (swDocumentTypes_e)swModel.GetType();

      if (docType != swDocumentTypes_e.swDocDRAWING) {
        swApp.SendMsgToUser2(string.Format("You need to have a drawing open. This is a '{0}'.", docType.ToString()),
            (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
      } else {

        if (System.IO.File.Exists(jsonPath)) {
          try {
            json = System.IO.File.ReadAllText(jsonPath);
          } catch (Exception e) {
            ErrMsg em = new ErrMsg(e);
            em.ShowDialog();
          }

          try {
            APathSet = Newtonsoft.Json.JsonConvert.DeserializeObject<PathSet>(json);
          } catch (Exception e) {
            ErrMsg em = new ErrMsg(e);
            em.ShowDialog();
          }
        }

        // Saving first.
        if (APathSet.SaveFirst) {

          try {
            swFrams.SetStatusBarText("Saving ... this is usually the most time consuming part ...");
            swModel.Save3((Int32)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
          } catch (ExportPDFException e) {
            //PDFArchiver._sendErrMessage(e, swApp);
            ErrMsg em = new ErrMsg(e);
            em.ShowDialog();
          } catch (Exception ex) {
            //PDFArchiver._sendErrMessage(ex, swApp);
            ErrMsg em = new ErrMsg(ex);
            em.ShowDialog();
          }
        }

        if (APathSet.ExportPDF) {
          if (ExportPdfs()) {
            if (!swModel.GetPathName().ToUpper().Contains("METAL MANUFACTURING")) {
              bool kohls = false;
              LayerMgr lm = (LayerMgr)swModel.GetLayerManager();
              Layer l = (Layer)lm.GetLayer("KOHLS");
              if (l != null) {
                if (l.Visible == true) {
                  kohls = true;
                }
              }
              switch (kohls) {
                case true:
                  ExportKohlsThumbnail();
                  break;
                case false:
                  ExportThumbnail();
                  break;
                default:
                  break;
              }
            }
            swFrams.SetStatusBarText(string.Format("Done exporting '{0}' to PDF Archive.", swModel.GetPathName().ToUpper()));
          } else {
            swApp.SendMsgToUser2("Failed to save PDF.",
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk);
          }
        }
        swFrams.SetStatusBarText(":-(");
      }

      swFrams.SetStatusBarText(":-)");

      // Solidworks won't usually exit this macro on some machines. This seems to help.
      System.GC.Collect(0, GCCollectionMode.Forced);
    }

    private bool ExportPdfs() {
      bool res = false;
      // Checking gauge, and exporting.
      try {
        PDFArchiver pda = new PDFArchiver(ref swApp, APathSet);
        GaugeSetter gs = new GaugeSetter(swApp, APathSet);
        gs.CheckAndUpdateGaugeNotes2();
        res = (pda.ExportPdfs());

        if (APathSet.ExportEDrw && !pda.MetalDrawing)
          res = res && pda.ExportEDrawings();
      } catch (MustHaveRevException mhre) {
        swApp.SendMsgToUser2(mhre.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
      } catch (MustSaveException mse) {
        swApp.SendMsgToUser2(mse.Message,
            (int)swMessageBoxIcon_e.swMbStop,
            (int)swMessageBoxBtn_e.swMbOk);
      } catch (GaugeSetterException gEx) {
        //PDFArchiver._sendErrMessage(gEx, swApp);
        ErrMsg em = new ErrMsg(gEx);
        em.ShowDialog();
      } catch (ExportPDFException e) {
        //PDFArchiver._sendErrMessage(e, swApp);
        ErrMsg em = new ErrMsg(e);
        em.ShowDialog();
      } catch (Exception ex) {
        //PDFArchiver._sendErrMessage(ex, swApp);

        ErrMsg em = new ErrMsg(ex);
        em.ShowDialog();
      }
      return res;
    }

    private void ExportKohlsThumbnail() {
      if (APathSet.ExportImg) {
        try {
          Thumbnailer thN = new Thumbnailer(swApp, APathSet);
          string[] sht_fmts = { Properties.Settings.Default.ShtFmtPath[0], Properties.Settings.Default.ShtFmtPath[1] };
          bool[] monochrome = { false, true };
          thN.CreateThumbnails(sht_fmts, monochrome);
        } catch (ThumbnailerException thEx) {
          ErrMsg em = new ErrMsg(thEx);
          em.ShowDialog();
        } catch (Exception ex) {
          ErrMsg em = new ErrMsg(ex);
          em.ShowDialog();
        }
      }
    }

    private void ExportThumbnail() {
      if (APathSet.ExportImg) {
        try {
          Thumbnailer thN = new Thumbnailer(swApp, APathSet);
          if (thN.assmbly) {
            thN.CreateThumbnail();
            thN.SaveAsJPG(APathSet.JPGPath);
            thN.CloseThumbnail();
          }
        } catch (ThumbnailerException thEx) {
          ErrMsg em = new ErrMsg(thEx);
          em.ShowDialog();
        } catch (Exception ex) {
          ErrMsg em = new ErrMsg(ex);
          em.ShowDialog();
        }
      }
    }

    private string GetSheetFormat(ModelDoc2 m) {
      DrawingDoc d = (DrawingDoc)m;
      string[] sht_names = (string[])d.GetSheetNames();
      Sheet s = d.get_Sheet(sht_names[0]);
      return s.GetTemplateName();
    }

    public SldWorks swApp;

    private PathSet _ps;

    private PathSet APathSet {
      get { return _ps; }
      set { _ps = value; }
    }
  }
}
