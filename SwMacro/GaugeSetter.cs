using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArchivePDF.csproj
{
    class GaugeSetter
    {
        private Excel.Application xlApp;
        private Excel.Worksheet xlSht;
        private Excel.Workbook xlWkBk;
        private Excel.Sheets xlShts;

        private SldWorks swApp;
        private Frame swFrame;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private View swView;
        private DisplayDimension swDispDim;
        private Dimension swDim;
        //private Annotation swAnn;

        private double ourGauge;
        private bool gaugeNotFound = true;
        
        public GaugeSetter(SldWorks sw)
        {
            swApp = sw;
            //String fileName = "\\\\AMSTORE-SVR-22\\Solid Works\\SolidWorks Data\\lang\\english\\gaugesetter.xls";
            String fileName = @"\\AMSTORE-SVR-22\cad\Solid Works\SolidWorks Data\lang\english\ARCHIVE\gaugesetter.xls";
            
            _setupSW();
            _loadXLFile(fileName);
        }

        public GaugeSetter(SldWorks sw, PathSet ps)
        {
            swApp = sw;
            this.APathSet = ps;
            string fileName = this.APathSet.GaugePath;
            
            _setupSW();
            _loadXLFile(fileName);
        }
        
        private void _loadXLFile(String wkBkPath)
        {
            swFrame.SetStatusBarText(String.Format("Loading {0}...", wkBkPath));
            xlApp = new Excel.Application();
            xlApp.Visible = false;

            Excel.XlUpdateLinks updateLinks = Microsoft.Office.Interop.Excel.XlUpdateLinks.xlUpdateLinksAlways;
            Boolean readOnly = true;
            Excel.XlFileFormat format = Excel.XlFileFormat.xlWK1;
            String passwd = String.Empty;
            String wtResPswd = String.Empty;
            Boolean ignoreReadOnly = false;
            Excel.XlPlatform platform = Microsoft.Office.Interop.Excel.XlPlatform.xlWindows;
            String delimiter = String.Empty;
            Boolean editable = false;
            Boolean notify = false;
            Int32 converter = 0;
            Boolean addToMru = false;
            Boolean local = false;
            Boolean corruptLoad = false;

            if (System.IO.File.Exists(wkBkPath))
            {
                xlWkBk = xlApp.Workbooks.Open(
                            wkBkPath,
                            updateLinks,
                            readOnly,
                            format,
                            passwd,
                            wtResPswd,
                            ignoreReadOnly,
                            platform,
                            delimiter,
                            editable,
                            notify,
                            converter,
                            addToMru,
                            local,
                            corruptLoad);
            }
            else
            {
                GaugeSetterException gEx = new GaugeSetterException(String.Format("Error opening '{0}'.", wkBkPath));
                //gEx.Data.Add("updateLinks", updateLinks);
                //gEx.Data.Add("readOnly", readOnly);
                //gEx.Data.Add("format", format);
                //gEx.Data.Add("passwd", passwd);
                //gEx.Data.Add("wtResPaswd", wtResPswd);
                //gEx.Data.Add("ignoreReadOnly", ignoreReadOnly);
                //gEx.Data.Add("platform", platform);
                //gEx.Data.Add("delimiter", delimiter);
                //gEx.Data.Add("editable", editable);
                //gEx.Data.Add("notify", notify);
                //gEx.Data.Add("converter", converter);
                //gEx.Data.Add("addToMru", addToMru);
                //gEx.Data.Add("local", local);
                //gEx.Data.Add("corruptLoad", corruptLoad);
                throw gEx;
            }

            xlShts = xlWkBk.Worksheets;
            xlSht = (Excel.Worksheet)xlShts.get_Item("Sheet1");
        }

        private void _setupSW()
        {
            swFrame = (Frame)swApp.Frame();
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            if (swDraw == null)
            {
                throw new GaugeSetterException("You must have a Drawing Document open.");
            }
        }
        /// <summary>
        /// I'd rather not do this.
        /// </summary>
        private void _killEmbeddedExcel()
        {
            System.Diagnostics.Debug.Print("_killEmbeddedExcel()");
            System.Management.WqlObjectQuery wqlQ =
                new System.Management.WqlObjectQuery("Select * from Win32_Process");
            System.Management.ManagementObjectSearcher searcher =
                new System.Management.ManagementObjectSearcher(wqlQ);



            foreach (System.Management.ManagementObject mo in searcher.Get())
            {
                if (mo["Description"].ToString() == "EXCEL.EXE" && mo["CommandLine"].ToString().Contains("-Embedding"))
                {
                    string spID = mo["ProcessID"].ToString();
                    int ipID = 0;
                    int.TryParse(spID, out ipID);

                    System.Diagnostics.Debug.Print("mo: {0}, {1}, {2}", mo["ProcessID"], mo["Description"], mo["CommandLine"]);
                    System.Diagnostics.Process msA = System.Diagnostics.Process.GetProcessById(ipID);

                    if (ipID > 0)
                        msA.Kill();
                }
            }
        }

        private void _releaseObject(Object o)
        {
            System.Diagnostics.Debug.Print("Releasing {0}.", o.ToString());
            try
            {
                Int32 count = 0;
                do
                {
                    count = System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
                    System.Diagnostics.Debug.Print(":{0}:", count);
                } while (count > 0);
                
                o = null;
            }
            catch (Exception e)
            {
                o = null;
                System.Diagnostics.Debug.Print("{0}", e.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        public void Close()
        {
            xlApp.Quit();
            _releaseObject(xlSht);
            _releaseObject(xlShts);
            _releaseObject(xlApp);
            _releaseObject(xlWkBk);
            //_killEmbeddedExcel();
        }

        public void CheckAndUpdateGaugeNotes()
        {
            String currPage = swDraw.GetCurrentSheet().ToString();
            Int32 shtCount = swDraw.GetSheetCount();
            String[] shtName = (String[])swDraw.GetSheetNames();

            for (int page = 0; page < shtCount; page++)
            {
                swFrame.SetStatusBarText(String.Format("Activating page {0}...", shtName[page]));
                swDraw.ActivateSheet(shtName[page]);
                swView = (View)swDraw.GetFirstView();

                while (swView != null)
                {
                    swDispDim = swView.GetFirstDisplayDimension5();

                    while (swDispDim != null)
                    {
                        swDim = (Dimension)swDispDim.GetDimension2(0);
                        swFrame.SetStatusBarText(String.Format("Processing '{0}' => '{1}'...", swDim.Name, swDim.Value.ToString()));

                        //System.Diagnostics.Debug.Print("::{0}:", swDim.FullName);

                        if (swDispDim.GetText((Int32)swDimensionTextParts_e.swDimensionTextCalloutBelow).EndsWith("GA"))
                        {
                            //System.Diagnostics.Debug.Print("GA");
                            Double og;
                            if (!Double.TryParse(swDim.GetSystemValue2("").ToString(), out og))
                            {
                                throw new GaugeSetterException("Couldn't parse dimension value.");
                            }
                            
                            ourGauge = og / 0.0254;

                            //System.Diagnostics.Debug.Print("::{0}::{1}", ourGauge, og);

                            for (int i = 8; i <= 31; i++)
                            {
                                Double dCellVal;
                                //String desiredCell = String.Format("B{0}", i);
                                Excel.Range rng = (Excel.Range)xlSht.Cells[i, 2];
                                //System.Diagnostics.Debug.Print("{0}: {1}", desiredCell, rng.Value2);

                                if (!Double.TryParse(rng.Value2.ToString(), out dCellVal))
                                {
                                    throw new GaugeSetterException("Couldn't parse gauge thickness.");
                                }
                                else
                                
                                if (Math.Abs(ourGauge - dCellVal) < 0.00003)
                                {
                                    String gaugeCell = String.Empty;
                                    String gaugeString = String.Empty;
                                    gaugeNotFound = false;

                                    if (swDispDim.GetText((Int32)swDimensionTextParts_e.swDimensionTextCalloutBelow).Contains("("))
                                    {
                                        Excel.Range gRng = (Excel.Range)xlSht.Cells[i, 1];
                                        gaugeCell = gRng.Value2.ToString();

                                        //System.Diagnostics.Debug.Print(gaugeCell);
                                        gaugeString = String.Format("({0} GA)", gaugeCell.Split('\u0020')[0]);
                                    }
                                    else
                                    {
                                        Excel.Range gRng = (Excel.Range)xlSht.Cells[i, 1];
                                        gaugeCell = gRng.Value2.ToString();

                                        //System.Diagnostics.Debug.Print(gaugeCell);
                                        gaugeString = String.Format("{0} GA", gaugeCell.Split('\u0020')[0]);
                                    }
                                    swDispDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutBelow, gaugeString);
                                }
                            }
                            if (gaugeNotFound)
                            {
                                if (!this.APathSet.SilenceGaugeErrors)
                                {
                                    StringBuilder sb = new StringBuilder("Nonstandard gauge thickness detected:\n");
                                    sb.AppendFormat("{0} in {1} = {2}", swDim.Name, swView.Name, ourGauge);
                                    swApp.SendMsgToUser2(sb.ToString()
                                        , (int)swMessageBoxIcon_e.swMbWarning
                                        , (int)swMessageBoxBtn_e.swMbOk);
                                }
                                gaugeNotFound = false;
                            }
                        }
                        swDispDim = swDispDim.GetNext5();
                    }
                    swView = (View)swView.GetNextView();
                }
            }
            swDraw.ActivateSheet(currPage);
            swFrame.SetStatusBarText("Rebuilding");
            swDraw.ForceRebuild();
            swFrame.SetStatusBarText(String.Empty);
        }

        private PathSet _pathSet;

        public PathSet APathSet
        {
            get { return _pathSet; }
            set { _pathSet = value; }
        }

    }
}
