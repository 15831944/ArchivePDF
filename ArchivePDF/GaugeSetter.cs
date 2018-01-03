using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace ArchivePDF.csproj {
  class GaugeSetter {

    private SldWorks swApp;
    private Frame swFrame;
    private ModelDoc2 swModel;
    private DrawingDoc swDraw;
    private View swView;
    private DisplayDimension swDispDim;
    private Dimension swDim;
    private Regex gaugeRegex;

    private double ourGauge;
    private bool gaugeNotFound = true;
		private string gaugeNote = @"GA";

    public GaugeSetter(SldWorks sw) {
      swApp = sw;

      gaugeRegex = new Regex(@"^[0-9]{1,2}\ ?GA\.?");

      _setupSW();
      _loadXMLFile(Properties.Settings.Default.GaugePath);
    }

    public GaugeSetter(SldWorks sw, PathSet ps) {
      swApp = sw;
      APathSet = ps;
      string fileName = APathSet.GaugePath;

      if (APathSet.GaugeRegex == string.Empty || APathSet.GaugeRegex == null) {
        gaugeRegex = new Regex(@"^[0-9]{1,2}\ ?GA\.?");
      } else {
        gaugeRegex =  new Regex(APathSet.GaugeRegex);
      }

      _setupSW();
      _loadXMLFile(Properties.Settings.Default.GaugePath);
    }

    private void _loadXMLFile(String xmlPath) {
      using (StreamReader sr = new StreamReader(xmlPath)) {
        XmlSerializer ser = new XmlSerializer(typeof(GaugeCollection));
        _gauges = (GaugeCollection)ser.Deserialize(sr);
      }
    }

    private void _setupSW() {
      swFrame = (Frame)swApp.Frame();
      swModel = (ModelDoc2)swApp.ActiveDoc;
      swDraw = (DrawingDoc)swModel;

      if (swDraw == null) {
        throw new GaugeSetterException("You must have a Drawing Document open.");
      }
    }

    public void CheckAndUpdateGaugeNotes2() {
      String currPage = swDraw.GetCurrentSheet().ToString();
      Int32 shtCount = swDraw.GetSheetCount();
      String[] shtName = (String[])swDraw.GetSheetNames();

      for (int page = 0; page < shtCount; page++) {
        swFrame.SetStatusBarText(String.Format("Activating page {0}...", shtName[page]));
        swDraw.ActivateSheet(shtName[page]);
        swView = (View)swDraw.GetFirstView();

        while (swView != null) {
          swDispDim = swView.GetFirstDisplayDimension5();

          while (swDispDim != null) {
            swDim = (Dimension)swDispDim.GetDimension2(0);
            swFrame.SetStatusBarText(String.Format("Processing '{0}' => '{1}'...", swDim.Name, swDim.Value.ToString()));
            string dimtext = swDispDim.GetText((Int32)swDimensionTextParts_e.swDimensionTextCalloutBelow);
						dimtext = dimtext.Replace(@"(", string.Empty).Replace(@")", string.Empty);
            if (gaugeRegex.IsMatch(dimtext)) {
							string [] texts_ = dimtext.Split(' ');
							gaugeNote = texts_[texts_.Length - 1];
              Double og;
              if (!Double.TryParse(swDim.GetSystemValue2("").ToString(), out og)) {
                throw new GaugeSetterException("Couldn't parse dimension value.");
              }
              ourGauge = og / 0.0254;

              for (int i = 0; i < Gauges.Gauge.Length; i++) {
                Double dCellVal;

                if (!Double.TryParse(Gauges.Gauge[i].Thickness, out dCellVal)) {
                  throw new GaugeSetterException("Couldn't parse gauge thickness.");
                } else

                  if (Math.Abs(ourGauge - dCellVal) < 0.00003) {
                    String gaugeCell = String.Empty;
                    String gaugeString = String.Empty;
                    gaugeNotFound = false;

										swDispDim.ShowParenthesis = true;
                    gaugeString = String.Format("{0} {1}", Gauges.Gauge[i].GaugeNumber, gaugeNote.ToUpper());
                    swDispDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutBelow, gaugeString);
                  }
              }
              if (gaugeNotFound) {
                if (!APathSet.SilenceGaugeErrors) {
                  StringBuilder sb = new StringBuilder("Non-standard gauge thickness detected:\n");
									sb.AppendFormat(get_format_txt(), swDim.Name, swView.Name, ourGauge);
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

      if (gaugeNotFound) { // Why waste time rebuilding?
        swFrame.SetStatusBarText("Rebuilding");
        swDraw.ForceRebuild(); 
      }
      swFrame.SetStatusBarText(String.Empty);
    }

		private string get_format_txt() {
			int prec_ = swModel.Extension.GetUserPreferenceInteger(
				(int)swUserPreferenceIntegerValue_e.swDetailingLinearDimPrecision,
				(int)swUserPreferenceOption_e.swDetailingLinearDimension);
			string format_txt_ = @"{0} in {1} = {2:0.";
			for (int i = 0; i < prec_; i++) {
				format_txt_ += @"0";
			}
			format_txt_ += @"}";
			return format_txt_;
		}

    private GaugeCollection _gauges;

    public GaugeCollection Gauges {
      get { return _gauges; }
      set { _gauges = value; }
    }

    private PathSet _pathSet;

    public PathSet APathSet {
      get { return _pathSet; }
      set { _pathSet = value; }
    }

  }
}
