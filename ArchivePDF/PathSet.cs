using System;

namespace ArchivePDF.csproj {
	public class PathSet {
		public PathSet() {
			Initialated = false;
		}

		public bool Initialated { get; set; }

		private string _gaugePath;

		public string GaugePath {
			get { return _gaugePath; }
			set { _gaugePath = value; }
		}

		private string _gaugeRegex;

		public string GaugeRegex {
			get { return _gaugeRegex; }
			set { _gaugeRegex = value; }
		}

		private string _shtFmtPath;

		public string ShtFmtPath {
			get { return _shtFmtPath; }
			set { _shtFmtPath = value; }
		}

		private string _jpgPath;

		public string JPGPath {
			get { return _jpgPath; }
			set { _jpgPath = value; }
		}

		private string _kPath;

		public string KPath {
			get { return _kPath; }
			set { _kPath = value; }
		}

		private string _gPath;

		public string GPath {
			get { return _gPath; }
			set { _gPath = value; }
		}

		private string _metalPath;

		public string MetalPath {
			get { return _metalPath; }
			set { _metalPath = value; }
		}

		private bool _saveFirst;

		public bool SaveFirst {
			get { return _saveFirst; }
			set { _saveFirst = value; }
		}

		private bool _silentGaugeErr;

		public bool SilenceGaugeErrors {
			get { return _silentGaugeErr; }
			set { _silentGaugeErr = value; }
		}

		private bool _exportPdf;

		public bool ExportPDF {
			get { return _exportPdf; }
			set { _exportPdf = value; }
		}

		private bool _exportEDrw;

		public bool ExportEDrw {
			get { return _exportEDrw; }
			set { _exportEDrw = value; }
		}

		private bool _exportImg;

		public bool ExportImg {
			get { return _exportImg; }
			set { _exportImg = value; }
		}

		private bool _writeToDb;

		public bool WriteToDb {
			get { return _writeToDb; }
			set { _writeToDb = value; }
		}

		public bool ExportSTEP { get; set; }
	}
}
