using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ArchivePDF.csproj {
	[Serializable()]
	[XmlRoot("GaugeCollection")]
	public class GaugeCollection {
		private Gauge[] _gauge;

		[XmlArray("Gauges")]
		[XmlArrayItem("Material", typeof(Gauge))]
		public Gauge[] Gauge {
			get { return _gauge; }
			set { _gauge = value; }
		}

	}
}
