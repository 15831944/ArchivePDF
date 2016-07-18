using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ArchivePDF.csproj {
  [Serializable()]
  public class Gauge {
    private string _gauge;

    [XmlElement("Gauge")]
    public string GaugeNumber {
      get { return _gauge; }
      set { _gauge = value; }
    }

    private string _thickness;

    [XmlElement("Thickness")]
    public string Thickness {
      get { return _thickness; }
      set { _thickness = value; }
    }
  }
}
