using System;
using System.Collections.Generic;
using System.Text;

namespace ArchivePDF.csproj
{
    class PathSet
    {
        private string _gaugePath;

        public string GaugePath
        {
            get { return _gaugePath; }
            set { _gaugePath = value; }
        }

        private string _shtFmtPath;

        public string ShtFmtPath
        {
            get { return _shtFmtPath; }
            set { _shtFmtPath = value; }
        }

        private string _jpgPath;

        public string JPGPath
        {
            get { return _jpgPath; }
            set { _jpgPath = value; }
        }

        private string _kPath;

        public string KPath
        {
            get { return _kPath; }
            set { _kPath = value; }
        }

        private string _gPath;

        public string GPath
        {
            get { return _gPath; }
            set { _gPath = value; }
        }

        private string _metalPath;

        public string MetalPath
        {
            get { return _metalPath; }
            set { _metalPath = value; }
        }
	
	
        //"GaugePath" : "\\\\AMSTORE-SVR-22\\cad\\Solid Works\\SolidWorks Data\\lang\\english\\ARCHIVE\\gaugesetter.xls"
        //"ShtFmtPath" : "\\\\AMSTORE-SVR-22\\cad\\Solid Works\\AMSTORE_SHEET_FORMATS\\zPostCard.slddrt"
        //"JPGPath" : "\\\\AMSTORE-SVR-01\\details\\JPGs\\"
        //"KPath" : "\\\\AMSTORE-SVR-01\\details\\"
        //"GPath" : "\\\\AMSTORE-SVR-22\\cad\\PDF ARCHIVE\\"
    }
}
