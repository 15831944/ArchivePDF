using System;
using System.Collections.Generic;
using System.Text;

namespace ArchivePDF.csproj
{
    [global::System.Serializable]
    public class ExportPDFException : Exception
    {
        //
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
        //

        public ExportPDFException() { }
        public ExportPDFException(string message) : base(message) { }
        public ExportPDFException(string message, Exception inner) : base(message, inner) { }
        protected ExportPDFException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}
