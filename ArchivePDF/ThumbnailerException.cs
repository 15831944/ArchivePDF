using System;
using System.Collections.Generic;
using System.Text;


namespace Thumbnail
{
    
    [global::System.Serializable]
    public class ThumbnailerException : Exception
    {
        //
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
        //

        public ThumbnailerException() { }
        public ThumbnailerException(string message) : base(message) { }
        public ThumbnailerException(string message, Exception inner) : base(message, inner) { }
        protected ThumbnailerException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }

        //public static void sendMessage(Exception e)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(String.Format("'{0}' threw error '{1}' in '{2}'\r\nStack trace:\n{3}\r\n"
        //        , e.Source
        //        , e.Message
        //        , e.TargetSite
        //        , e.StackTrace));

        //    foreach (System.Collections.DictionaryEntry item in e.Data)
        //        sb.AppendFormat("{0}: {1}\r\n", item.Key, item.Value.ToString());

        //    SolidworksBMPWrapper.ErrForm ef = new SolidworksBMPWrapper.ErrForm();
        //    if (e.InnerException != null)
        //        sb.AppendFormat("\r\n\r\n{0}\r\n\r\nInner Stack Trace:\r\n{1}", e.InnerException.Message, e.InnerException.StackTrace);

        //    ef.fillErrMsg(sb.ToString());
        //    ef.Text = String.Format("Error in {0}", e.Source);
        //    ef.ShowDialog();

        //    //System.Windows.Forms.MessageBox.Show(sb.ToString()
        //    //    , "Error"
        //    //    , System.Windows.Forms.MessageBoxButtons.OK
        //    //    , System.Windows.Forms.MessageBoxIcon.Error);
        //}

        //public static void sendMessage(string str)
        //{
        //    SolidworksBMPWrapper.ErrForm ef = new SolidworksBMPWrapper.ErrForm();
        //    ef.fillErrMsg(str);
        //    ef.Text = String.Format("Message for {0}", Environment.UserName) ;
        //    ef.ShowDialog();
        //    //System.Windows.Forms.MessageBox.Show(sb.ToString()
        //    //    , "Error"
        //    //    , System.Windows.Forms.MessageBoxButtons.OK
        //    //    , System.Windows.Forms.MessageBoxIcon.Error);
        //}
    }
}
