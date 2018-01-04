using System;
using System.Collections.Generic;
using System.Text;

namespace ArchivePDF {
	[Serializable]
	public class MustHaveRevException : Exception {
		public MustHaveRevException() { }
		public MustHaveRevException(string message) : base(message) { }
		public MustHaveRevException(string message, Exception inner) : base(message, inner) { }
		protected MustHaveRevException(
		System.Runtime.Serialization.SerializationInfo info,
		System.Runtime.Serialization.StreamingContext context)
			: base(info, context) { }
	}
}
