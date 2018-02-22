using System;

namespace FormatFixtureBook {
	[Serializable]
	public class PDFCreatorFileOpenException : Exception {
		public PDFCreatorFileOpenException() { }
		public PDFCreatorFileOpenException(string message) : base(message) { }
		public PDFCreatorFileOpenException(string message, Exception inner) : base(message, inner) { }
		protected PDFCreatorFileOpenException(
		System.Runtime.Serialization.SerializationInfo info,
		System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
	}
}
