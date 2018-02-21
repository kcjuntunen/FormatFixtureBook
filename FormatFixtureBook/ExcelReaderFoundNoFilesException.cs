using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormatFixtureBook {
	[Serializable]
	public class ExcelReaderFoundNoFilesException : Exception {
		public ExcelReaderFoundNoFilesException() { }
		public ExcelReaderFoundNoFilesException(string message) : base(message) { }
		public ExcelReaderFoundNoFilesException(string message, Exception inner) : base(message, inner) { }
		protected ExcelReaderFoundNoFilesException(
		System.Runtime.Serialization.SerializationInfo info,
		System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
	}
}
