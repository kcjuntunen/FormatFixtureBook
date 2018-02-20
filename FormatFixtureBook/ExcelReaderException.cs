using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormatFixtureBook {
	[Serializable]
	public class ExcelReaderException : Exception {
		public ExcelReaderException() { }
		public ExcelReaderException(string message) : base(message) { }
		public ExcelReaderException(string message, Exception inner) : base(message, inner) { }
		protected ExcelReaderException(
		System.Runtime.Serialization.SerializationInfo info,
		System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
	}
}
