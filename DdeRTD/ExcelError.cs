using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace DdeRTD
{
	public static class ExcelError
	{
		public static readonly ErrorWrapper Div0 = new ErrorWrapper(-2146826281);
		public static readonly ErrorWrapper NA = new ErrorWrapper(-2146826246);
		public static readonly ErrorWrapper Name = new ErrorWrapper(-2146826259);
		public static readonly ErrorWrapper Null = new ErrorWrapper(-2146826288);
		public static readonly ErrorWrapper Num = new ErrorWrapper(-2146826252);
		public static readonly ErrorWrapper Ref = new ErrorWrapper(-2146826265);
		public static readonly ErrorWrapper Value = new ErrorWrapper(-2146826273);
	};
}
