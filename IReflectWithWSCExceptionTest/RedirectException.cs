using System;
using System.Runtime.InteropServices;

namespace IReflectWithWSCExceptionTest
{
	[ComVisible(true)]
	public sealed class RedirectException : COMSurvivableException
	{
		public RedirectException(string message) : base(message, reviver) { }

		private static COMSurvivableException reviver(string message)
		{
			return new RedirectException(message);
		}

		protected override byte UniqueErrorCode { get { return 12; } }
	}
}
