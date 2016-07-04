using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace IReflectWithWSCExceptionTest
{
	[ComVisible(true)]
	public sealed class Response
	{
		public object CreateObject(string progId)
		{
			if (progId == "ETWP.TestComponent2")
			{
				var wscFile = new FileInfo("TestComponent2.wsc");
				var wsc = Interaction.GetObject("script:" + wscFile.FullName, null);
				return new FromFileWSCWrapper(wsc, typeof(ITestComponent), new ConsoleLogger());
			}

			throw new ArgumentException("Unsupported progId: " + progId);
		}

		public void Write(string message)
		{
			Console.WriteLine(message);
		}

		public void Redirect(string url)
		{
			throw new RedirectException(url);
		}
	}
}
