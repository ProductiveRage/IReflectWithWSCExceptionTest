using System;
using System.IO;
using Microsoft.VisualBasic;

namespace IReflectWithWSCExceptionTest
{
	class Program
	{
		static void Main(string[] args)
		{
			// To try to imitate the issue in the Engine as closely as possible, the following is done:
			//  1. A WSC is created (this is like a Control on Page)
			//  2. This WSC gets a Response reference and its Go method is called, both over IDispatch
			//  3. Inside the WSC, it will call Response.CreateObject to get another WSC - this one is wrapped in a FromFileWSCWrapper (so it is like a Booking Component)
			//  4. That WSC gets its Response reference set by VBScript in the first WSC and then "Go" is called on the second WSC, which calls Response.Redirect, which throws a RedirectException
			//  5. The error is caught in the IDispatchAccess layer, which was used when the first component called the second's "Go" method (since that invocation went through IReflect and so
			//     through the FromFileWSCWrapper and called into the second component over IDispatch). The error has all of the message content, so it may be succesfully translated back into
			//     a RedirectException.
			//  6. If we call the first component using IDispatchAccess then the error is caught again there - the error code is correct (it may be mapped back onto a RedirectException) but its
			//     bstrDescription is null and so we can not retrieve the URL to redirect to (using COMSurvivableException.RethrowAsOriginalIfPossible will result in a RedirectException being
			//     raised with a generic message "Exception of type 'System.Runtime.InteropServices.COMException' was thrown" which is not good
			//     - However, if we call the first component's "Go" method using dynamic then we DO get the full RedirectException being throw, which is very good
			// Note: If the second component is not wrapped in a FromFileWSCWrapper then we get the full RedirectException whether we use IDispatch or dynamic
			var callOuterComponentMethodsUsingDynamic = false;

			var wscFile = new FileInfo("TestComponent.wsc");
			var wsc = Interaction.GetObject("script:" + wscFile.FullName, null);

			try
			{
				var response = new WrappedResponse(new Response());
				IDispatchAccess.SetProperty(wsc, "Response", response); // Have to use IDispatchAccess here since dynamic sets the property using PutDispProperty instead of PutRefDispProperty (.net bug, I think)
				if (callOuterComponentMethodsUsingDynamic)
					((dynamic)wsc).Go();
				else
					IDispatchAccess.CallMethod(wsc, "Go");
			}
			catch (RedirectException e)
			{
				Console.WriteLine("Redirect: " + e.Message);
				if (e.Message.StartsWith("Exception of type"))
					Console.WriteLine("^ ********** This is a disappointment");
				else
					Console.WriteLine("^ ********** THIS IS WHAT WE WANT!");
			}
			catch (Exception e)
			{
				Console.WriteLine("FAIL: " + e.Message + " (" + e.GetType().Name + ")");
				Console.WriteLine("^ ********** This is a disappointment");
			}
			Console.ReadLine();
		}
	}
}
