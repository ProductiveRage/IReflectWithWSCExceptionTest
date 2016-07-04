using System;

namespace IReflectWithWSCExceptionTest
{
	public sealed class ConsoleLogger : ILogger
	{
		public void LogIgnoringAnyError(LogLevel logLevel, Func<string> contentGenerator)
		{
			Console.WriteLine(contentGenerator());
		}
	}
}
