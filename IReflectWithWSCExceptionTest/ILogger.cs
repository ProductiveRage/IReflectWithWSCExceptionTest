using System;

namespace IReflectWithWSCExceptionTest
{
	public interface ILogger
	{
		void LogIgnoringAnyError(LogLevel logLevel, Func<string> contentGenerator);
	}
}
