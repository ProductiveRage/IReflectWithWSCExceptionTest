using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace IReflectWithWSCExceptionTest
{
	/// <summary>
	/// When exceptions are raised in managed code, caught in unmanaged code (ie. WSCs) and then come back out of that unmanaged code it will be as a COMException - type
	/// meta data and state are lost. To workaround this, custom exceptions may be derived from this base class and then an additional try..catch wrapped around calls to
	/// the unmanaged code where the catch block calls the static ThrowAsMostSpecificPossible function on this class, restoring the original exception.
	/// </summary>
	[Serializable]
	public abstract class COMSurvivableException : COMException
	{
		private static readonly Dictionary<byte, Reviver> _revivers = new Dictionary<byte, Reviver>();
		protected COMSurvivableException(string messageWithAnyStateData, Reviver reviver) : base(messageWithAnyStateData)
		{
			if (string.IsNullOrWhiteSpace(messageWithAnyStateData))
				throw new ArgumentException("Null/blank messageWithAnyStateData specified");
			if (reviver == null)
				throw new ArgumentNullException("reviver");

			lock (_revivers)
			{
				_revivers[UniqueErrorCode] = reviver;
			}
			HResult = CustomErrorHResultGenerator.GetHResult(UniqueErrorCode);
		}

		/// <summary>
		/// Derived types are responsible for knowing how to map back to an instance of themselves from a message-with-state-data string associated with the UniqueErrorCode
		/// for the exception class. Since each derived type has a unique error code, it should not be possible for a string to be received that can not be deceiphered - so
		/// implementations of this delegate should never return null, an exception should be raised if the state data is invalid.
		/// </summary>
		protected delegate COMSurvivableException Reviver(string messageWithAnyStateData);

		/// <summary>
		/// It is important that this return a consistent value unique to the derived class and that this value be available before the constructor is executed
		/// </summary>
		protected abstract byte UniqueErrorCode { get; }

		protected COMSurvivableException(SerializationInfo info, StreamingContext context) : base(info, context) { }

		/// <summary>
		/// If this COMException was translated from a COMSurvivableException that has passed up through unmanaged code then this method will re-throw as the original exception
		/// type (which will also be derived from COMSurvivableException). If not then this will perform no action, so long as the specified exception is not null (which is an
		/// error condition and will result in an argument exception being thrown).
		/// </summary>
		[DebuggerStepThrough] // No benefit to the debugger breaking here, if an exception is raised then we want the debugger to stop where the translated exception is caught
		public static void RethrowAsOriginalIfPossible(COMException e)
		{
			if (e == null)
				throw new ArgumentNullException("e");

			var uniqueErrorCode = CustomErrorHResultGenerator.GetErrorCode(e.HResult);
			Reviver reviver;
			lock (_revivers)
			{
				if (!_revivers.TryGetValue(uniqueErrorCode, out reviver))
					return;
			}
			var revivedException = reviver(e.Message);
			if (revivedException != null)
				throw revivedException;
		}

		/// <summary>
		/// If all calls to unmanaged code are wrapped in a try..catch where the catch block passes any exception into this method then any exceptions that need reconstructing
		/// after passing through the unmanaged code will be restored (as long as they are derived from this class) any any other exceptions will be re-thrown - eg.
		/// 
		///   try { unmanagedCodeComponent.Go(); } catch (Exception e) { COMSurvivableException.ThrowAsMostSpecificPossible(e); }
		/// </summary>
		[DebuggerStepThrough] // No benefit to the debugger breaking here, if an exception is raised then we want the debugger to stop where the translated exception is caught
		public static void ThrowAsMostSpecificPossible(Exception e)
		{
			if (e == null)
				throw new ArgumentNullException("e");

			var baseComException = e.GetBaseException() as COMException;
			if (baseComException != null)
				RethrowAsOriginalIfPossible(baseComException);
			throw e;
		}

		public static bool IsComSurvivableHResult(int hresult)
		{
			return CustomErrorHResultGenerator.IsComSurvivableHResult(hresult);
		}

		private static class CustomErrorHResultGenerator
		{
			// See https://msdn.microsoft.com/en-us/library/cc231198.aspx
			//  Bit 0      Severity (1 = fail vs 0 = success)
			//  Bit 1      Reserved (set to 0 if NTSTATUS is 0)
			//  Bit 2      Customer (set to 1 for non-Microsoft status)
			//  Bit 3      NTSTATUS (set to 0 for non-Microsoft status)
			//  Bit 4      Reserved (set to 0)
			//  Bits 5-15  Facility (zero is default, most applicable for Customer codes)
			//  Bits 16-31 Code 
			private const int CUSTOMERROR_BASE = unchecked((int)0xA000B600); // Severity = 1, Customer = 1, Upper byte of Code = 182 (arbitrary for NM|t)
			public static int GetHResult(byte errorCode) { return CUSTOMERROR_BASE | errorCode; }

			private const int ERROR_CODE_MASK = byte.MaxValue;
			public static byte GetErrorCode(int hresult) { return (byte)(hresult & ERROR_CODE_MASK); }

			public static bool IsComSurvivableHResult(int hresult)
			{
				return (hresult >= CUSTOMERROR_BASE && hresult <= (CUSTOMERROR_BASE | ERROR_CODE_MASK));
			}
		}
	}
}