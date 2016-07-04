﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace IReflectWithWSCExceptionTest
{
	public static class IDispatchAccess
	{
		[DllImport(@"oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
		static extern Int32 VariantClear(IntPtr pvarg);

		private const int LOCALE_SYSTEM_DEFAULT = 2048;
		private const int DISPID_PROPERTYPUT = -3;
		private const int SizeOfNativeVariant = 16;
		private static readonly ComTypes.DISPPARAMS EmptyDISPPARAMS = new ComTypes.DISPPARAMS()
		{
			cArgs = 0,
			cNamedArgs = 0,
			rgdispidNamedArgs = IntPtr.Zero,
			rgvarg = IntPtr.Zero
		};

		public enum InvokeFlags : ushort
		{
			DISPATCH_METHOD = 1,
			DISPATCH_PROPERTYGET = 2,
			DISPATCH_PROPERTYPUT = 4,
			DISPATCH_PROPERTYPUTREF = 8
		}

		public static bool ImplementsIDispatch(object source)
		{
			if (source == null)
				throw new ArgumentNullException("source");

			return source is IDispatch;
		}

		public static T CallMethod<T>(object source, string name, params object[] args)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");
			if (args == null)
				throw new ArgumentNullException("args");

			return Invoke<T>(source, InvokeFlags.DISPATCH_METHOD, GetDispId(source, name), args);
		}

		public static object CallMethod(object source, string name, params object[] args)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");
			if (args == null)
				throw new ArgumentNullException("args");

			return CallMethod<object>(source, name, args);
		}

		public static T CallDefaultMethod<T>(object source, params object[] args)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (args == null)
				throw new ArgumentNullException("args");

			return Invoke<T>(source, InvokeFlags.DISPATCH_METHOD, 0, args);
		}

		public static T GetProperty<T>(object source, string name)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			return Invoke<T>(source, InvokeFlags.DISPATCH_PROPERTYGET, GetDispId(source, name));
		}

		public static object GetProperty(object source, string name)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			return GetProperty<object>(source, name);
		}

		public static T GetProperty<T>(object source, string name, params object[] indices)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			return Invoke<T>(
					source,
					InvokeFlags.DISPATCH_PROPERTYGET,
					GetDispId(source, name),
					(indices ?? new object[0])
			);
		}

		public static object GetProperty(object source, string name, params object[] indices)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			return GetProperty<object>(source, name, indices);
		}

		public static T GetDefaultProperty<T>(object source, params object[] args)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (args == null)
				throw new ArgumentNullException("args");

			return Invoke<T>(source, InvokeFlags.DISPATCH_PROPERTYGET, 0, args);
		}

		public static void SetProperty(object source, string name, object value, params object[] indices)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			Invoke<object>(
				source,
				IsVBScriptValueType(value) ? IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUT : IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUTREF,
				GetDispId(source, name),
				(indices ?? new object[0]).Concat(new[] { value }).ToArray()
			);
		}

		private static bool IsVBScriptValueType(object o)
		{
			return ((o == null) || (o == DBNull.Value) || (o is ValueType) || (o is string) || (o.GetType().IsArray));
		}

		public static bool DeclaresMember(object source, string name)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			var dispIdDetails = GetDispIdDetails(source, name);
			return (dispIdDetails.HrRet == 0);
		}

		private static int GetDispId(object source, string name)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			var dispIdDetails = GetDispIdDetails(source, name);
			if (dispIdDetails.HrRet == 0)
				return dispIdDetails.DispId;

			var message = "Invalid member \"" + name + "\"";
			var errorType = GetErrorMessageForHResult(dispIdDetails.HrRet);
			if (errorType != CommonErrors.Unknown)
				message += " [" + errorType.ToString() + "]";
			throw new ArgumentException(message);
		}

		private static DispIdDetails GetDispIdDetails(object source, string name)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException("Null/blank name specified");

			var IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
			var rgDispId = new int[1] { 0 }; // This will be populated with a the DispId of the named member (if available)
			var hrRet = ((IDispatch)source).GetIDsOfNames
			(
					ref IID_NULL,
					new string[1] { name },
					1, // number of names to get ids for
					LOCALE_SYSTEM_DEFAULT,
					rgDispId
			);
			return new DispIdDetails
			{
				HrRet = hrRet,
				DispId = (hrRet == 0) ? rgDispId[0] : 0
			};
		}

		private struct DispIdDetails
		{
			/// <summary>
			/// This will be zero for a successfully-retrieved DispId (non-zero indicates an error)
			/// </summary>
			public int HrRet;
			public int DispId;
		}

		private static T Invoke<T>(object source, InvokeFlags invokeFlags, int dispId, params object[] args)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (!Enum.IsDefined(typeof(InvokeFlags), invokeFlags))
				throw new ArgumentOutOfRangeException("invokeFlags");
			if (args == null)
				throw new ArgumentNullException("args");

			var memoryAllocationsToFree = new List<IntPtr>();
			IntPtr rgdispidNamedArgs;
			int cNamedArgs;
			if ((invokeFlags == InvokeFlags.DISPATCH_PROPERTYPUT) || (invokeFlags == InvokeFlags.DISPATCH_PROPERTYPUTREF))
			{
				// There must be at least one argument specified; only one if it is a non-indexed property and multiple if there are index values as well as the
				// value to set to
				if (args.Length < 1)
					throw new ArgumentException("At least one argument must be specified when DISPATCH_PROPERTYPUT is requested");

				var pdPutID = Marshal.AllocCoTaskMem(sizeof(Int64));
				Marshal.WriteInt64(pdPutID, DISPID_PROPERTYPUT);
				memoryAllocationsToFree.Add(pdPutID);

				rgdispidNamedArgs = pdPutID;
				cNamedArgs = 1;
			}
			else
			{
				rgdispidNamedArgs = IntPtr.Zero;
				cNamedArgs = 0;
			}

			var variantsToClear = new List<IntPtr>();
			IntPtr rgvarg;
			if (args.Length == 0)
				rgvarg = IntPtr.Zero;
			else
			{
				// We need to allocate enough memory to store a variant for each argument (and then populate this memory)
				rgvarg = Marshal.AllocCoTaskMem(SizeOfNativeVariant * args.Length);
				memoryAllocationsToFree.Add(rgvarg);
				for (var index = 0; index < args.Length; index++)
				{
					// Note: The "IDispatch::Invoke method (Automation)" page (http://msdn.microsoft.com/en-us/library/windows/desktop/ms221479(v=vs.85).aspx)
					// states that "Arguments are stored in pDispParams->rgvarg in reverse order" so we'll reverse them here
					var arg = args[(args.Length - 1) - index];

					// According to http://stackoverflow.com/a/1866268 it seems like using ToInt64 here will be valid for both 32 and 64 bit machines. While
					// this may apparently not be the most performant approach, it should do the job.
					// Don't think we have to worry about pinning any references when we do this manipulation here since we are allocating the array in
					// unmanaged memory and so the garbage collector won't be moving anything around (and GetNativeVariantForObject copies the reference
					// and automatic pinning will prevent the GC from interfering while this is happening).
					var pVariant = new IntPtr(
							rgvarg.ToInt64() + (SizeOfNativeVariant * index)
					);
					Marshal.GetNativeVariantForObject(arg, pVariant);
					variantsToClear.Add(pVariant);
				}
			}

			var dispParams = new ComTypes.DISPPARAMS()
			{
				cArgs = args.Length,
				rgvarg = rgvarg,
				cNamedArgs = cNamedArgs,
				rgdispidNamedArgs = rgdispidNamedArgs
			};

			try
			{
				var IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
				UInt32 pArgErr = 0;
				object varResult;
				var excepInfo = new ComTypes.EXCEPINFO();
				var hrRet = ((IDispatch)source).Invoke
				(
						dispId,
						ref IID_NULL,
						LOCALE_SYSTEM_DEFAULT,
						(ushort)invokeFlags,
						ref dispParams,
						out varResult,
						ref excepInfo,
						out pArgErr
				);
				if (hrRet != 0)
				{
					if (excepInfo.bstrDescription == null)
						Console.WriteLine($"Exception thrown while accessing {TypeDescriptor.GetClassName(source)} has null bstrDescription");
					else
						Console.WriteLine($"Exception thrown while accessing {TypeDescriptor.GetClassName(source)} has bstrDescription: \"{excepInfo.bstrDescription}\"");

					// Try to translate the exception back into a COMSurvivableException - if this is not possible then fall through to the code below
					COMSurvivableException.RethrowAsOriginalIfPossible(
						new COMException(excepInfo.bstrDescription, excepInfo.scode)
					);
					var message = "Failing attempting to invoke method with DispId " + dispId + ": ";
					var errorType = GetErrorMessageForHResult(hrRet);
					if (errorType == CommonErrors.DISP_E_MEMBERNOTFOUND)
						message += "Member not found";
					else if (!string.IsNullOrWhiteSpace(excepInfo.bstrDescription))
						message += excepInfo.bstrDescription;
					else
					{
						message += "Unspecified error";
						if (errorType != CommonErrors.Unknown)
							message += " [" + errorType.ToString() + "]";
					}
					throw new ArgumentException(message);
				}
				return (T)varResult;
			}
			finally
			{
				foreach (var variantToClear in variantsToClear)
					VariantClear(variantToClear);

				foreach (var memoryAllocationToFree in memoryAllocationsToFree)
					Marshal.FreeCoTaskMem(memoryAllocationToFree);
			}
		}

		private static CommonErrors GetErrorMessageForHResult(int hrRet)
		{
			if (Enum.IsDefined(typeof(CommonErrors), hrRet))
				return (CommonErrors)hrRet;

			return CommonErrors.Unknown;
		}

		// http://blogs.msdn.com/b/eldar/archive/2007/04/03/a-lot-of-hresult-codes.aspx
		private enum CommonErrors
		{
			Unknown = 0,

			E_UNEXPECTED = -2147418113,
			E_NOTIMPL = -2147467263,
			E_OUTOFMEMORY = -2147024882,
			E_INVALIDARG = -2147024809,
			E_NOINTERFACE = -2147467262,
			E_POINTER = -2147467261,
			E_HANDLE = -2147024890,
			E_ABORT = -2147467260,
			E_FAIL = -2147467259,
			E_ACCESSDENIED = -2147024891,
			E_PENDING = -2147483638,

			DISP_E_UNKNOWNINTERFACE = -2147352575,
			DISP_E_MEMBERNOTFOUND = -2147352573,
			DISP_E_PARAMNOTFOUND = -2147352572,
			DISP_E_TYPEMISMATCH = -2147352571,
			DISP_E_UNKNOWNNAME = -2147352570,
			DISP_E_NONAMEDARGS = -2147352569,
			DISP_E_BADVARTYPE = -2147352568,
			DISP_E_EXCEPTION = -2147352567,
			DISP_E_OVERFLOW = -2147352566,
			DISP_E_BADINDEX = -2147352565,
			DISP_E_UNKNOWNLCID = -2147352564,
			DISP_E_ARRAYISLOCKED = -2147352563,
			DISP_E_BADPARAMCOUNT = -2147352562,
			DISP_E_PARAMNOTOPTIONAL = -2147352561,
			DISP_E_BADCALLEE = -2147352560,
			DISP_E_NOTACOLLECTION = -2147352559,
			DISP_E_DIVBYZERO = -2147352558,
			DISP_E_BUFFERTOOSMALL = -2147352557
		}

		[ComImport()]
		[Guid("00020400-0000-0000-C000-000000000046")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		private interface IDispatch
		{
			[PreserveSig]
			int GetTypeInfoCount(out int Count);

			[PreserveSig]
			int GetTypeInfo
			(
					[MarshalAs(UnmanagedType.U4)] int iTInfo,
					[MarshalAs(UnmanagedType.U4)] int lcid,
					out ComTypes.ITypeInfo typeInfo
			);

			[PreserveSig]
			int GetIDsOfNames
			(
					ref Guid riid,
					[MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)]
								string[] rgsNames,
					int cNames,
					int lcid,
					[MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
			);

			[PreserveSig]
			int Invoke
			(
					int dispIdMember,
					ref Guid riid,
					uint lcid,
					ushort wFlags,
					ref ComTypes.DISPPARAMS pDispParams,
					out object pVarResult,
					ref ComTypes.EXCEPINFO pExcepInfo,
					out UInt32 pArgErr
			);
		}
	}
}
