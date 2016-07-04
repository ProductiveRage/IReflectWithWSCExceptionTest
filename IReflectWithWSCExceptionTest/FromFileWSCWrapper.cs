using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace IReflectWithWSCExceptionTest
{
	[ClassInterface(ClassInterfaceType.AutoDispatch)]
	[ComVisible(true)]
	public sealed class FromFileWSCWrapper : IReflect
	{
		private const string DispIdZeroIdentifier = "[DISPID=0]";

		private readonly object _target;
		private readonly Type _template;
		private readonly ILogger _logger;
		public FromFileWSCWrapper(object target, Type template, ILogger logger)
		{
			if (target == null)
				throw new ArgumentNullException("target");
			if (template == null)
				throw new ArgumentNullException("template");
			if (logger == null)
				throw new ArgumentNullException("logger");

			// 2016-06-30 DWR: We could consider caching the available properties and methods that the specified template has, rather than performing reflection calls every
			// time that InvokeMember is called (either caching against the current instance or building a static ConcurrentDictionary that caches the member data for each
			// template) but I'm not sure that the extra complication if worth it considering we're already paying the cost of IDispatch accesses for every operation
			// (since they're already very slow, the extra reflection costs probably aren't too much to worry about).
			_target = target;
			_template = template;
			_logger = logger;
		}

		public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
			try
			{
				return InvokeMemberInner(name, invokeAttr, binder, target, args, modifiers, culture, namedParameters);
			}
			catch (Exception e)
			{
				var errorToLog = GetErrorToLog(e);
				var requireDefaultMember = string.IsNullOrWhiteSpace(name) || (name == DispIdZeroIdentifier);
				_logger.LogIgnoringAnyError(
					LogLevel.Error,
					() => $"FromFileWSCWrapper.InvokeMember failure for {(requireDefaultMember ? "{DefaultMember}" : name)}: {errorToLog.GetType().Name} {errorToLog.Message}"
				);
				throw errorToLog;
			}
		}

		private static Exception GetErrorToLog(Exception e)
		{
			if (e == null)
				throw new ArgumentNullException("e");

			var comException = e as COMException;
			if (comException == null)
				return e;

			try
			{
				COMSurvivableException.RethrowAsOriginalIfPossible(comException);
			}
			catch(Exception comSurvivableException)
			{
				return comSurvivableException;
			}
			return e;
		}

		private object InvokeMemberInner(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
			if (name == null)
				throw new ArgumentNullException("name");
			if (args == null)
				throw new ArgumentNullException("args");

			var requireDefaultMember = string.IsNullOrWhiteSpace(name) || (name == DispIdZeroIdentifier);
			if (!requireDefaultMember && invokeAttr.HasFlag(BindingFlags.GetProperty) && invokeAttr.HasFlag(BindingFlags.InvokeMethod))
			{
				// When VBScript tries to access a member that is not clearly a property or a method (eg. "a = x.Name") then it will include binding flags GetProperty AND
				// AND InvokeMethod but we need to try to work out which of the two it is now because the IDispatchAccess code doesn't like that ambiguity.
				var nameComparison = invokeAttr.HasFlag(BindingFlags.IgnoreCase) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
				if (_template.GetProperties().Any(p => p.Name.Equals(name, nameComparison)))
					invokeAttr = invokeAttr ^ BindingFlags.InvokeMethod; // It's a property so remove InvokeMethod by xor'ing
				else
					invokeAttr = invokeAttr ^ BindingFlags.GetProperty; // It's not a property so remove GetProperty by xor'ing
			}

			if (invokeAttr.HasFlag(BindingFlags.GetProperty))
			{
				object value;
				if (requireDefaultMember)
					value = IDispatchAccess.GetDefaultProperty<object>(_target, args);
				else
					value = IDispatchAccess.GetProperty(_target, name, args);
				return CurrencyToFloatWhereApplicable(value);
			}

			if (requireDefaultMember)
				throw new Exception("Currently there is only support for default member (DispId zero) on GetProperty requests");

			if (invokeAttr.HasFlag(BindingFlags.SetProperty) || invokeAttr.HasFlag(BindingFlags.PutDispProperty) || invokeAttr.HasFlag(BindingFlags.PutRefDispProperty))
			{
				var value = args[args.Length - 1];
				if ((value == null) && invokeAttr.HasFlag(BindingFlags.PutRefDispProperty))
				{
					// When a VBScript statement "Set x = Nothing" becomes an IReflect.InvokeMember call, the invokeAttr will have value PutRefDispProperty (because it's
					// a property setter call for an *object* reference, rather than PutDispProperty - which is for a non-object type) but the value will have been
					// interpreted as null (.net null, not VBScript null). This will be problematic if we try to pass it on to the underlying reference through
					// IDispatch because it will be like us saying "Set x = Empty", which will fail because Empty is not an object. So we need to look out
					// for this case and transform the .net null back into VBScript Nothing.
					value = new DispatchWrapper(null);
				}
				var argsWithoutValues = new object[args.Length - 1];
				Array.Copy(args, argsWithoutValues, args.Length - 1);
				IDispatchAccess.SetProperty(_target, name, value, argsWithoutValues);
				return null;
			}

			if (invokeAttr.HasFlag(BindingFlags.InvokeMethod))
			{
				return CurrencyToFloatWhereApplicable(IDispatchAccess.CallMethod(_target, name, args));
			}

			throw new Exception("Don't know what to do with invokeAttr " + invokeAttr);
		}

		/// <summary>
		/// VBScript throws its toys out the pram in strange ways if a Decimal is returned, but that's what WILL be returned if the VBScript value is a "Currency" (in
		/// VBScript speak). The best that we can do is covnert this into a float and hope that there's no loss in precision for us to worry about.
		/// </summary>
		private static object CurrencyToFloatWhereApplicable(object value)
		{
			if (value is decimal)
				return Convert.ToSingle(value);
			return value;
		}

		public FieldInfo GetField(string name, BindingFlags bindingAttr) { return _template.GetField(name, bindingAttr); }
		public FieldInfo[] GetFields(BindingFlags bindingAttr) { return _template.GetFields(bindingAttr); }
		public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) { return _template.GetMember(name, bindingAttr); }
		public MemberInfo[] GetMembers(BindingFlags bindingAttr) { return _template.GetMembers(bindingAttr); }
		public MethodInfo GetMethod(string name, BindingFlags bindingAttr) { return _template.GetMethod(name, bindingAttr); }
		public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
		{
			return _template.GetMethod(name, bindingAttr, binder, types, modifiers);
		}
		public MethodInfo[] GetMethods(BindingFlags bindingAttr) { return _template.GetMethods(bindingAttr); }
		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr) { return _template.GetProperty(name, bindingAttr); }
		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
		{
			return _template.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
		}
		public PropertyInfo[] GetProperties(BindingFlags bindingAttr) { return _template.GetProperties(bindingAttr); }
		public Type UnderlyingSystemType { get { return _target.GetType().UnderlyingSystemType; } }
	}
}