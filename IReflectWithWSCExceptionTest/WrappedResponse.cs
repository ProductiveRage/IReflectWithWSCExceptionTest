using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace IReflectWithWSCExceptionTest
{
	[ClassInterface(ClassInterfaceType.AutoDispatch)]
	[ComVisible(true)]
	public sealed class WrappedResponse : IReflect
	{
		private readonly Response _target;
		public WrappedResponse(Response target)
		{
			if (target == null)
				throw new ArgumentNullException(nameof(target));

			_target = target;
		}


		public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
			var numberOfUnnamedArgs = (args == null) ? 0 : args.Length;
			var numberOfNamedArgs = (namedParameters == null) ? 0 : namedParameters.Length;

			var isSupportedInvocation =
				((name == "CreateObject") && (target == this) && (numberOfUnnamedArgs == 1) && invokeAttr.HasFlag(BindingFlags.InvokeMethod) && (numberOfNamedArgs == 0)) ||
				((name == "Write") && (target == this) && (numberOfUnnamedArgs == 1) && invokeAttr.HasFlag(BindingFlags.InvokeMethod) && (numberOfNamedArgs == 0)) ||
				((name == "Redirect") && (target == this) && (numberOfUnnamedArgs == 1) && invokeAttr.HasFlag(BindingFlags.InvokeMethod) && (numberOfNamedArgs == 0));
			if (isSupportedInvocation)
				return _target.GetType().InvokeMember(name, invokeAttr, binder, _target, args);

			throw new MissingMemberException($"Invalid InvokeMember call ({name})");
		}


		public Type UnderlyingSystemType { get { return _target.GetType().UnderlyingSystemType; } }

		public FieldInfo GetField(string name, BindingFlags bindingAttr) { return _target.GetType().GetField(name, bindingAttr); }

		public FieldInfo[] GetFields(BindingFlags bindingAttr) { return _target.GetType().GetFields(bindingAttr); }

		public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) { return _target.GetType().GetMember(name, bindingAttr); }

		public MemberInfo[] GetMembers(BindingFlags bindingAttr) { return _target.GetType().GetMembers(bindingAttr); }

		public MethodInfo GetMethod(string name, BindingFlags bindingAttr) { return _target.GetType().GetMethod(name, bindingAttr); }

		public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
		{
			return _target.GetType().GetMethod(name, bindingAttr, binder, types, modifiers);
		}

		public MethodInfo[] GetMethods(BindingFlags bindingAttr) { return _target.GetType().GetMethods(bindingAttr); }

		public PropertyInfo[] GetProperties(BindingFlags bindingAttr) { return _target.GetType().GetProperties(bindingAttr); }

		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr) { return _target.GetType().GetProperty(name, bindingAttr); }

		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
		{
			return _target.GetType().GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
		}
	}
}
