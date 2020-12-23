using System;
using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ASCOM.Standard.COM.DriverAccess
{
    public class DynamicAccess : DynamicObject
    {
        private object device;

        internal object Device
        {
            get => device;
        }

        public DynamicAccess(string ProgID)
        {
            Type type = Type.GetTypeFromProgID(ProgID);
            device = Activator.CreateInstance(type);
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            try
            {
                result = Unwrap().GetType().InvokeMember(
             binder.Name,
             BindingFlags.InvokeMethod,
             Type.DefaultBinder,
             Unwrap(),
             args
         );

                return true;
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException?.HResult == ErrorCodes.NotImplemented)
                {
                    var member = (string)ex.InnerException.GetType().InvokeMember("PropertyOrMethod", BindingFlags.Default | BindingFlags.GetProperty, null, ex.InnerException, new object[] { }, CultureInfo.InvariantCulture);
                    //TL.LogMessageCrLf(binder.Name, "  Throwing MethodNotImplementedException: '" + member + "'");
                    throw new MethodNotImplementedException(member, ex.InnerException);
                }

                CheckDotNetExceptions(binder.Name, ex);

                throw;
            }
            catch
            {
                throw;
            }
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            try
            {
                Unwrap().GetType().InvokeMember(
            binder.Name,
            BindingFlags.SetProperty,
            Type.DefaultBinder,
            Unwrap(),
            new object[] { value }
        );
                return true;
            }
            catch (TargetInvocationException ex)
            {
                CheckDotNetExceptions(binder.Name, ex);

                if (ex.InnerException?.HResult == ErrorCodes.NotImplemented)
                {
                    var member = (string)ex.InnerException.GetType().InvokeMember("PropertyOrMethod", BindingFlags.Default | BindingFlags.GetProperty, null, ex.InnerException, new object[] { }, CultureInfo.InvariantCulture);
                    //TL.LogMessageCrLf(binder.Name, "  Throwing PropertyNotImplementedException: '" + member + "'");
                    throw new PropertyNotImplementedException(member, false, ex.InnerException);
                }

                throw;
            }
            catch
            {
                throw;
            }
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            try
            {
                result = Unwrap().GetType().InvokeMember(
                binder.Name,
                BindingFlags.GetProperty,
                Type.DefaultBinder,
                Unwrap(),
                new object[] { },
                CultureInfo.InvariantCulture
            );

                return true;
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException?.HResult == ErrorCodes.NotImplemented)
                {
                    var member = (string)ex.InnerException.GetType().InvokeMember("PropertyOrMethod", BindingFlags.Default | BindingFlags.GetProperty, null, ex.InnerException, new object[] { }, CultureInfo.InvariantCulture);
                    //TL.LogMessageCrLf(binder.Name, "  Throwing PropertyNotImplementedException: '" + member + "'");
                    throw new PropertyNotImplementedException(member, false, ex.InnerException);
                }

                CheckDotNetExceptions(binder.Name, ex);

                throw;
            }
            catch
            {
                throw;
            }
        }

        private object Unwrap()
       => Device ?? throw new ObjectDisposedException(nameof(Device));

        /// <summary>
        /// Checks for ASCOM exceptions returned as inner exceptions of TargetInvocationException. When new ASCOM exceptions are created
        /// they must be added to this method. They will then be used in all three cases of Property Get, Property Set and Method call.
        /// </summary>
        /// <param name="memberName">The name of the invoked member</param>
        /// <param name="e">The thrown TargetInvocationException including the inner exception</param>
        private void CheckDotNetExceptions(string memberName, Exception e)
        {
            throw new NotImplementedException();
        }
    }
}