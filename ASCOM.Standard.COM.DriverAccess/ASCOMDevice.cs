﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ASCOM.Standard.COM.DriverAccess
{
    public abstract class ASCOMDevice : Interfaces.IAscomDevice, IDisposable
    {
        private dynamic device;
        internal dynamic Device
        {
            get => device;
        }

        private short? interfaceVersion = null;

        public ASCOMDevice(string progid)
        {
            device = new DynamicAccess(progid);
        }

        public bool Connected { get => Device.Connected; set => Device.Connected = value; }

        public string Description => Device.Description;

        public string DriverInfo => Device.DriverInfo;

        public string DriverVersion => Device.DriverVersion;

        public short InterfaceVersion
        {
            get
            {
                // Test whether the interface version has already been retrieved
                if (!interfaceVersion.HasValue) // This is the first time the method has been called so get the interface version number from the driver and cache it
                {
                    try { interfaceVersion = Device.InterfaceVersion; } // Get the interface version
                    catch { interfaceVersion = 1; } // The method failed so assume that the driver has a version 1 interface where the InterfaceVersion method is not implemented
                }

                return interfaceVersion.Value; // Return the newly retrieved or already cached value
            }
        }

        public string Name => Device.Name;

        public IList<string> SupportedActions
        {
            get
            {
                if(InterfaceVersion == 1)
                {
                    return new List<string>();
                }
                return (Device.SupportedActions as IEnumerable).Cast<string>().ToList();
            }
        }

        public string Action(string ActionName, string ActionParameters)
        {
            return Device.Action(ActionName, ActionParameters);
        }

        public void CommandBlind(string Command, bool Raw = false)
        {
            Device.CommandBlind(Command, Raw);
        }

        public bool CommandBool(string Command, bool Raw = false)
        {
            return Device.CommandBool(Command, Raw);
        }

        public string CommandString(string Command, bool Raw = false)
        {
            return Device.CommandString(Command, Raw);
        }

        public void Dispose()
        {
            Device.Dispose();
        }

        public void SetupDialog()
        {
            Device.SetupDialog();
        }

        internal void AssertMethodImplemented(int MinimumVersion, string Message)
        {
            if(this.InterfaceVersion < MinimumVersion)
            {
                throw new ASCOM.MethodNotImplementedException(Message);
            }
        }

        internal void AssertPropertyImplemented(int MinimumVersion, string Message)
        {
            if (this.InterfaceVersion < MinimumVersion)
            {
                throw new ASCOM.PropertyNotImplementedException(Message);
            }
        }
    }
}
