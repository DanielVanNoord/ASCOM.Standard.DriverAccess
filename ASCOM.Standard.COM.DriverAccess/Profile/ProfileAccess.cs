using Microsoft.Win32;
using System;
using System.Collections.Generic;

namespace ASCOM.Standard.COM.DriverAccess
{
    public class ProfileAccess
    {
        public static string DriverString(DriverTypes driverType)
        {
            switch (driverType)
            {
                case DriverTypes.Camera:
                    return "Camera";

                case DriverTypes.CoverCalibrator:
                    return "CoverCalibrator";

                case DriverTypes.Dome:
                    return "Dome";

                case DriverTypes.FilterWheel:
                    return "FilterWheel";

                case DriverTypes.Focuser:
                    return "Focuser";

                case DriverTypes.ObservingConditions:
                    return "ObservingConditions";

                case DriverTypes.Rotator:
                    return "Rotator";

                case DriverTypes.SafetyMonitor:
                    return "SafetyMonitor";

                case DriverTypes.Switch:
                    return "Switch";

                case DriverTypes.Telescope:
                    return "Telescope";

                case DriverTypes.Video:
                    return "Video";
            }
            throw new Exception("Unknown device type requested;");
        }

        public static List<ASCOMRegistration> GetDrivers(DriverTypes DeviceType)
        {
            List<ASCOMRegistration> Drivers = new List<ASCOMRegistration>();

            using (var localmachine32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                            RegistryView.Registry32))
            {
                using (var ASCOMKeys = localmachine32.OpenSubKey($"SOFTWARE\\ASCOM\\{DriverString(DeviceType)} Drivers", false))
                {
                    foreach (var key in ASCOMKeys.GetSubKeyNames())
                    {
                        string name = string.Empty;
                        using (var DriverKey = ASCOMKeys.OpenSubKey(key, false))
                        {
                            foreach (var subkey in DriverKey.GetValueNames())
                            {
                                if (subkey == string.Empty)
                                {
                                    name = DriverKey.GetValue(subkey).ToString();
                                }
                            }

                            if (name != string.Empty)
                            {
                                Drivers.Add(new ASCOMRegistration(key, name));
                            }
                        }
                    }
                }
            }

            return Drivers;
        }
    }
}