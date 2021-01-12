# ASCOM.Standard.COM.DriverAccess

An experimental library to access ASCOM COM drivers from .Net Standard / Core / 5.0. This then translates the interfaces to use the ASCOM Standard version of the interface to behave the same as Alpaca devices.

Currently this reads the registry directly looking for device ProgIDs. It does not check the bitness of the registered drivers. Because .Net Core is not supposed to load from the GAC this uses reflection for all access and does not use any platform interfaces.