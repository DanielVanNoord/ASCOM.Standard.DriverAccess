using System;

namespace TestCOMAccess
{
    class Program
    {
        static void Main(string[] args)
        {

            var test2 = ASCOM.Standard.COM.DriverAccess.Camera.Cameras;

            using (var covcal = new ASCOM.Standard.COM.DriverAccess.CoverCalibrator("ASCOM.Simulator.CoverCalibrator"))
            {
                covcal.SetupDialog();
                var test = covcal.SupportedActions;
            }

            using (var Camera = new ASCOM.Standard.COM.DriverAccess.Camera("ASCOM.Simulator.Camera"))
            {
                Camera.SetupDialog();
                var test = Camera.SupportedActions;
            }

            using (var Telescope = new ASCOM.Standard.COM.DriverAccess.Telescope("ASCOM.Simulator.Telescope"))
            {
                Telescope.Connected = true;
                var test5 = Telescope.AxisRates(ASCOM.Standard.Interfaces.TelescopeAxis.Primary);
                var test4 = Telescope.TrackingRates;
                var test = Telescope.SupportedActions;
            }

            using (var Switch = new ASCOM.Standard.COM.DriverAccess.Switch("ASCOM.Simulator.Switch"))
            {
                Switch.Connected = true;
                var test = Switch.SupportedActions;
            }

            using (var SafetyMonitor = new ASCOM.Standard.COM.DriverAccess.SafetyMonitor("ASCOM.Simulator.SafetyMonitor"))
            {
                SafetyMonitor.Connected = true;
                var test = SafetyMonitor.SupportedActions;
            }


            using (var ObservingConditions = new ASCOM.Standard.COM.DriverAccess.ObservingConditions("ASCOM.Simulator.ObservingConditions"))
            {
                ObservingConditions.Connected = true;
                var test = ObservingConditions.SupportedActions;
            }

            using (var Dome = new ASCOM.Standard.COM.DriverAccess.Dome("ASCOM.Simulator.Dome"))
            {
                var test = Dome.SupportedActions;
            }

            using (var covcal = new ASCOM.Standard.COM.DriverAccess.CoverCalibrator("ASCOM.Simulator.CoverCalibrator"))
            {
                var test = covcal.SupportedActions;
            }

            using (var rotator = new ASCOM.Standard.COM.DriverAccess.Rotator("ASCOM.SimulatorLS.CoverCalibrator"))
            {
                var test = rotator.SupportedActions;
            }

            using (var wheel = new ASCOM.Standard.COM.DriverAccess.FilterWheel("ASCOM.Simulator.FilterWheel"))
            {
                wheel.Connected = true;
                Console.WriteLine(wheel.Position);
            }

            using (var focuser = new ASCOM.Standard.COM.DriverAccess.Focuser("ASCOM.Simulator.Focuser"))
            {

                try
                {
                    focuser.Connected = true;
                    Console.WriteLine(focuser.Connected);
                    focuser.Halt();
                    Console.WriteLine(focuser.StepSize);
                }
                catch (ASCOM.PropertyNotImplementedException ex)
                {

                }
                catch (ASCOM.NotConnectedException)
                {

                }
                catch (Exception ex)
                {
                    var HResult = (UInt32)ex.HResult;

                    var type = ex.GetType();

                }

                focuser.SetupDialog();

                Console.WriteLine(focuser.Name);
            }

            Console.ReadLine();
        }
    }
}
