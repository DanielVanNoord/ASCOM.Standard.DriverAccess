using System;

namespace TestCOMAccess
{
    class Program
    {
        static void Main(string[] args)
        {
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
