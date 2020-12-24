using System;
using System.Collections.Generic;
using System.Text;

namespace ASCOM.Standard.COM.DriverAccess
{
    public class ASCOMRegistration
    {
        internal ASCOMRegistration(string progID, string name)
        {
            this.ProgID = progID;
            this.Name = name;
        }

        public string ProgID
        {
            get;
            private set;
        }

        public string Name
        {
            get;
            private set;
        }
    }
}
