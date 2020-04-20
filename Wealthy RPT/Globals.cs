using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wealthy_RPT
{
    public static class Globals
    {
        // Access
        public static bool blnAccess = false;  // set with mnuAccessibility_Click()
        // form event boolean
        public static bool blnIgnoreEvents = false; // set with MnuNewCase_Click()
        // version number variables
        public static long gnVersion = 0;
        public static string gsVersion = "";
        public static long CurrentVersion = 0;

        //public fixed int gs_CRM[2];
        //Public gs_CRM(2) As Integer

    }
}
