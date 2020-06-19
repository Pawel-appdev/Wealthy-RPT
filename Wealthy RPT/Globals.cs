using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wealthy_RPT
{
    public class Globals
    {
        // Access
        public static bool blnAccess = false;  // set with mnuAccessibility_Click()
        // form event boolean
        public static bool blnIgnoreEvents = false; // set with MnuNewCase_Click()
        // version number variables
        public static long CurrentVersion = 0;
        // CRM List
        public static List<int> gn_CRM = new List<int>();
        public static bool blnOfficeWiped = false; // set with RPTDetail CboOffice_SelectionChanged()
        public static DataTable dtGraph;
        public static DataTable dtGrid;
    }
}
