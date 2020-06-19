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

        public static bool blnUpdate = false;
        // CRM List
        public static List<List<String>> gs_CRM = new List<List<String>>(); // list within a list

        public static bool blnOfficeWiped = false; // set with RPTDetail CboOffice_SelectionChanged()

        public static string[,] arrHist = new string[1,2];
        public static string[] arrPopChanges = new string[1]; // Historical Population Changes

        public static DataTable dtGraph;
        public static DataTable dtGrid;
    }
}
