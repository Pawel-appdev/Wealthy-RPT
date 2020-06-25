using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Wealthy_RPT
{
    class Forms
    {
        public static void reloadForm(double dUTR, int intTab, double dPercentile, string sPop)
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w.Name == "frmRPTdetail")
                {
                    w.Close();
                }
            }

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            RPT_Detail rptDetail = new RPT_Detail();  // initialise form
            RPT.RPT_Data rpt = new RPT.RPT_Data(); // initialise data
            rpt.GetRPDData(dUTR, intTab, dPercentile, sPop);
            rptDetail.DataContext = rpt;

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            rptDetail.tabRPD.SelectedIndex = intTab;
            rptDetail.Show();
        }
    }
}
