using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace Wealthy_RPT
{
    class TextBoxWholeNum
    {
        public System.Windows.Controls.TextBox tb = new System.Windows.Controls.TextBox();

        public bool TextBox_KeyDown(System.Windows.Controls.TextBox tb, System.Windows.Input.KeyEventArgs e)
        {

            bool blnReturn = true;
            int ascii = KeyInterop.VirtualKeyFromKey(e.Key);

            if (tb.Text.Length == 10)
            {
                ascii = ((ascii == 13) || (ascii == 8)) ? ascii : 0;  /* backspace OR horizontal tab */
            }

            switch (ascii)
            {
                case  8:  /* horizontal tab*/
                    blnReturn = false;
                    break;
                case 13:  /* return moves to next tab index*/
                    SendKeys.Send("\t");
                    ascii = 0;
                    blnReturn = false;
                    break;
                case int n when (n < 32) || (n >= 48 && n <= 57):  /* space OR nos 0 to 9*/
                    blnReturn = false;
                    break;
                default:
                    ascii = 0;
                    break;
            }
            return blnReturn;
        }
    }
}
