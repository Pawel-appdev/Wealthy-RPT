using System;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;

namespace Wealthy_RPT
{
    /*
        TextBoxDate Class
    */
    public class TextBoxDate
    {
        public System.Windows.Controls.TextBox tb = new System.Windows.Controls.TextBox();
        private string days;
        private string months;
        private string years;
        private string textdate;

        private bool IsDate(string inputDate)
        {
            DateTime dt;
            return DateTime.TryParse(inputDate, out dt);
        }

        public bool TextBox_KeyDown(System.Windows.Input.KeyEventArgs e)
        {

            bool blnReturn = true;
            int ascii = KeyInterop.VirtualKeyFromKey(e.Key);
            switch (ascii)
            {
                case int n when (n == 8) || (n == 9) || (n == 191) || (n >= 47 && n <= 57):  /* backspace OR tab OR forward slash OR nos 0 to 9*/
                    blnReturn = false;
                    break;
                default:
                    break;
            }
            return blnReturn;
        }

        public bool TextBox_LostFocus(System.Windows.Controls.TextBox tb)
        {
            bool bContinue = false;

            string strtbName = "";
            try { strtbName = tb.Tag.ToString(); } catch { strtbName = "Date"; }

                switch (tb.Text.Length)
                {
                    case 6: // 101102
                        days = tb.Text.Substring(0, 2);
                        years = tb.Text.Substring(tb.Text.Length - 2, 2);
                        months = tb.Text.Substring(3, 2);
                        break;
                    case 8:
                        if (tb.Text.IndexOf("/") > 0) // 10/11/02
                        {
                            days = tb.Text.Substring(0, 2);
                            years = tb.Text.Substring(tb.Text.Length - 2, 2);
                            months = tb.Text.Substring(4, 2);
                        }
                        else // 10112002
                        {
                            days = tb.Text.Substring(0, 2);
                            years = tb.Text.Substring(tb.Text.Length - 4, 4);
                            months = tb.Text.Substring(3, 2);
                        }
                        break;
                    case 10: // 10/11/2002
                        days = tb.Text.Substring(0, 2);
                        years = tb.Text.Substring(tb.Text.Length - 4, 4);
                        months = tb.Text.Substring(3, 2);
                        break;
                    case 11: // 10 Aug 2015 or 10/Aug/2015
                        days = tb.Text.Substring(0, 2);
                        years = tb.Text.Substring(tb.Text.Length - 4, 4);
                        months = tb.Text.Substring(4, 3);
                        break;
                    default:
                        break;
                }
                textdate = days + "/" + months + "/" + years;

                if (!DateTime.TryParse(textdate, out DateTime dt))
                {
                    bContinue = false;
                }
                else
                {
                    bContinue = true;
                    tb.Text = dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }

            if (tb.Text.Length == 0)
            {
                bContinue = true;
            }

            if (bContinue == false)
                {
                    DialogResult warning = MessageBox.Show("The date entered is not a valid date", strtbName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                //tb.Focus(); 
                //tb.SelectAll();
            }
            return bContinue;
        }

    }
}
