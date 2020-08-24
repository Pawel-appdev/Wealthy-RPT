using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for About.xaml
    /// </summary>
    public partial class About : Window
    {
        public About()
        {
            InitializeComponent();

            //lblApplicationName.Content = Global.ApplicationName;
            lblBDAppName.Content = "BDApp Name : " + Global.ApplicationName;
            lblBDAppNumber.Content = "BDApp Number : " + Global.BDAppNo;
            lblResolver.Content = "Resolver Group : " + Global.ResolverGroup;
            lblVersion.Content = "Version : " + Global.CurrentVersion;
        }

        private void ImgLogo_MouseDown(object sender, MouseButtonEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("This will initiate the reinstall process.  Click Ok to proceed", Global.ApplicationName, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);

            if (result == MessageBoxResult.OK)
            {
                Global.updateApplication();
            }
        }
    }
}
