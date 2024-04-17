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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ServerEventLogReader
{
    class WPFMessagesClass
    {
        //Public method to get the information
        public void ErrorMessage(string strErrorMessage)
        {
            MessageBox.Show(strErrorMessage, "Please Correct", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        public void InformationMessage(string strInformationMessage)
        {
            MessageBox.Show(strInformationMessage, "Thank You", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        public void UnderDevelopment()
        {
            MessageBox.Show("The Module Is Under Development", "Thank You", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        public void CloseTheProgram()
        {
            const string message = "Are you sure that you would like to close the program?";
            const string caption = "Form Closing";
            MessageBoxResult result = MessageBox.Show(message, caption, MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }

        }
        public void LaunchHelpSite()
        {
            System.Diagnostics.Process.Start("file://bjc/shares/Documents/WAREHOUSE/WhseTrac%20Manual/index.html");
        }
        public void LaunchHelpDeskTickets()
        {
            System.Diagnostics.Process.Start("https://bluejay.on.spiceworks.com/portal/tickets");
        }
        
    }
}
