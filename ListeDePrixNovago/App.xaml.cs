using ListeDePrixNovago.Utility.TeamsAuthHelper;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ListeDePrixNovago
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        static App()
        {
            _clientApp = new PublicClientApplication(ClientId, "https://login.microsoftonline.com/common/", TokenCacheHelper.GetUserCache());
            //_clientApp = new PublicClientApplication(ClientId);
        }
        //Below is the clientId of your app registration. 
        //You have to replace the below with the Application Id for your app registration
        private static string ClientId = "CLIENT_ID";

        private static PublicClientApplication _clientApp;

        public static PublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
