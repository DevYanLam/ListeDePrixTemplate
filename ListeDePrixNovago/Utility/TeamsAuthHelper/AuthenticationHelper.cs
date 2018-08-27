//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ListeDePrixNovago.Utility.TeamsAuthHelper
{
    public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        public static string[] Scopes = { "User.Read", "Sites.Read.All", "Group.Read.All","Files.Read.All", "Directory.Read.All", "Directory.AccessAsUser.All","User.ReadBasic.All" };

        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static async Task<GraphServiceClient> GetAuthenticatedClientAsync()
        {
            var myApp = App.PublicClientApp;
            AuthenticationResult authenticationResult = await myApp.AcquireTokenAsync(Scopes).ConfigureAwait(false);
            GraphServiceClient graphClient = new GraphServiceClient(
              new DelegateAuthenticationProvider(
                  async (requestMessage) =>
                  {
                      // Append the access token to the request.
                      requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);
                  }));


            return graphClient;
        }
    }
}
