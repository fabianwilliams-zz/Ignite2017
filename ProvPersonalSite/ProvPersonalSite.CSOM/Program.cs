using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ProvPersonalSite.CSOM
{
    class Program
    {
        static void Main(string[] args)
        {
            Program _spoPersonalSiteCreator = new Program();
            String _adminTenantUrl = "https://fwilliams-admin.sharepoint.com";
            String _userName = "UrOwnUserName@fwilliams.onmicrosoft.com";
            String _passWord = "UrOwnPassword";

            //Create Personal Sites using CSOM
            string[] _usersToCreate = { "ignite2017userone@FWilliams.onmicrosoft.com", "ignite2017usertwo@FWilliams.onmicrosoft.com" };
            _spoPersonalSiteCreator.CreatePersonalSiteUsingCSOM(_adminTenantUrl, _userName, _passWord, _usersToCreate);
        }


        public void CreatePersonalSiteUsingCSOM(string tenantAdminUrl, string userName, string password, string[] emailIDs)
        {
            using (ClientContext _context = new ClientContext(tenantAdminUrl))
            {
                try
                {
                    SharePointOnlineCredentials _creds = new SharePointOnlineCredentials(userName, ConvertToSecureString(password));
                    _context.Credentials = _creds;
                    _context.ExecuteQuery();

                    ProfileLoader _profileLoader = ProfileLoader.GetProfileLoader(_context);
                    _profileLoader.CreatePersonalSiteEnqueueBulk(emailIDs);
                    _profileLoader.Context.ExecuteQuery();
                }
                catch (Exception _ex)
                {
                    throw;
                }
            }
        }

        private SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }

    }
}
