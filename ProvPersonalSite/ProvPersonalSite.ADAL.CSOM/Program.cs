using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProvPersonalSite.ADAL.CSOM
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var ctx = new AuthenticationManager().GetAzureADNativeApplicationAuthenticatedContext(
            "https://fwilliams-admin.sharepoint.com",
            "Your Own App ID",
            "Your Own Redirect URL"))
            {
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                //string title = web.Title;
                //Console.WriteLine(title);
                //Console.ReadLine();

                Program _spoPersonalSiteCreator = new Program();
                String _adminTenantUrl = "https://fwilliams-admin.sharepoint.com";

                string[] _usersToCreate = { "ignite2017userone@FWilliams.onmicrosoft.com", "ignite2017usertwo@FWilliams.onmicrosoft.com" };
                _spoPersonalSiteCreator.CreatePersonalSiteUsingCSOM(_adminTenantUrl, _usersToCreate);

            }
        }

        public void CreatePersonalSiteUsingCSOM(string tenantAdminUrl, string[] emailIDs)
        {
            using (ClientContext _context = new ClientContext(tenantAdminUrl))
            {
                try
                {

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


    }
}
