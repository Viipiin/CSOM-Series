using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using System.Net;

namespace SharePointCSOMRestSeries
{
    class CSOMDay1
    {
        string siteUrl = ConfigurationManager.AppSettings["siteUrl"];
        string userName = ConfigurationManager.AppSettings["userName"];
        SecureString securePass = GetSecurePassword();

        private static SecureString GetSecurePassword()
        {
            string pass = ConfigurationManager.AppSettings["userPassword"];
            SecureString secPass = new SecureString();
            foreach (char c in pass)
            {
                secPass.AppendChar(c);
            }
            return secPass;
        }

        public void ConnectSharePointOnline()
        {
            
            using (SP.ClientContext context=new SP.ClientContext(siteUrl))
            {
                context.Credentials = new SP.SharePointOnlineCredentials(userName, securePass);
                SP.Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();
                Console.WriteLine("Web Title is "+web.Title);
            }
        }
        
    }
   
}
