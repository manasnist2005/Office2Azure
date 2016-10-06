using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Office2Azure
{
    class SettingsHelper
    {
        //Connection String for Azure storage account : Contains Storage Account name and Primary Access Key
        public static string AzureConnectionString = ConfigurationManager.AppSettings["StorageConnectionString"];

        //A temp Folder to which the files will be downloded from Office 365 and there after will be uploded to Azure
        public static string LocalFolder = ConfigurationManager.AppSettings["sourceFolder"];

        //The name of the container created in the Azure Storage account, which will contain the data
        public static string DestinationContainer = ConfigurationManager.AppSettings["destFolder"];

        //The SharePoint online URL. eg: https://<Any name>.sharepoint.com
        public static string O365SiteUrl = ConfigurationManager.AppSettings["O365SiteUrl"];

        //The user name which has access to the share point online to access data eg: testuser@Tieto708.onmicrosoft.com
        public static string O365UserName = ConfigurationManager.AppSettings["O365UserName"];
        //Password of the Office 365 User
        public static string O365Password = ConfigurationManager.AppSettings["O365Password"];
        //Name of the Document library from which the files will be uploaded to Azure
        public static string SPDocLibName = ConfigurationManager.AppSettings["SPDocLibName"];

        //Name of the Document library which contains the excel sheet
        public static string SPDocLibExcel = ConfigurationManager.AppSettings["SPDocLibExcel"];
        //Name of the Document library which contains the excel sheet
        public static string ExcelSheetName = ConfigurationManager.AppSettings["ExcelSheetName"];

        //Azure Sql database connection string
        public static string AzureDBConnStr = ConfigurationManager.ConnectionStrings["AzureDBConnStr"].ConnectionString;


        /// <summary>
        /// Get the User context in order to connect to SP online
        /// </summary>
        /// <returns></returns>
        public static ClientContext GetUserContext()
        {
            //Secure password
            var o365Pwd = new SecureString();
            foreach (var c in SettingsHelper.O365Password)
            {
                o365Pwd.AppendChar(c);
            }
            var o365Credential = new SharePointOnlineCredentials(SettingsHelper.O365UserName, o365Pwd);
            var o365Context = new ClientContext(SettingsHelper.O365SiteUrl);
            o365Context.Credentials = o365Credential;

            return o365Context;
        }
    }
}
