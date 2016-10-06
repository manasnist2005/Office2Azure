using ALEX.Library.SpreadsheetDocument;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Office2Azure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;



namespace Office2Azure
{
    class Program
    {
        static void Main(string[] args)
        {
            string tempFileLocation = System.IO.Directory.GetCurrentDirectory() + SettingsHelper.LocalFolder;
            #region  Copy binary files from Office 365 to Azure Storage
            /*
            List<SPFiles> spFiles = GetOfficeFiles(SettingsHelper.SPDocLibName);
            if (spFiles.Count > 0)
            {                
                DownloadFiles(spFiles, tempFileLocation);
                UploadToAzure(tempFileLocation);
            }
            else
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: No files found to Upload");
            */
            #endregion

            #region Copy data from excel to Azure SQL DB           
            List<SPFiles> spFiles = GetOfficeFiles(SettingsHelper.SPDocLibExcel);
            DownloadFiles(spFiles, tempFileLocation);
            string datapath = tempFileLocation + @"\Data.xlsx";
            List<Person> lstPerson = GetExcelData(datapath);
            DBHelper.SaveData(lstPerson);
            string[] excelFiles = Directory.GetFiles(tempFileLocation);
            foreach(var file in excelFiles)
            {
                System.IO.File.Delete(file);
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: File deleted");
            }
            #endregion
        }
       
        

        /// <summary>
        /// Reads data from excel and fills in the Person class
        /// </summary>
        /// <param name="datapath">Excel file location</param>
        /// <returns></returns>
        public static List<Person> GetExcelData(string datapath)
        {
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* Data fetching from excel started. *******");
            List<Person> lstPerson = new List<Person>();
            try
            {
                var fileName = datapath;
                var sheetName = SettingsHelper.ExcelSheetName; // Existing tab name.
                using (var document = SpreadsheetDocument.Open(fileName, isEditable: false))
                {
                    var workbookPart = document.WorkbookPart;
                    var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(s => s.Name == sheetName);
                    var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Number of row found : " + sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Count().ToString());
                    
                    foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        if (Convert.ToInt32(row.RowIndex.InnerText) > 1)
                        {                           
                            Person prsn = new Person();
                            int columnIndex = 0;
                            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                            {                             
                                if (columnIndex == (int)Columns.Name)
                                {
                                    prsn.Name = GetCellValue(cell, workbookPart);
                                }
                                if (columnIndex == (int)Columns.Age)
                                {
                                    prsn.Age = Convert.ToInt32(GetCellValue(cell, workbookPart));
                                }
                                if (columnIndex == (int)Columns.Gender)
                                {
                                    prsn.Gender = GetCellValue(cell, workbookPart);
                                }
                                if (columnIndex == (int)Columns.City)
                                {
                                    prsn.City = GetCellValue(cell, workbookPart);
                                }
                                if (columnIndex == (int)Columns.Company)
                                {
                                    prsn.Company = GetCellValue(cell, workbookPart);
                                }
                                columnIndex++; 
                            }
                            lstPerson.Add(prsn);
                        }
                        
                    }
                }
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* Data fetching from excel completed. *******");
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + "::Error while reading data from excel. :" + ex.Message);
            }
            return lstPerson;
        }
        /// <summary>
        /// Get cell value
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static string GetCellValue(DocumentFormat.OpenXml.Spreadsheet.Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null)
            {
                return null;
            }

            var value = cell.CellFormula != null
                ? cell.CellValue.InnerText
                : cell.InnerText.Trim();

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    // For shared strings, look up the value in the
                    // shared strings table.
                    var stringTable =
                        workbookPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                    // If the shared string table is missing, something 
                    // is wrong. Return the index that is in
                    // the cell. Otherwise, look up the correct text in 
                    // the table.
                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            return value;
        }
        

        /// <summary>
        /// Connects to Sharepoint online to get the file URLs
        /// </summary>
        /// <returns></returns>
        public static List<SPFiles> GetOfficeFiles(string docLibName)
        {
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File list fetching process from Office 365 started..*******");
            List<SPFiles> spfiles = new List<SPFiles>();
            var context = SettingsHelper.GetUserContext();
            var lstDocs = context.Web.Lists.GetByTitle(docLibName);
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Fetching file details from the Document library : "+ docLibName);
            Folder folderDocs = lstDocs.RootFolder;
            context.Load(lstDocs.RootFolder.Folders);
            context.Load(folderDocs);           
            context.Load(lstDocs.RootFolder.Files);
            context.ExecuteQuery();
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Total files found : "+ folderDocs.Files.Count().ToString());
            foreach (var file in folderDocs.Files)
            {                
                spfiles.Add(new SPFiles { FileName = file.Name, FileUrl = SettingsHelper.O365SiteUrl+ file.ServerRelativeUrl });
            }
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File list fetching process from Office 365 completed..*******");
            return spfiles;
        }
       
        /// <summary>
        /// Download the SP files to a temporary directory
        /// </summary>
        /// <param name="myFiles"></param>
        /// <param name="tempFileLocation"></param>                 
        private static void DownloadFiles(List<SPFiles> myFiles, string tempFileLocation)
        {
            try
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File downloading from Office 365 started..*******");
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Total number of files: " + myFiles.Count.ToString());
                var securePassword = new SecureString();
                foreach (var ch in SettingsHelper.O365Password) securePassword.AppendChar(ch);
                var credentials = new SharePointOnlineCredentials(SettingsHelper.O365UserName, securePassword);
                if (!Directory.Exists(tempFileLocation))
                {
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(tempFileLocation);
                }
                foreach (var file in myFiles)
                {
                    using (var client = new WebClient())
                    {
                        client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        client.Headers.Add("User-Agent: Other");
                        client.Credentials = credentials;
                        client.DownloadFile(file.FileUrl, tempFileLocation + file.FileName);
                        Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: " + file.FileName + " downloaded");
                    }
                }

                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File downloading from Office 365 completed..*******");
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Error while downoading. " + ex.Message );
            }
        }

        /// <summary>
        /// Upload files to Azure from temp location
        /// </summary>
        /// <param name="tempFileLocation"></param>
        private static void UploadToAzure(string tempFileLocation)
        {
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File uploading to Azure started..*******");
            try
            {
                //Get a reference to the storage           
                CloudStorageAccount sa = CloudStorageAccount.Parse(SettingsHelper.AzureConnectionString);
                CloudBlobClient blobClient = sa.CreateCloudBlobClient();

                //Get a reference to the container           
                CloudBlobContainer container = blobClient.GetContainerReference(SettingsHelper.DestinationContainer);

                //Create container if not exist
                container.CreateIfNotExists();

                string[] files = Directory.GetFiles(tempFileLocation);
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Total number of files to be uploaded : " + files.Count().ToString());
                foreach (var file in files)
                {
                    string key = DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + "-" + Path.GetFileName(file);
                    CloudBlockBlob block = container.GetBlockBlobReference(key);
                    using (var fs = System.IO.File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        //Upload file to azure
                        block.UploadFromStream(fs);
                        Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: " + Path.GetFileName(file) + " uploaded");
                    }
                    //Delete files from the temporary location
                    System.IO.File.Delete(file);
                }
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* File uploading to Azure completed..*******");
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Error while uploading. " + ex.Message);
            }
        }

        /*
        public static void GetSiteLists()
        {
         List<SPFiles> myFiles = new List<SPFiles>();
            myFiles.Add(new SPFiles { FileName = "Book1.xlsx", FileUrl = @"https://tieto708-my.sharepoint.com/personal/mpanda_tieto708_onmicrosoft_com/Documents/Book1.xlsx" });
            myFiles.Add(new SPFiles { FileName = "Userinfo.xlsx", FileUrl = @"https://tieto708-my.sharepoint.com/personal/mpanda_tieto708_onmicrosoft_com/Documents/UserInfo.xlsx" });
            myFiles.Add(new SPFiles { FileName = "Apple6S_1.jpg", FileUrl = @"https://tieto708-my.sharepoint.com/personal/mpanda_tieto708_onmicrosoft_com/Documents/Apple6S_1.jpg" });
            myFiles.Add(new SPFiles { FileName = "Architecture.png", FileUrl = @"https://tieto708-my.sharepoint.com/personal/mpanda_tieto708_onmicrosoft_com/Documents/Architecture.png" });


            //using (var clientContext = new ClientContext("https://tieto708.sharepoint.com"))
            //{
            //    // SharePoint Online Credentials 
            var securePassword = new SecureString();
            foreach (var ch in SettingsHelper.O365Password) securePassword.AppendChar(ch);
            //    clientContext.Credentials = new SharePointOnlineCredentials(SettingsHelper.O365UserName, securePassword);
            //    // Get the SharePoint web  
            //    Web web = clientContext.Web;
            //    // Load the Web properties  
            //    clientContext.Load(web);
            //    // Execute the query to the server.  
            //    clientContext.ExecuteQuery();
            //    // Web properties - Display the Title and URL for the web  
            //    Console.WriteLine("Title: " + web.Title + "; URL: " + web.Url);
            //    Console.ReadLine();
            //}
            var context = SettingsHelper.GetUserContext();
            Web wb = context.Web;
            context.Load(wb);
            Site site = context.Site;
            context.Load(site);

            var lists = context.Web.Lists;
            var list = context.Web.Lists.GetByTitle("TTList");
            var lstDocs = context.Web.Lists.GetByTitle("TTDocs");

            CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
            Folder ff = lstDocs.RootFolder;

            FolderCollection fcol = list.RootFolder.Folders;
            List<string> lstFile = new List<string>();


            
            context.Load(lists);
            context.Load(list);
            context.Load(lstDocs);

            context.Load(lstDocs.RootFolder.Folders);
            context.Load(ff);
            context.Load(lstDocs);
            context.Load(lstDocs.RootFolder);
            context.Load(lstDocs.RootFolder.Folders);
            context.Load(lstDocs.RootFolder.Files);
            context.ExecuteQuery();

            Console.WriteLine("Root : " + ff.Name + "\r\n");
            Console.WriteLine(" ItemCount : " + ff.ItemCount.ToString());
            Console.WriteLine(" Folder Count : " + ff.Folders.Count.ToString());
            Console.WriteLine(" File Count : " + ff.Files.Count.ToString());
            Console.WriteLine(" URL : " + ff.ServerRelativeUrl);

            foreach(var file in ff.Files)
            {
                Console.WriteLine("File Url::" + file.ServerRelativeUrl.ToString());
            }

            //foreach (Folder f in fcol)
            //{
            //    if (f.Name == "Item")
            //    {
            //        context.Load(f.Files);
            //        context.ExecuteQuery();
            //        FileCollection fileCol = f.Files;
            //        foreach (Microsoft.SharePoint.Client.File file in fileCol)
            //        {
            //            lstFile.Add(file.Name);
            //            Console.WriteLine(" File Name : " + file.Name);
            //        }
            //    }


            //}


            //camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
            List sslst = lstDocs;
            foreach (var lst in lists)
            {
                Console.WriteLine(lst.Title + "::" + lst.ItemCount.ToString());
                if (lst.Title == "TTDocs")
                {
                    string tt = lst.Title;
                    var dd = lst.GetItems(camlQuery);
                    string fdf = "";
                }

            }

            var kl = lstDocs.GetItems(camlQuery);
            string ffs = "";
            //var ss = context.Site;
            //var usrs = context.Web.SiteUsers;

            // Microsoft.SharePoint.Client.

            // SitesSoapClient ssc = new SitesSoapClient("SitesSoap");
            // //ssc.ClientCredentials = new NetworkCredential(SettingsHelper.O365UserName, securePassword);
            //var ss =  ssc.GetSite("https://tieto708.sharepoint.com/SitePagesaa/Home.aspx");


            // ClientContext clientContext = SettingsHelper.GetUserContext();
            // List list = clientContext.Web.Lists.GetByTitle("TTList");

            // CamlQuery camlQuery = new CamlQuery();
            // camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
            // Folder ff = list.RootFolder;


            // FolderCollection fcol = list.RootFolder.Folders;
            // List<string> lstFile = new List<string>();

            // clientContext.Load(list.RootFolder.Folders);
            // clientContext.Load(ff);
            // clientContext.Load(list);
            // clientContext.Load(list.RootFolder);
            // clientContext.Load(list.RootFolder.Folders);
            // clientContext.Load(list.RootFolder.Files);

            // Console.WriteLine("Root : " + ff.Name + "\r\n");
            // Console.WriteLine(" ItemCount : " + ff.ItemCount.ToString());
            // Console.WriteLine(" Folder Count : " + ff.Folders.Count.ToString());
            // Console.WriteLine(" File Count : " + ff.Files.Count.ToString());
            // Console.WriteLine(" URL : " + ff.ServerRelativeUrl);

            // foreach (Folder f in fcol)
            // {
            //     if (f.Name == "Testing")
            //     {
            //         clientContext.Load(f.Files);
            //         clientContext.ExecuteQuery();
            //         FileCollection fileCol = f.Files;
            //         foreach (Microsoft.SharePoint.Client.File file in fileCol)
            //         {
            //             lstFile.Add(file.Name);
            //             Console.WriteLine(" File Name : " + file.Name);
            //         }
            //     }
            // }

            Console.ReadKey();

        }

        public static List<Person> GetExcelData(string datapath)
        {
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* Data fetching from excel started. *******");
            List<Person> lstPerson = new List<Person>();
            try
            {               
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
               
                string str;
                int rCnt = 0;
                int cCnt = 0;

                xlApp = new Excel.Application();

                xlWorkBook = xlApp.Workbooks.Open(datapath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Number of row found : "+range.Rows.Count.ToString());
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Number of column found : " + range.Columns.Count.ToString());

                for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    Person prsn = new Person();
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {

                        if (cCnt == (int)Columns.Name)
                        {
                            prsn.Name = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        }
                        if (cCnt == (int)Columns.Age)
                        {
                            prsn.Age = (int)((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                        }
                        if (cCnt == (int)Columns.Gender)
                        {
                            prsn.Gender = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        }
                        if (cCnt == (int)Columns.City)
                        {
                            prsn.City = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        }
                        if (cCnt == (int)Columns.Company)
                        {
                            prsn.Company = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        }
                    }
                    lstPerson.Add(prsn);
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: ******* Data fetching from excel completed. *******");
            }
            catch(Exception ex)
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + "::Error while reading data from excel. :"+ex.Message );
            }
           
            return lstPerson;
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        */

    }

    public class SPFiles
    {
        public string FileName { get; set; }
        public string FileUrl { get; set; }
    }

    public class Person
    {
        public string Name { get; set; }
        public string Gender { get; set; }
        public string City { get; set; }
        public string Company { get; set; }
        public int Age { get; set; }
       
    }
    
    public enum Columns
    {
        Name, Gender, Age, City,Company
    }
}
