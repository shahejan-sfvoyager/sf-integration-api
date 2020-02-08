using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CaseContentExpoter.sforce;
using System.Net;
using System.Web.Services.Protocols;
using System.IO;

namespace CaseContentExpoter
{
    class PartnerSforce
    {
        private CaseContentExpoter.sforce.SforceService binding;
        static void Main(string[] args)
        {
            Console.WriteLine("Please log in to salesforce");
            PartnerSforce pSforce = new PartnerSforce();
            if (pSforce.login())
            {
                //Fetch list of cases
                Console.WriteLine("\nFetching Case Records");
                List<sObject> listofCases = new List<sObject>();
                listofCases = pSforce.queryRecords(pSforce.binding, "Select ID,CaseNumber from Case");
                Console.WriteLine("\nFetched Records :" + listofCases.Count );
                Dictionary<String, String> mapOfCases = new Dictionary<string, string>();
                foreach(sObject caseRec in listofCases)
                {
                    mapOfCases.Add(caseRec.Any[0].InnerText, caseRec.Any[1].InnerText);
                }
                //Fetch attached content
                Console.WriteLine("\nFetching Case Attchaments Records");
                List<sObject> filesofCases = new List<sObject>();
                List<String> caseIds = new List<string>();
                caseIds.AddRange(mapOfCases.Keys);
                String queryStr= "Select ID,body,Name,ParentId from Attachment where ParentId in ('" +string.Join("','", caseIds)+"')";
                filesofCases = pSforce.queryRecords(pSforce.binding, queryStr);
                //Create Files
                Console.WriteLine("\nDownloading Case Attchaments Files");
                pSforce.writeFile(filesofCases, mapOfCases);
            }

        }

        private bool login()
        {
            Console.Write("Enter username: ");
            string username = Console.ReadLine();
            Console.Write("Enter password: ");
            string password = Console.ReadLine();

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            // Create a service object 
            binding = new CaseContentExpoter.sforce.SforceService();

            // Timeout after a minute 
            binding.Timeout = 60000;

            // Try logging in   
            CaseContentExpoter.sforce.LoginResult lr;
            try
            {

                Console.WriteLine("\nLogging in...\n");
                lr = binding.login(username, password);
            }

            // ApiFault is a proxy stub generated from the WSDL contract when     
            // the web service was imported 
            catch (SoapException e)
            {
                // Write the fault code to the console 
                Console.WriteLine(e.Code);

                // Write the fault message to the console 
                Console.WriteLine("An unexpected error has occurred: " + e.Message);

                // Write the stack trace to the console 
                Console.WriteLine(e.StackTrace);

                // Return False to indicate that the login was not successful 
                return false;
            }

            // Check if the password has expired 
            if (lr.passwordExpired)
            {
                Console.WriteLine("An error has occurred. Your password has expired.");
                return false;
            }

            // Set the returned service endpoint URL
            binding.Url = lr.serverUrl;

            // Set the SOAP header with the session ID returned by
            // the login result. This will be included in all
            // API calls.
            binding.SessionHeaderValue = new CaseContentExpoter.sforce.SessionHeader();
            binding.SessionHeaderValue.sessionId = lr.sessionId;

            // Return true to indicate that we are logged in, pointed  
            // at the right URL and have our security token in place.     
            return true;
        }

        private List<sObject> queryRecords(CaseContentExpoter.sforce.SforceService binding,String queryStr)
        {
            List<sObject> listofResponse = new List<sObject>();
            try
            {
                QueryResult qr = null;
                binding.QueryOptionsValue = new sforce.QueryOptions();
                binding.QueryOptionsValue.batchSize = 250;
                binding.QueryOptionsValue.batchSizeSpecified = true;

                qr = binding.query(queryStr);

                bool done = false;
                //int loopCount = 0;
                while (!done)
                {
                    // Process the query results
                    listofResponse.AddRange(qr.records);
                    
                    if (qr.done)
                        done = true;
                    else
                        qr = binding.queryMore(qr.queryLocator);
                }
            }
            catch (SoapException e)
            {
                Console.WriteLine("An unexpected error has occurred: " + e.Message +
                    " Stack trace: " + e.StackTrace);
            }
            return listofResponse;
        }

        private void writeFile(List<sObject> listRecords, Dictionary<String, String> mapOfCases)
        {
            string folderPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName+"\\Cases";
            //manuplating fetched response 
            foreach (sObject attFile in listRecords)
            {
                string filefolder = mapOfCases[attFile.Any[3].InnerText];
                byte[] obj2 = Convert.FromBase64String(attFile.Any[1].InnerText);
                string fileName = attFile.Any[2].InnerText;
                string path = folderPath + "\\" + filefolder + "\\" + fileName;
                bool exists = System.IO.Directory.Exists(folderPath + "\\" + filefolder);
                if (!exists)
                {
                    System.IO.Directory.CreateDirectory(folderPath + "\\" + filefolder);
                }
                FileStream fs1 = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                BinaryWriter bw1 = new BinaryWriter(fs1);
                bw1.Write(obj2);
                bw1.Flush();
                fs1.Close();
                Console.WriteLine("Please Find Exported data in Folder " + folderPath);
                Console.ReadLine();
            }
        }
    }
}
