
using Microsoft.Graph;
using Microsoft;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Net;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Win32;

namespace GraphAPIConsole
{
    //JSONConfig class so I can use work with the config easier.
    class JSONConfig
    {
        public string tenantID;
        public string clientID;
        public string clientSecret;
        public string email;
        public string signatureAddressHTML;
        public string signatureAddressTXT;
        public string XMLFilesZip;
        public string encapsulator;
        public bool firstRun;

        public bool LoadConfig(string pathToConfig)
        {
            if(System.IO.File.Exists(pathToConfig))
            {
                string config = string.Concat(System.IO.File.ReadAllLines(pathToConfig));
                //Retrieving the RAW data for config, so it isn't an awful one liner.
                try
                {
                    dynamic JSONConfig = JObject.Parse(config);
                    tenantID = JSONConfig.tenantID;
                    clientID = JSONConfig.clientID;
                    clientSecret = JSONConfig.clientSecret;
                    email = JSONConfig.email;
                    signatureAddressHTML = JSONConfig.signatureAddressHTML;
                    signatureAddressTXT = JSONConfig.signatureAddressTXT;
                    XMLFilesZip = JSONConfig.XMLFilesZip;
                    encapsulator = JSONConfig.encapsulator;
                    firstRun = JSONConfig.firstRun;
                }
                //try parsing JSONObject, if it works parse it and return true.
                catch
                {
                    return false;
                }
                return true;
            }
            return false;
        }
        public void RewriteFirstRun(string pathToConfig)
        {
            this.firstRun = false;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            //Getting path to my AppData folder.
            string pathToAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Test");
            JSONConfig jSONConfig = new JSONConfig();
            //If a config exists, continue and log it.
            if (jSONConfig.LoadConfig(pathToAppData + "\\config.json"))
            {
                WriteLog(pathToAppData, "Found config, continuing");
            }
            //If a config doesn't exist, log it and exit, retry on next run.
            else
            {
                WriteLog(pathToAppData, "did not find config.json in AppData");
                System.Environment.Exit(0);
            }
            //Try to retrieve users if it fails, log error and exit.
            try
            {
                getUsersAsync(jSONConfig).GetAwaiter().GetResult(); //retrieving all users to work with them.
            }
            catch (Exception ex)
            {
                Console.WriteLine(jSONConfig.clientID);
                Console.WriteLine(jSONConfig.tenantID);
                Console.WriteLine(jSONConfig.clientSecret);
                WriteLog(pathToAppData, ex.Message + "; Couldn't retrieve any data, check tenantID, clientSecret and ClientID.");
                System.Environment.Exit(0);
            }

            var users = getUsersAsync(jSONConfig).GetAwaiter().GetResult(); //retrieving all users to work with them.
            User signingUser = null;
            //sorting through all the users to find the right one. also foundUser is set to true if I find it, if I do not log and exit in the next if.
            foreach (var user in users)
            {
                if (user.Mail.ToString().ToLower() == jSONConfig.email.ToLower())
                {
                    signingUser = user;
                    WriteLog(pathToAppData, $"Found {signingUser.DisplayName}, with mail \"{signingUser.Mail}\" as the signing user.");
                }
            }

            //If I do not find the right user log and exit.
            if (signingUser == null)
            {
                WriteLog(pathToAppData, $"No user with {jSONConfig.email} has been found.");
                System.Environment.Exit(0);
            }

            //try downloading the files, if it fails log and exit.
            try
            {
                using (var client = new WebClient())
                {
                    client.DownloadFile(jSONConfig.signatureAddressHTML, pathToAppData + "\\test.html");
                }
            }
            catch (Exception ex)
            {
                WriteLog(pathToAppData, $"couldn't download the file in \"{jSONConfig.signatureAddressHTML}\" check network path and connection");
                Environment.Exit(0);
            }

            //downloading the files to use for signatures
            using (var client = new WebClient())
            {
                client.DownloadFile(jSONConfig.signatureAddressHTML, pathToAppData + "\\signature.html");
            }

            using (var client = new WebClient())
            {
                client.DownloadFile(jSONConfig.signatureAddressHTML, pathToAppData + "\\signature.txt");
            }

            using (var client = new WebClient())
            {
                client.DownloadFile(jSONConfig.signatureAddressHTML, pathToAppData + "\\XMLZipFiles.zip");
            }

            string rawSignatureString = System.IO.File.ReadAllText(pathToAppData + "\\signature.html");
            rawSignatureString = RawReplacer(pathToAppData,rawSignatureString, jSONConfig, signingUser);
            System.IO.File.WriteAllText(pathToAppData + "\\signature.html", rawSignatureString);

            #region HTML to RTF, doesn't work yet.
            //Console.WriteLine(signingUser.JobTitle);

            //string convertFromToRtf = System.IO.File.ReadAllText(pathToAppData + "\\signature.html"); // html content
            //RichTextBox rtbTemp = new RichTextBox();
            //WebBrowser wb = new WebBrowser();
            //wb.Navigate("about:blank");
            //wb.Document.Write(convertFromToRtf);

            //while (wb.ReadyState != WebBrowserReadyState.Complete)
            //{
            //    Thread.Sleep(200);
            //    System.Windows.Forms.Application.DoEvents();
            //}

            //wb.Document.ExecCommand("SelectAll", false, null);
            //wb.Document.ExecCommand("Copy", false, null);
            //rtbTemp.SelectAll();
            //rtbTemp.Paste();
            //convertFromToRtf = rtbTemp.ToString();
            //rtbTemp.SaveFile(pathToAppData + "\\signature.rtf", RichTextBoxStreamType.RichText);
            //string a = wb.Document.ToString();
            #endregion
            //I don't actually need an RTF signature, it works with HTM only, so I only need to create a .txt signature and it should be baller
            //I do need to create an RTF signature, the .htm file has a load of Microsoft bullshit Imma ignore it, so does the %%signatureName%%_files folder,
            //it's a bunch of shit I don't know what is, but it works if I copy it, so it should be gucci.

            //Rewrite registryKey only on firstRun of this program, add FirstRun to conf.json
            string keyName = "SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Profiles\\Outlook\\9375CFF0413111d3B88A00104B2A6676";
            RegistryKey keyForOutlookSignature = Registry.CurrentUser.OpenSubKey(keyName);
            string[] subkeys = keyForOutlookSignature.GetSubKeyNames();

            //If firstRun is set to true I am rewriting the registry entry for the signatures
            if (jSONConfig.firstRun)
            {
                WriteLog(pathToAppData, "performing firstRun");
                foreach (string subkeyNames in subkeys)
                {
                    WriteLog(pathToAppData, $"Currently comparing {subkeyNames}");
                    RegistryKey checkingRegistry = Registry.CurrentUser.OpenSubKey(keyName + $"\\{subkeyNames}",true);
                    if ((string)checkingRegistry.GetValue("Account Name") == jSONConfig.email)
                    {
                        WriteLog(pathToAppData,"Found the right address to change keys for.");
                        checkingRegistry.SetValue("5New Signature", "signature");
                        checkingRegistry.SetValue("Reply-Forward Signature", "signature");
                    }
                    Console.WriteLine();
                }
            }

            Console.WriteLine(DateTime.Now.ToLongDateString());
            Console.ReadLine();

        }

        public static string RawReplacer(string pathToAppData, string rawSignatureString, JSONConfig jSONConfig, User signingUser)
        {
            string searchWord = null;
            // foreach Property in SigningUser, this is to replace %%displayName%% into an actual display name, if the value of it is null log and exit

            foreach (PropertyInfo propInfo in signingUser.GetType().GetProperties())
            {
                Console.WriteLine(propInfo.Name + " - " + propInfo.GetValue(signingUser, null));
                //if the property is included in the signature.html I will execute the code inside
                searchWord = Encapsulate(propInfo.Name, jSONConfig.encapsulator);
                if (rawSignatureString.Contains(searchWord))
                {
                    //checking if the value of the property is null, if it is then log and exit
                    if (propInfo.GetValue(signingUser, null) == null)
                    {
                        WriteLog(pathToAppData, $"{propInfo.Name} is null, will not replace");
                        System.Environment.Exit(0);
                    }
                    // if it isn't null I will replace and exit
                    else
                    {
                        WriteLog(pathToAppData, $"{propInfo.Name.ToString()} replacing with {propInfo.GetValue(signingUser, null)}");
                        rawSignatureString = rawSignatureString.Replace(searchWord, propInfo.GetValue(signingUser, null).ToString());
                    }
                }
            }

            return rawSignatureString;
        }

        //if no log exists, create one, if it does append text
        public static void WriteLog(string pathToAppData, string message)
        {
            string pathToLog = pathToAppData + "\\log.txt";
            if (!System.IO.File.Exists(pathToLog))
            {
                System.IO.File.Create(pathToLog);
                Console.WriteLine("");
            }
            message = $"{DateTime.Now.ToLocalTime()} - " + message + Environment.NewLine;
            Thread.Sleep(100);
            System.IO.File.AppendAllText(pathToLog, message);
        }

        //Get all users from the API, after this I read the config.json file and only work with one of the accounts selected
        public async static Task<List<User>> getUsersAsync(dynamic JSONConfig)
        {
            string clientID = JSONConfig.clientID;
            string tenantID = JSONConfig.tenantID;
            string clientSecret = JSONConfig.clientSecret;

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientID)
                .WithTenantId(tenantID)
                .WithClientSecret(clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var users = await graphClient.Users
                        .Request()
                        .Select("*")
                        .GetAsync();
            return users.ToList();
        }

        public static string Encapsulate(string toEncapsulate, string encapsulator)
        {
            return encapsulator + toEncapsulate + encapsulator;
        }
    }
}