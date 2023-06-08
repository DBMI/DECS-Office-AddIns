using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.ComponentModel;
using log4net.Core;
using log4net;

namespace DecsWordAddIns
{
    internal class GitLabHandler
    {
        private const string BASE_ADDRESS = @"https://ctri-gitlab.ucsd.edu/api/v4/projects/238/repository/files/";
        private const string DIVIDER = "%2F";
        private const string QUOTES = "\"";
        private string token;
        private string userName;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        internal GitLabHandler()
        {
            LogManager.GetRepository().Threshold = Level.Debug;
            log.Debug("Instantiating GitLabHandler object.");
            GetGitLabToken();
        }

        private void GetGitLabToken()
        {
            // If we can't read an existing token ...
            if (!ReadGitLabToken())
            {
                log.Debug("Asking user to create new GitLab token.");

                // ...ask user to create a new one.
                using (var form = new TokenForm())
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        this.token = form.token;

                        //  and save it in file for next time.
                        SaveGitLabToken();
                    }
                }
            }
        }

        internal bool PushFileExe(string path)
        {
            bool success = false;

            // Compiled Python script expects "/" as path separators.
            string pathCorrected = path.Replace(@"\", "/");

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = "--file " + QUOTES + pathCorrected + QUOTES;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = @"Resources\git_uploader.exe";

            if (!File.Exists(startInfo.FileName))
            {
                throw new FileNotFoundException(startInfo.FileName);
            }

            // https://stackoverflow.com/a/31650828/18749636
            startInfo.UseShellExecute = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                    int returnValue = exeProcess.ExitCode;
                    success = returnValue == 0;
                }
            }
            catch
            {
                log.Error("Received error when starting external process.");
            }

            return success;
        }

        //internal async Task<bool> PushFile(string path)
        //{
        //    string fullProjectDirectory = Path.GetDirectoryName(path);
        //    string projectDirectory = Path.GetFileName(fullProjectDirectory);
        //    string justTheFilenameAndExt = Path.GetFileName(path);
        //    string urlExtended = BASE_ADDRESS + projectDirectory + DIVIDER + justTheFilenameAndExt;

        //    Dictionary<string, string> parameters = new Dictionary<string, string>();
        //    parameters.Add("branch", "master");
        //    parameters.Add("author_email", this.userName + "@ucsd.edu");
        //    parameters.Add("author_name", Utilities.TranslateLoginName(this.userName));
        //    parameters.Add("commit_message", "Automated project setup");
        //    parameters.Add("content", ReadFile(path));
        //    var parametersJson = JsonSerializer.Serialize(parameters);
        //    var data = new StringContent(parametersJson, Encoding.UTF8, "application/json");

        //    var productValue = new ProductInfoHeaderValue("python-requests", "2.28.2");

        //    // https://stackoverflow.com/a/48930280/18749636
        //    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        //    using (var client = new HttpClient())
        //    {
        //        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", this.token);
        //        client.DefaultRequestHeaders.Accept.Clear();
        //        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
        //        client.DefaultRequestHeaders.AcceptEncoding.Add(new StringWithQualityHeaderValue("gzip"));
        //        client.DefaultRequestHeaders.Add("Connection", "keep-alive");
        //        client.DefaultRequestHeaders.UserAgent.Add(productValue);

        //        HttpResponseMessage response = await client.PostAsync(urlExtended, data);

        //        return response.IsSuccessStatusCode;
        //    }
        //}

        private string ReadFile(string path)
        {
            if (!File.Exists(path))
            {
                log.Error("Unable to find file '" + path + "'.");
                throw new FileNotFoundException("Unable to find file '" + path + "'.");
            }

            // Escape any single quotes.
            string contents = File.ReadAllText(path).Replace("'", "\'");
            return contents;
        }

        private bool ReadGitLabToken()
        {
            bool success = false;
            this.userName = Utilities.GetUserName();
            string tokenFilename = Path.Combine(@"C:\Users", this.userName, ".ssh", "gitlab_api_token.txt");

            if (File.Exists(tokenFilename))
            {
                try
                {
                    this.token = File.ReadAllText(tokenFilename);
                    success = !string.IsNullOrEmpty(this.token);
                    log.Debug("Reading file '" + tokenFilename + "' resulted in " + success.ToString());
                }
                catch 
                {
                    log.Error("Error when trying to read file '" + tokenFilename + "'.");
                }
            }

            return success;
        }

        internal bool Ready()
        {
            return !string.IsNullOrEmpty(this.token);
        }

        private void SaveGitLabToken()
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string tokenFilename = Path.Combine(@"C:\Users", userName, ".ssh", "gitlab_api_token.txt");

            using (StreamWriter writer = new StreamWriter(tokenFilename))
            {
                writer.WriteLine(this.token);
            }
        }
    }
}