using System;
using System.Collections.Generic;
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

namespace DecsWordAddIns
{
    internal class GitLabHandler
    {
        private const string BASE_ADDRESS = @"https://ctri-gitlab.ucsd.edu/api/v4/projects/238/repository/files/";
        private string token;
        private string userName;

        internal GitLabHandler()
        {
            GetGitLabToken();
        }

        private void GetGitLabToken()
        {
            // If we can't read an existing token ...
            if (!ReadGitLabToken())
            {
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

        internal async Task PushFile(string path)
        {
            string fullProjectDirectory = Path.GetDirectoryName(path);
            string projectDirectory = Path.GetFileName(fullProjectDirectory);
            string justTheFilename = Path.GetFileName(path);
            string urlExtended = Path.Combine(projectDirectory, justTheFilename);

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("branch", "master");
            parameters.Add("author_email", this.userName + "@ucsd.edu");
            parameters.Add("author_name", TranslateLoginName(this.userName));
            parameters.Add("commit_message", "Automated project setup");
            parameters.Add("content", ReadFile(path));

            // https://stackoverflow.com/a/48930280/18749636
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(BASE_ADDRESS);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", this.token);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //var contentObj = new FormUrlEncodedContent(parameters);
                var result = await client.PostAsJsonAsync(urlExtended, parameters);
                string resultContent = await result.Content.ReadAsStringAsync();
            }
        }

        private string ReadFile(string path)
        {
            if (!File.Exists(path))
            {
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
                }
                catch 
                {
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
        private string TranslateLoginName(string loginName)
        {
            Dictionary<string, string> userNamesList = Utilities.ReadUserNamesFile();

            if (userNamesList != null && userNamesList.ContainsKey(loginName))
            {
                return userNamesList[loginName];
            }

            return loginName;
        }
    }
}