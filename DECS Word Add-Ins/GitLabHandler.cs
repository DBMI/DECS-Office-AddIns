using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
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
    /**
     * @brief Communicates with GitLab to instantiate project.
     */ 
    internal class GitLabHandler
    {
        private const string DIVIDER = "%2F";
        private const string QUOTES = "\"";
        private const string VIEW_ADDRESS =
            @"https://ctri-gitlab.ucsd.edu/cdwr/data-concierge/-/tree/master/";
        private string token;
        private string userName;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        /// <summary>
        /// Constructor
        /// Gets GitLab token, either by looking up saved value or asking user for it.
        /// </summary>
        internal GitLabHandler()
        {
            LogManager.GetRepository().Threshold = Level.Debug;
            log.Debug("Instantiating GitLabHandler object.");
            GetGitLabToken();
        }

        /// <summary>
        /// Lets external code get the GitLab URL.
        /// </summary>
        /// <returns></returns>
        internal static string Address()
        {
            return VIEW_ADDRESS;
        }

        /// <summary>
        /// Gets the GitLab authorization token, either by reading the saved value (if possible) or
        /// asking the user to provide it.
        /// </summary>
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

        /// <summary>
        /// Uses externally-compiled .exe file to push file to GitLab.
        /// </summary>
        /// <param name="path">path to SQL file</param>
        /// <returns>bool</returns>
        internal bool PushFileExe(string path)
        {
            bool success = false;

            // Compiled Python script expects "/" as path separators.
            string pathCorrected = path.Replace(@"\", "/");

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = "--file " + QUOTES + pathCorrected + QUOTES;
            startInfo.CreateNoWindow = true;
            var fullpath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Resources",
                "git_uploader.exe"
            );

            if (!File.Exists(fullpath))
            {
                log.Error("Unable to find GitLabUploader executable '" + fullpath + "'.");
                return false;
            }

            startInfo.FileName = fullpath;

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

        //private string ReadFile(string path)
        //{
        //    if (!File.Exists(path))
        //    {
        //        log.Error("Unable to find file '" + path + "'.");
        //        throw new FileNotFoundException("Unable to find file '" + path + "'.");
        //    }

        //    // Escape any single quotes.
        //    string contents = File.ReadAllText(path).Replace("'", "\'");
        //    return contents;
        //}

        /// <summary>
        /// Reads GitLab token from its expected location.
        /// </summary>
        /// <returns>bool</returns>
        private bool ReadGitLabToken()
        {
            bool success = false;
            this.userName = Utilities.GetUserName();
            string tokenFilename = Path.Combine(
                @"C:\Users",
                this.userName,
                ".ssh",
                "gitlab_api_token.txt"
            );

            if (File.Exists(tokenFilename))
            {
                try
                {
                    this.token = File.ReadAllText(tokenFilename);
                    success = !string.IsNullOrEmpty(this.token);
                    log.Debug(
                        "Reading file '" + tokenFilename + "' resulted in " + success.ToString()
                    );
                }
                catch
                {
                    log.Error("Error when trying to read file '" + tokenFilename + "'.");
                }
            }

            return success;
        }

        /// <summary>
        /// Have we retrieved the token?
        /// </summary>
        /// <returns>bool</returns>
        internal bool Ready()
        {
            return !string.IsNullOrEmpty(this.token);
        }

        /// <summary>
        /// Saves a newly-provided token to the file where we can find it next time.
        /// </summary>
        private void SaveGitLabToken()
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string tokenFilename = Path.Combine(
                @"C:\Users",
                userName,
                ".ssh",
                "gitlab_api_token.txt"
            );

            using (StreamWriter writer = new StreamWriter(tokenFilename))
            {
                writer.WriteLine(this.token);
            }
        }
    }
}
