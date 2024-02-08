using log4net;
using log4net.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Resources.ResXFileRef;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using MsOutlook = Microsoft.Office.Interop.Outlook;

namespace DecsWordAddIns
{
    /**
     * @brief Parses the project Scope of Work document, extracts key field & sets up the DECS project:
     * - creates local project directory
     * - initializes output Excel file with desired disclaimer page
     * - initializes stub of SQL file with project infomation
     * - converts SlicerDicer code (if applicable)
     * - drafts the completion email
     * - pushes SQL file to GitLab
     */ 
    internal class ScopeOfWorkParser
    {
        private string dataSetName;
        private string dataSource;
        private string documentDirectoryName;
        private string outputFileName;
        private string principalInvestigatorEmail;
        private string principalInvestigatorGivenName;
        private string principalInvestigatorSurname;
        private DirectoryInfo projectDirectory;
        private string projectTriple;
        private string requesterEmail;
        private string requesterName;
        private Document scopeOfWork;
        private string sqlFilename;
        private string taskNumber;
        private string studyName;

        private const string DATA_SET_NAME_HEADING = "Data Set Name:";
        private const string DATA_SOURCE_HEADING = "Data Source:";
        private const int MAX_STRING_LENGTH = 32;
        private const string PEOPLE_PATTERN =
            @"(?<surname>[\w ]+), ?(?<given_name>[\w \.]+?) ?(?<email>[\d\w\.]+@[\d\w\.]+)";
        private const string PRINCIPAL_INVESTIGATOR_HEADING =
            "Principal Investigator (Name, E-mail):";
        private const string REQUESTER_HEADING = "Requester (Name, E-mail):";
        private const string STUDY_NAME_HEADING = "Study Name:";
        private const string TASK_NUMBER_HEADING = "DECS Request #:";
        private const string TASK_NUMBER_PATTERN = @"DECS-(\d+)";

        private const string SLICER_DICER_CARDINALITY = "FORCE_DEFAULT_CARDINALITY_ESTIMATION";
        private const string SLICER_DICER_COUNT_BIG = "COUNT_BIG";
        private const string SLICER_DICER_DURABLE_KEY = "DurableKey";
        private const string SLICER_DICER_GROUP_BY = "GROUP BY";
        private const string SLICER_DICER_ISOLATION = "SET TRANSACTION ISOLATION LEVEL SNAPSHOT";

        private Regex decsNumberRegex;
        private Regex peopleRegex;

        private ProgressForm progressForm;

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        internal ScopeOfWorkParser()
        {
            LogManager.GetRepository().Threshold = Level.Debug;
            log.Debug("Instantiating ScopeOfWorkParser.");
            BuildRegex();
            progressForm = new ProgressForm();
            progressForm.Show();
        }

        /// <summary>
        /// Create all the reusable Regex objects. 
        /// </summary>
        private void BuildRegex()
        {
            decsNumberRegex = new Regex(TASK_NUMBER_PATTERN);
            peopleRegex = new Regex(PEOPLE_PATTERN);
        }

        /// <summary>
        /// Drafts SQL file with project header.
        /// </summary>
        /// <returns></returns>
        private bool BuildSqlFile()
        {
            sqlFilename = Path.Combine(
                projectDirectory.FullName,
                projectTriple + ".sql"
            );
            log.Debug("Will build file '" + sqlFilename + "'.");

            if (!InsertSqlHeader())
            {
                return false;
            }

            string slicerDicerFilename = GetSlicerDicerFilename();

            if (string.IsNullOrEmpty(slicerDicerFilename))
            {
                // If there IS no Slicer Dicer file, then we were successful.
                return true;
            }

            // Since there IS a SlicerDicer file to convert, enable this section.
            progressForm.EnableSlicerDicer();

            if (!CopySqlBody(slicerDicerFilename))
            {
                return false;
            }

            if (!WriteConsentSection())
            {
                return false;
            }

            // Can't do this in SetupProject, because it can't tell if this method returned true
            // because it converted the SlicerDicer file or because there isn't one.
            progressForm.CheckOffConvertSlicerDicer();
            progressForm.LinkConvertedSlicerDicerFile(sqlFilename);

            return true;
        }

        /// <summary>
        /// Removes control characters from string.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string Clean(string input)
        {
            return new string(input.Where(c => !char.IsControl(c)).ToArray()).Trim();
        }

        /// <summary>
        /// Copies existing SlicerDicer code, converting to proper format.
        /// </summary>
        /// <param name="slicerDicerFile"></param>
        /// <returns>bool </returns>
        private bool CopySqlBody(string slicerDicerFile)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(sqlFilename, append: true))
                {
                    foreach (var line in File.ReadLines(slicerDicerFile))
                    {
                        if (line.Trim().Length == 0)
                        {
                            // Don't copy blank lines.
                            continue;
                        }

                        if (line.Contains(SLICER_DICER_ISOLATION))
                        {
                            // Don't copy this line to new file.
                            continue;
                        }

                        if (line.Contains(SLICER_DICER_COUNT_BIG))
                        {
                            // Substitute this line instead.
                            writer.WriteLine("\t\tDurableKey");
                            continue;
                        }

                        if (line.Contains(SLICER_DICER_GROUP_BY))
                        {
                            if (!line.Contains(SLICER_DICER_DURABLE_KEY))
                            {
                                writer.WriteLine(line.Replace("\n", "") + ", DurableKey\n");
                            }

                            continue;
                        }

                        if (line.Contains(SLICER_DICER_CARDINALITY))
                        {
                            // Then we're done copying.
                            // Copy this line over & skip the rest.
                            writer.WriteLine(line);
                            break;
                        }

                        // Then just copy input to output;
                        writer.WriteLine(line);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Creates draft Excel file to hold output.
        /// </summary>
        /// <returns>bool</returns>
        private bool CreateOutputFile()
        {
            try
            {
                log.Debug("About to copy file to '" + projectDirectory.FullName + "'.");
                outputFileName = Path.Combine(
                    projectDirectory.FullName,
                    projectTriple + ".xlsx"
                );

                // Copy results template to project directory, allowing overwrite.
                var fullpath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Resources",
                    "results_template.xlsx"
                );

                if (File.Exists(fullpath))
                {
                    File.Copy(fullpath, outputFileName, true);
                    return true;
                }

                log.Error("Unable to find file '" + fullpath + "'.");
            }
            catch (Exception ex)
            {
                log.Error("Unable to copy file to project directory because: " + ex.Message);
            }

            return false;
        }

        /// <summary>
        /// Creates the project directory path. Returns whether the new directory exists.
        /// </summary>
        /// <returns>bool</returns>
        private bool CreateProjectDirectory()
        {
            string targetDirectory = Path.Combine(documentDirectoryName, projectTriple);
            projectDirectory = Directory.CreateDirectory(targetDirectory);
            return projectDirectory.Exists;
        }

        /// <summary>
        /// Tests whether all setup steps were completed successfully.
        /// </summary>
        /// <returns>bool</returns>
        private bool Done()
        {
            bool haveStudyNameOrDataSetName =
                !string.IsNullOrEmpty(studyName) || !string.IsNullOrEmpty(dataSetName);
            return haveStudyNameOrDataSetName
                && !string.IsNullOrEmpty(dataSource)
                && !string.IsNullOrEmpty(principalInvestigatorEmail)
                && !string.IsNullOrEmpty(principalInvestigatorGivenName)
                && !string.IsNullOrEmpty(principalInvestigatorSurname)
                && !string.IsNullOrEmpty(requesterEmail)
                && !string.IsNullOrEmpty(requesterName)
                && !string.IsNullOrEmpty(taskNumber);
        }

        /// <summary>
        /// Pulls the DECS number out of text.
        /// </summary>
        /// <param name="text"></param>
        private void ExtractDecsNumber(string text)
        {
            Match decsNumberMatch = decsNumberRegex.Match(text);

            if (decsNumberMatch.Success)
            {
                taskNumber = decsNumberMatch.Groups[1].Value.ToString().Trim();
            }
        }

        /// <summary>
        /// Pulls the principal investigator's name & email from text.
        /// </summary>
        /// <param name="text"></param>
        private void ExtractPI(string text)
        {
            Match piMatch = peopleRegex.Match(text);

            if (piMatch.Success)
            {
                principalInvestigatorGivenName = piMatch.Groups["given_name"].Value
                    .ToString()
                    .Trim();
                principalInvestigatorSurname = piMatch.Groups["surname"].Value
                    .ToString()
                    .Trim();
                principalInvestigatorEmail = piMatch.Groups["email"].Value.ToString().Trim();
            }
        }

        /// <summary>
        /// Pulls the requestor's name & email from text.
        /// </summary>
        /// <param name="text"></param>
        private void ExtractRequester(string text)
        {
            Match requesterMatch = peopleRegex.Match(text);

            if (requesterMatch.Success)
            {
                string requesterSurname = requesterMatch.Groups["surname"].Value.ToString().Trim();
                string requesterGivenName = requesterMatch.Groups["given_name"].Value
                    .ToString()
                    .Trim();
                requesterName = requesterGivenName + " " + requesterSurname;
                requesterEmail = requesterMatch.Groups["email"].Value.ToString().Trim();
            }
        }

        /// <summary>
        /// Gets & cleans the next paragraph.
        /// </summary>
        /// <param name="desiredIndex"></param>
        /// <returns>string</returns>
        private string GetNextLine(int desiredIndex)
        {
            string nextLine = String.Empty;

            if (desiredIndex >= 0 && desiredIndex < scopeOfWork.Paragraphs.Count)
            {
                nextLine = Clean(scopeOfWork.Paragraphs[desiredIndex].Range.Text.ToString());
            }

            return nextLine;
        }

        /// <summary>
        /// Creates @c OpenFileDialog form to ask for location of SlicerDicer file.
        /// </summary>
        /// <returns>string</returns>
        private string GetSlicerDicerFilename()
        {
            string filePath = string.Empty;

            using (var form = new YesNoForm())
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK && form.fileExists)
                {
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        // Because we don't specify an opening directory,
                        // the dialog will open in the last directory used.
                        openFileDialog.Filter = "SQL files (*.sql)|*.sql";
                        openFileDialog.RestoreDirectory = true;

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            // Get the path of specified file.
                            filePath = openFileDialog.FileName;
                        }
                    }
                }
            }

            return filePath;
        }

        /// <summary>
        /// Initializes the SQL file with header containing:
        /// - project triple: PI-number-description
        /// - task number
        /// - PI name & email
        /// - requester name & email
        /// - author name & email
        /// - date created
        /// - database name
        /// </summary>
        /// <returns>bool</returns>
        private bool InsertSqlHeader()
        {
            try
            {
                log.Debug("About to use Streamwriter to create SQL file.");

                using (StreamWriter writer = new StreamWriter(sqlFilename))
                {
                    writer.WriteLine("/*");
                    writer.WriteLine("** " + projectTriple + ".sql");
                    writer.WriteLine("** Task: " + taskNumber);
                    writer.WriteLine(
                        "** Principal Investigator: "
                            + principalInvestigatorGivenName
                            + " "
                            + principalInvestigatorSurname
                            + ", "
                            + principalInvestigatorEmail
                    );
                    writer.WriteLine(
                        "** Requester: " + requesterName + ", " + requesterEmail
                    );
                    writer.WriteLine(
                        "** Author: "
                            + Environment.UserName
                            + ", "
                            + UserPrincipal.Current.EmailAddress
                    );
                    writer.WriteLine("** Created: " + DateTime.Now.ToString("yyyy-MM-dd"));
                    writer.WriteLine("** Database: " + dataSource);
                    writer.WriteLine("*/");
                    writer.WriteLine("");
                    writer.WriteLine("USE [" + dataSource.ToUpper() + "];");
                }

                return true;
            }
            catch (Exception ex)
            {
                log.Error("Unable to use StreamWriter to create SQL file because: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Runs through the Word document, looking for project information.
        /// </summary>
        /// <returns>bool</returns>
        internal bool Parse()
        {
            int index = 1;
            string nextLine;

            while (!Done() && index < scopeOfWork.Paragraphs.Count)
            {
                Paragraph paragraph = scopeOfWork.Paragraphs[index];

                if (paragraph != null)
                {
                    string text = paragraph.Range.Text.ToString().Trim();

                    if (text != null)
                    {
                        string textCleaned = Clean(text);

                        switch (textCleaned)
                        {
                            case DATA_SET_NAME_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                dataSetName = nextLine;
                                break;

                            case DATA_SOURCE_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                dataSource = nextLine;
                                break;

                            case PRINCIPAL_INVESTIGATOR_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                ExtractPI(nextLine);
                                break;

                            case REQUESTER_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                ExtractRequester(nextLine);
                                break;

                            case STUDY_NAME_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                studyName = nextLine;
                                break;

                            case TASK_NUMBER_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                ExtractDecsNumber(nextLine);
                                break;

                            default:
                                break;
                        }
                    }
                }

                index++;
            }

            projectTriple = ProjectTriple();
            return Done();
        }

        /// <summary>
        /// Forms the string PI-task number-description
        /// </summary>
        /// <returns>string</returns>
        private string ProjectTriple()
        {
            log.Debug(
                "Building project triple from"
                    + " surname: "
                    + principalInvestigatorSurname
                    + " task number: "
                    + taskNumber
                    + " study name: "
                    + StudyName()
            );
            string triple =
                principalInvestigatorSurname + "-" + taskNumber + "-" + StudyName();
            triple = triple.Replace("&", "and");
            triple = triple.Replace(' ', '_');
            triple = triple.Replace(',', '_');
            triple = triple.Replace("__", "_");

            if (triple.Length > MAX_STRING_LENGTH)
            {
                triple = triple.Substring(0, MAX_STRING_LENGTH);
            }

            return triple;
        }

        /// <summary>
        /// Main method: parses the Scope of Work, extracting project information, then:
        /// - creates local project directory
        /// - initializes output Excel file with desired disclaimer page
        /// - initializes stub of SQL file with project infomation
        /// - converts SlicerDicer code(if applicable)
        /// - drafts the completion email
        /// - pushes SQL file to GitLab
        /// </summary>
        /// <param name="doc"></param>
        internal void SetupProject(Document doc)
        {
            log.Debug("Setting up project.");
            scopeOfWork = doc;
            documentDirectoryName = Path.GetDirectoryName(doc.FullName);
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            string message;
            DialogResult result;

            if (progressForm.StopSignaled())
                return;

            // 1. Extract key information from Scope of Work.
            if (!Parse())
            {
                message = "Unable to parse document.";
                progressForm.ReportProgress(message);
                log.Error(message);
                result = MessageBox.Show(message, "Parse Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            if (progressForm.StopSignaled())
                return;

            // 2. Create project directory.
            if (!CreateProjectDirectory())
            {
                message = "Unable to create project directory.";
                progressForm.MarkFailedCreateProjectDirectory();
                progressForm.ReportProgress(message);
                log.Error(message);
                result = MessageBox.Show(message, "Create Directory Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            if (progressForm.StopSignaled())
                return;

            progressForm.CheckOffCreateProjectDirectory();
            progressForm.LinkProjectDirectory(documentDirectoryName);

            // 3. Create Excel file to hold output.
            if (!CreateOutputFile())
            {
                message = "Unable to create output file.";
                progressForm.MarkFailedInitializeExcelFile();
                progressForm.ReportProgress(message);
                log.Error(message);
                result = MessageBox.Show(message, "Create Output File Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            if (progressForm.StopSignaled())
                return;

            progressForm.CheckOffInitializeExcelFile();
            progressForm.LinkExcelFile(outputFileName);

            // 4. Initialize SQL file with project info in header.
            if (!BuildSqlFile())
            {
                message = "Unable to build SQL file.";
                progressForm.MarkFailedInitializeSqlFile();
                progressForm.ReportProgress(message);
                log.Error(message);
                result = MessageBox.Show(message, "Create SQL File Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            if (progressForm.StopSignaled())
                return;

            progressForm.CheckOffInitializeSqlFile();
            progressForm.LinkSqlFile(sqlFilename);

            // 5. Ask user how results will be delivered.
            DeliveryType deliveryType;

            using (var form = new DeliveryTypeForm())
            {
                result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    deliveryType = form.deliveryType;
                }
                else
                {
                    // User declined to specify, so we can't proceed.
                    message = "User did not specify the project delivery type.";
                    progressForm.ReportProgress(message);
                    log.Error(message);
                    return;
                }
            }

            if (progressForm.StopSignaled())
                return;

            // 6. Add taskNumber folder to Outlook.
            MsOutlook.Application app = new MsOutlook.Application();
            MsOutlook.Folder folder = app.Session.GetDefaultFolder(
            MsOutlook.OlDefaultFolders.olFolderInbox) as MsOutlook.Folder;
            MsOutlook.Folders folders = folder.Folders;
            MsOutlook.Folder decsFolder = null;

            try
            {
                decsFolder = (MsOutlook.Folder)folders.Add("DECS");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // That's OK--probably already exists.
            }

            // Now work in the DECS folder.
            try
            {
                decsFolder = folders["DECS"] as MsOutlook.Folder;

                if (decsFolder != null)
                {
                    try
                    {
                        decsFolder.Folders.Add(taskNumber);
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        log.Error(ex.Message);
                        result = MessageBox.Show(ex.Message, "Create folder Failed", buttons);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                log.Error(ex.Message);
                result = MessageBox.Show(ex.Message, "Create folder Failed", buttons);
            }

            // 7. Draft email reporting project completion.
            Emailer emailer = new Emailer(
                deliveryType: deliveryType,
                projectDirectory: projectDirectory.ToString(),
                requestorSalutation: Utilities.SalutationFromName(requesterName),
                taskNumber: taskNumber
            );

            if (
                !emailer.DraftOutlookEmail(
                    subject: "Your DECS Request is Ready: DECS-" + taskNumber,
                    recipients: requesterEmail
                )
            )
            {
                message = "Unable to draft email.";
                progressForm.MarkFailedDraftEmail();
                progressForm.ReportProgress(message);
                log.Error(message);
                result = MessageBox.Show(message, "Create email Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            progressForm.CheckOffDraftEmail();
            progressForm.LinkEmail(emailer);

            // 8. Push SQL file to GitLab.
            GitLabHandler gitLabHandler = new GitLabHandler();

            if (gitLabHandler.Ready())
            {
                if (!gitLabHandler.PushFileExe(sqlFilename))
                {
                    message = "Unable to upload SQL file to GitLab.";
                    progressForm.MarkFailedPushToGitLab();
                    progressForm.ReportProgress(message);
                    log.Error(message);
                    result = MessageBox.Show(message, "GitLab upload Failed", buttons);

                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
            }

            if (progressForm.StopSignaled())
                return;

            // Since it's a web address, use Uri class to convert path separators to fwd slash.
            Uri gitLabProjectAddress = new Uri(
                Path.Combine(GitLabHandler.Address(), projectTriple)
            );
            progressForm.CheckOffPushToGitLab();
            progressForm.LinkGitLab(gitLabProjectAddress.ToString());

            progressForm.EnableOkButton();
            progressForm.ReportProgress("Completed project " + taskNumber + " setup.");
        }

        /// <summary>
        /// Returns the study name, if available. If not, the dataset name.
        /// </summary>
        /// <returns>string</returns>
        private string StudyName()
        {
            if (string.IsNullOrEmpty(studyName))
            {
                return dataSetName;
            }

            return studyName;
        }

        /// <summary>
        /// Adds patient consent section to @c SlicerDicer SQL.
        /// </summary>
        /// <returns>bool</returns>
        private bool WriteConsentSection()
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(sqlFilename, append: true))
                {
                    writer.WriteLine("\n-- FINAL OUTPUT:");
                    writer.WriteLine("SELECT DISTINCT");
                    writer.WriteLine("    pid.IdentityId AS MRN");
                    writer.WriteLine("FROM #resultSet rs");
                    writer.WriteLine("JOIN dbo.PatientDim p\n    ON p.durableKey=rs.durableKey");
                    writer.WriteLine("-- Get MRN:");
                    writer.WriteLine(
                        "JOIN dbo.PatientIdentityDimX pid \n    ON pid.patientId=p.PatientEpicId AND pid.identityTypeId=2"
                    );
                    writer.WriteLine("-- Research-Eligble:");
                    writer.WriteLine(
                        "JOIN [prd-clarity].[clarity_prod].dbo.REGISTRY_DATA_INFO rdi"
                    );
                    writer.WriteLine("    ON rdi.NETWORKED_ID = p.PatientEpicId");
                    writer.WriteLine(
                        "JOIN [prd-clarity].[clarity_prod].dbo.REG_DATA_MEMBERSHP rdm"
                    );
                    writer.WriteLine(
                        "    ON rdm.RECORD_ID = rdi.RECORD_ID AND rdm.REGISTRY_ID = '100468'"
                    );
                    writer.WriteLine("WHERE");
                    writer.WriteLine("	p.PatientEpicId NOT IN");
                    writer.WriteLine(
                        "        (SELECT pat_id from \n[prd-clarity].[clarity_prod].ucsd_research.unconsented_patient)"
                    );
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
