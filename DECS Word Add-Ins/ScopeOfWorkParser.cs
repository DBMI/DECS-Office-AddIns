using DecsWordAddIns.Properties;
using log4net;
using log4net.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace DecsWordAddIns
{
    internal class ScopeOfWorkParser
    {
        private string dataSetName;
        private string dataSource;
        private string documentDirectoryName;
        private string principalInvestigatorEmail;
        private string principalInvestigatorGivenName;
        private string principalInvestigatorSurname;
        private DirectoryInfo projectDirectory;
        private string requesterEmail;
        private string requesterName;
        private Document scopeOfWork;
        private string sqlFilename;
        private string taskNumber;
        private string studyName;

        private const string DATA_SET_NAME_HEADING = "Data Set Name:";
        private const string DATA_SOURCE_HEADING = "Data Source:";
        private const int MAX_STRING_LENGTH = 32;
        private const string PEOPLE_PATTERN = @"(?<surname>[\w ]+), ?(?<given_name>[\w \.]+?) ?(?<email>[\d\w\.]+@[\d\w\.]+)";
        private const string PRINCIPAL_INVESTIGATOR_HEADING = "Principal Investigator (Name, E-mail):";
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

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        internal ScopeOfWorkParser()
        {
            LogManager.GetRepository().Threshold = Level.Debug;
            log.Debug("Instantiating ScopeOfWorkParser.");
            BuildRegex();
        }

        // Create all the reusable Regex objects.
        private void BuildRegex()
        {
            this.decsNumberRegex = new Regex(TASK_NUMBER_PATTERN);
            this.peopleRegex = new Regex(PEOPLE_PATTERN);
        }

        private bool BuildSqlFile()
        {
            this.sqlFilename = Path.Combine(this.projectDirectory.FullName, ProjectTriple() + ".sql");
            log.Debug("Will build file '" + this.sqlFilename + "'.");

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

            if (!CopySqlBody(slicerDicerFilename))
            {
                return false;
            }

            if (!WriteConsentSection())
            {
                return false;
            }

            return true;
        }

        private string Clean(string input)
        {
            return new string(input.Where(c => !char.IsControl(c)).ToArray()).Trim();
        }

        // Adapt the existing Slicer Dicer SQL code.
        private bool CopySqlBody(string slicerDicerFile)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(this.sqlFilename, append: true))
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

        private bool CreateOutputFile()
        {
            try
            {
                log.Debug("About to copy file to '" + projectDirectory.FullName + "'.");
                string targetFile = Path.Combine(projectDirectory.FullName, ProjectTriple() + ".xlsx");

                // Copy results template to project directory, allowing overwrite.
                File.Copy(@"Resources\results_template.xlsx", targetFile, true);
                return true;
            }
            catch (Exception ex)
            {
                log.Error("Unable to copy file to project directory because: " + ex.Message);
                return false;
            }
        }

        private bool CreateProjectDirectory()
        {
            string targetDirectory = Path.Combine(this.documentDirectoryName, ProjectTriple());
            this.projectDirectory = Directory.CreateDirectory(targetDirectory);
            return projectDirectory.Exists;
        }

        private bool Done()
        {
            bool haveStudyNameOrDataSetName = !string.IsNullOrEmpty(this.studyName) || !string.IsNullOrEmpty(this.dataSetName);
            return haveStudyNameOrDataSetName && 
                !string.IsNullOrEmpty(this.dataSource) &&
                !string.IsNullOrEmpty(this.principalInvestigatorEmail) &&
                !string.IsNullOrEmpty(this.principalInvestigatorGivenName) &&
                !string.IsNullOrEmpty(this.principalInvestigatorSurname) &&
                !string.IsNullOrEmpty(this.requesterEmail) &&
                !string.IsNullOrEmpty(this.requesterName) &&
                !string.IsNullOrEmpty(this.taskNumber);
        }

        private void ExtractDecsNumber(string text)
        {
            Match decsNumberMatch = this.decsNumberRegex.Match(text);

            if (decsNumberMatch.Success)
            {
                this.taskNumber = decsNumberMatch.Groups[1].Value.ToString().Trim();
            }
        }

        private void ExtractPI(string text)
        {
            Match piMatch = this.peopleRegex.Match(text);

            if (piMatch.Success)
            {
                this.principalInvestigatorGivenName = piMatch.Groups["given_name"].Value.ToString().Trim();
                this.principalInvestigatorSurname = piMatch.Groups["surname"].Value.ToString().Trim();
                this.principalInvestigatorEmail = piMatch.Groups["email"].Value.ToString().Trim();
            }
        }

        private void ExtractRequester(string text)
        {
            Match requesterMatch = this.peopleRegex.Match(text);

            if (requesterMatch.Success)
            {
                string requesterSurname = requesterMatch.Groups["surname"].Value.ToString().Trim();
                string requesterGivenName = requesterMatch.Groups["given_name"].Value.ToString().Trim();
                this.requesterName = requesterGivenName + " " + requesterSurname;
                this.requesterEmail = requesterMatch.Groups["email"].Value.ToString().Trim();
            }
        }

        private string GetNextLine(int desiredIndex)
        {
            string nextLine = String.Empty;

            if (desiredIndex >= 0 && desiredIndex < this.scopeOfWork.Paragraphs.Count)
            {
                nextLine = Clean(this.scopeOfWork.Paragraphs[desiredIndex].Range.Text.ToString());
            }

            return nextLine;
        }

        private string GetSlicerDicerFilename()
        {
            string filePath = string.Empty;
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            string message = "Is there a Slicer Dicer SQL file to be adapted?";
            DialogResult result = MessageBox.Show(message, "Slicer Dicer SQL File", buttons);

            if (result == DialogResult.Yes)
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

            return filePath;
        }

        private bool InsertSqlHeader() 
        {
            try
            {
                log.Debug("About to use Streamwriter to create SQL file.");

                using (StreamWriter writer = new StreamWriter(this.sqlFilename))
                {
                    writer.WriteLine("/*");
                    writer.WriteLine("** " + ProjectTriple() + ".sql");
                    writer.WriteLine("** Task: " + this.taskNumber);
                    writer.WriteLine("** Principal Investigator: " +
                        this.principalInvestigatorGivenName + " " +
                        this.principalInvestigatorSurname + ", " +
                        this.principalInvestigatorEmail);
                    writer.WriteLine("** Requester: " +
                        this.requesterName + ", " +
                        this.requesterEmail);
                    writer.WriteLine("** Author: " + Environment.UserName + ", " + UserPrincipal.Current.EmailAddress);
                    writer.WriteLine("** Created: " + DateTime.Now.ToString("yyyy-MM-dd"));
                    writer.WriteLine("** Database: " + this.dataSource);
                    writer.WriteLine("*/");
                    writer.WriteLine("");
                    writer.WriteLine("USE [" + this.dataSource.ToUpper() + "];");
                }
                
                return true;
            }
            catch (Exception ex)
            {
                log.Error("Unable to use StreamWriter to create SQL file because: " + ex.Message);
                return false;
            }
        }

        internal bool Parse()
        {
            int index = 1;
            string nextLine;

            while (!Done() && index < this.scopeOfWork.Paragraphs.Count)
            {
                Paragraph paragraph = this.scopeOfWork.Paragraphs[index];

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
                                this.dataSetName = nextLine;
                                break;

                            case DATA_SOURCE_HEADING:
                                index++;
                                nextLine = GetNextLine(index);
                                this.dataSource = nextLine;
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
                                this.studyName = nextLine;
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

            return Done();
        }

        private string ProjectTriple()
        {
            log.Debug("Building project triple from" + 
                      " surname: " + this.principalInvestigatorSurname + 
                      " task number: " + this.taskNumber + 
                      " study name: " + StudyName());
            string triple = this.principalInvestigatorSurname + "-" + this.taskNumber + "-" + StudyName();
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

        internal void SetupProject(Document doc)
        {
            log.Debug("Setting up project.");
            this.scopeOfWork = doc;
            this.documentDirectoryName = Path.GetDirectoryName(doc.FullName);
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            string message;
            DialogResult result;

            // 1. Extract key information from Scope of Work.
            if (!Parse())
            {
                message = "Unable to parse document.";
                log.Error(message);
                result = MessageBox.Show(message, "Parse Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            // 2. Create project directory.
            if (!CreateProjectDirectory())
            {
                message = "Unable to create project directory.";
                log.Error(message);
                result = MessageBox.Show(message, "Create Directory Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            // 3. Create Excel file to hold output.
            if (!CreateOutputFile())
            {
                message = "Unable to create output file.";
                log.Error(message);
                result = MessageBox.Show(message, "Create Output File Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            // 4. Initialize SQL file with project info in header.
            if (!BuildSqlFile()) 
            {
                message = "Unable to build SQL file.";
                log.Error(message);
                result = MessageBox.Show(message, "Create SQL File Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            // 5. Push SQL file to GitLab.
            GitLabHandler gitLabHandler = new GitLabHandler();

            if (gitLabHandler.Ready())
            {
                if (!gitLabHandler.PushFileExe(sqlFilename))
                {
                    message = "Unable to upload SQL file to GitLab.";
                    log.Error(message);
                    result = MessageBox.Show(message, "GitLab upload Failed", buttons);

                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
            }

            // 6. Ask user how results will be delivered.
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
                    log.Error("Usere did not specify the project delivery type.");
                    return;
                }
            }

            // 7. Draft email reporting project completion.
            Emailer emailer = new Emailer(deliveryType: deliveryType,
                                          projectDirectory: this.projectDirectory.ToString(),
                                          requestorSalutation: Utilities.SalutationFromName(this.requesterName),
                                          taskNumber: this.taskNumber);

            if (!emailer.DraftOutlookEmail(subject: "Your DECS Request is Ready: DECS-" + this.taskNumber,
                                           recipients: this.requesterEmail))
            {
                message = "Unable to draft email.";
                log.Error(message);
                result = MessageBox.Show(message, "Create email Failed", buttons);

                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            message = "Completed project " + this.taskNumber + " setup.";
            log.Debug(message);
            result = MessageBox.Show(message, "Success", buttons);

            if (result == DialogResult.OK)
            {
                return;
            }
        }

        private string StudyName()
        {
            if (string.IsNullOrEmpty(this.studyName))
            {
                return this.dataSetName;
            }

            return this.studyName;
        }

        private bool WriteConsentSection()
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(this.sqlFilename, append: true))
                {
                    writer.WriteLine("\n-- FINAL OUTPUT:");
                    writer.WriteLine("SELECT DISTINCT");
                    writer.WriteLine("    pid.IdentityId AS MRN");
                    writer.WriteLine("FROM #resultSet rs");
                    writer.WriteLine("JOIN dbo.PatientDim p\n    ON p.durableKey=rs.durableKey");
                    writer.WriteLine("-- Get MRN:");
                    writer.WriteLine("JOIN dbo.PatientIdentityDimX pid \n    ON pid.patientId=p.PatientEpicId AND pid.identityTypeId=2");
                    writer.WriteLine("-- Research-Eligble:");
                    writer.WriteLine("JOIN [prd-clarity].[clarity_prod].dbo.REGISTRY_DATA_INFO rdi");
                    writer.WriteLine("    ON rdi.NETWORKED_ID = p.PatientEpicId");
                    writer.WriteLine("JOIN [prd-clarity].[clarity_prod].dbo.REG_DATA_MEMBERSHP rdm");
                    writer.WriteLine("    ON rdm.RECORD_ID = rdi.RECORD_ID AND rdm.REGISTRY_ID = '100468'");
                    writer.WriteLine("WHERE");
                    writer.WriteLine("	p.PatientEpicId NOT IN");
                    writer.WriteLine("        (SELECT pat_id from \n[prd-clarity].[clarity_prod].ucsd_research.unconsented_patient)");
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