﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DecsWordAddIns
{
    /// <summary>
    /// Creates custom form to show user the progress of DECS project setup.
    /// </summary>
    internal partial class ProgressForm : Form
    {
        private const string CHECKED_BOX = "☑";
        private const string RED_X = "❌";

        private bool stopExecution = false;

        private Emailer emailer;

        internal ProgressForm()
        {
            InitializeComponent();
            ShowVersion();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.stopExecution = true;
            this.Close();
        }

        internal void CheckOffConvertSlicerDicer()
        {
            this.convertSlicerDicerStatusLabel.Text = CHECKED_BOX;
        }

        internal void CheckOffCreateProjectDirectory()
        {
            this.createProjectDirectoryStatusLabel.Text = CHECKED_BOX;
        }

        internal void CheckOffDraftEmail()
        {
            this.draftEmailStatusLabel.Text = CHECKED_BOX;
        }

        internal void CheckOffInitializeExcelFile()
        {
            this.initializeExcelFileStatusLabel.Text = CHECKED_BOX;
        }

        internal void CheckOffInitializeSqlFile()
        {
            this.initializeSqlFileStatusLabel.Text = CHECKED_BOX;
        }

        internal void CheckOffPushToGitLab()
        {
            this.pushToGitLabStatusLabel.Text = CHECKED_BOX;
        }

        private void draftEmailLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.emailer.ShowDraftEmail();
        }

        internal void EnableOkButton()
        {
            this.okButton.Enabled = true;
        }

        internal void EnableSlicerDicer()
        {
            this.convertSlicerDicerStatusLabel.Enabled = true;
            this.convertSlicerDicerDescriptionLabel.Enabled = true;
            this.convertSlicerDicerLinkLabel.Enabled = true;
        }

        internal void LinkConvertedSlicerDicerFile(string filePath)
        {
            this.convertSlicerDicerLinkLabel.Text = filePath;
        }

        internal void LinkEmail(Emailer emailer)
        {
            this.emailer = emailer;
            this.draftEmailLinkLabel.Text = this.emailer.Subject();
        }

        internal void LinkExcelFile(string filePath)
        {
            this.initializeExcelFileLinkLabel.Text = filePath;
        }

        internal void LinkGitLab(string webAddress)
        {
            this.pushToGitLabLinkLabel.Text = webAddress;
        }

        private void LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel linkLabel = (LinkLabel)sender;
            linkLabel.LinkVisited = true;
            System.Diagnostics.Process.Start(linkLabel.Text);
        }

        internal void LinkProjectDirectory(string directoryPath)
        {
            this.createProjectDirectoryLinkLabel.Text = directoryPath;
        }

        internal void LinkSqlFile(string filePath)
        {
            this.initializeSqlFileLinkLabel.Text = filePath;
        }

        internal void MarkFailedConvertSlicerDicer()
        {
            this.convertSlicerDicerStatusLabel.Text = RED_X;
        }

        internal void MarkFailedCreateProjectDirectory()
        {
            this.createProjectDirectoryStatusLabel.Text = RED_X;
        }

        internal void MarkFailedDraftEmail()
        {
            this.draftEmailStatusLabel.Text = RED_X;
        }

        internal void MarkFailedInitializeExcelFile()
        {
            this.initializeExcelFileStatusLabel.Text = RED_X;
        }

        internal void MarkFailedInitializeSqlFile()
        {
            this.initializeSqlFileStatusLabel.Text = RED_X;
        }

        internal void MarkFailedPushToGitLab()
        {
            this.pushToGitLabStatusLabel.Text = RED_X;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        internal void ReportProgress(string message)
        {
            this.progressLabel.Text = message;
        }

        private void ShowVersion()
        {
            System.Version version = System.Reflection.Assembly
                .GetExecutingAssembly()
                .GetName()
                .Version;
            string ver = String.Format(
                "{0}.{1}.{2}",
                version.Major,
                version.Minor,
                version.Revision
            );

            string filename = Assembly.GetExecutingAssembly().Location;
            FileInfo fi = new FileInfo(filename);
            DateTime modifiedDate = fi.LastWriteTime;
            this.versionLabel.Text = "ver. " + ver + " " + modifiedDate.ToString("yyyy-MM-dd");
        }

        internal bool StopSignaled()
        {
            return this.stopExecution;
        }
    }
}
