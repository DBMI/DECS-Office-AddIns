﻿using DecsWordAddIns.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsOutlook = Microsoft.Office.Interop.Outlook;

namespace DecsWordAddIns
{
    internal enum DeliveryType
    {
        OneDrive,
        VRD
    }

    // https://csharpexamples.com/c-send-an-email-using-outlook-program/
    internal class Emailer
    {
        private string htmlBody;

        private const string PROJECT_DIRECTORY = "{{ cookiecutter.__directory_name }}";
        private const string SALUTATION = "{{ cookiecutter.__requestor_salutation }}";
        private const string TASK_NUMBER = "{{ cookiecutter.task_number }}";

        private MsOutlook.MailItem mail;

        internal Emailer(
            DeliveryType deliveryType,
            string projectDirectory,
            string requestorSalutation,
            string taskNumber
        )
        {
            // Read in the boilerplate HTML & substitute actual values for the placeholders.
            ReadEmailBody(deliveryType);
            this.htmlBody = this.htmlBody.Replace(PROJECT_DIRECTORY, projectDirectory);
            this.htmlBody = this.htmlBody.Replace(SALUTATION, requestorSalutation);
            this.htmlBody = this.htmlBody.Replace(TASK_NUMBER, taskNumber);
        }

        internal bool DraftOutlookEmail(string subject, string recipients)
        {
            try
            {
                // create the outlook application.
                MsOutlook.Application outlookApp = new MsOutlook.Application();

                if (outlookApp == null)
                {
                    return false;
                }

                // create a new mail item.
                this.mail = (MsOutlook.MailItem)
                    outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);

                // add the body of the email
                this.mail.HTMLBody = this.htmlBody;

                this.mail.Subject = subject;
                this.mail.To = recipients;

                this.mail.Display(true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ReadEmailBody(DeliveryType deliveryType)
        {
            if (deliveryType == DeliveryType.OneDrive)
            {
                this.htmlBody = Resources.one_drive_email_body;
                return;
            }

            if (deliveryType == DeliveryType.VRD)
            {
                this.htmlBody = Resources.vrd_email_body;
                return;
            }
        }

        internal void ShowDraftEmail()
        {
            this.mail.Display(true);
        }

        internal string Subject()
        {
            return this.mail.Subject;
        }
    }
}