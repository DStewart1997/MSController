using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Outlook, can send emails with multiple attachments.
    /// </summary>
    public class OutlookHandler
    {
        Outlook.Application app = null;
        Outlook.MAPIFolder root = null;
        Outlook.MailItem mailItem = null;

        /// <summary>
        /// Constructor for creating initial COM objects.
        /// </summary>
        public OutlookHandler()
        {
            app = new Outlook.Application();
            root = app.Session.DefaultStore.GetRootFolder();
            mailItem = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
        }

        // SendMail
        #region
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, string recipient, int importance=1)
        {
            SendMail(subject, body, new List<string>() { recipient }, new List<string>(), importance);
        }
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, List<string> recipients, int importance=1)
        {
            SendMail(subject, body, recipients, new List<string>(), importance);
        }
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipeint">Single recipient of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, string recipeint, string attachmentPath, int importance=1)
        {
            SendMail(subject, body, new List<string>() { recipeint }, new List<string>() { attachmentPath }, importance);
        }
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, List<string> recipients, string attachmentPath, int importance=1)
        {
            SendMail(subject, body, recipients, new List<string>() { attachmentPath }, importance);
        }
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email.</param>
        /// <param name="attachmentPaths">List of attachment paths of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, string recipient, List<string> attachmentPaths, int importance=1)
        {
            SendMail(subject, body, new List<string>() { recipient }, attachmentPaths, importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPaths">List of the attachment paths of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void SendMail(string subject, string body, List<string> recipients, List<string> attachmentPaths, int importance=1)
        {
            try
            {
                mailItem.Subject = subject;
                mailItem.Body = body;
                mailItem.To = string.Join("; ", recipients.ToArray());

                foreach (string attachment in attachmentPaths)
                    mailItem.Attachments.Add(attachment);

                switch (importance)
                {
                    case 0:
                        mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                        break;
                    case 1:
                        mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
                        break;
                    case 2:
                        mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                        break;
                }
                
                mailItem.Display(false);

                mailItem.Send();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseCOMObjects();
            }

        }
        #endregion

        /// <summary>
        /// Returns all folders from current session (Including subfolders).
        /// </summary>
        /// <param name="root"></param>
        public void GetFolders(Outlook.MAPIFolder root = null)
        {
            if (root == null)
                root = this.root;

            foreach (Outlook.MAPIFolder folder in root.Folders)
            {
                Console.WriteLine(folder.Name);
                GetFolders(folder);
            }
        }


        /// <summary>
        /// Releases all used COM objects, useful for when try, catch, finally blocks to ensure all COM objects are released.
        /// </summary>
        public void ReleaseCOMObjects()
        {
            if (app != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            if (mailItem != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);

            // The internet said do this and it works so it's here
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
