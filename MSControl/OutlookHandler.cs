﻿using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Outlook, can send emails with multiple attachments.
    /// </summary>
    public class OutlookHandler
    {
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, string recipient, int importance=1)
        {
            sendMail(subject, body, new List<string>() { recipient }, new List<string>(), importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, List<string> recipients, int importance=1)
        {
            sendMail(subject, body, recipients, new List<string>(), importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipeint">Single recipient of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, string recipeint, string attachmentPath, int importance=1)
        {
            sendMail(subject, body, new List<string>() { recipeint }, new List<string>() { attachmentPath }, importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, List<string> recipients, string attachmentPath, int importance=1)
        {
            sendMail(subject, body, recipients, new List<string>() { attachmentPath }, importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email.</param>
        /// <param name="attachmentPaths">List of attachment paths of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, string recipient, List<string> attachmentPaths, int importance=1)
        {
            sendMail(subject, body, new List<string>() { recipient }, attachmentPaths, importance);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPaths">List of the attachment paths of the email.</param>
        /// <param name="importance">The importance level of the email. Can be low (0), normal (1) or high (2).</param>
        public void sendMail(string subject, string body, List<string> recipients, List<string> attachmentPaths, int importance=1)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
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

        }
    }
}
