using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Net.Mail;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using SendGrid.Helpers.Mail;

//  12/21/2022 by Mark Gamber for Clipper Magazine
//  This is an Azure conversion of an old console
//  program run by a job scheduler

namespace EmailQTF
{
    public class Function1
    {
        [FunctionName("Function1")]
        public void Run([TimerTrigger("0 * * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation("Timer execution at " + DateTime.Now.ToString());
            List<Classes.EmailQueue> emails = GetSendableEmail(log);    //  Is there anything to be sent?
            if(emails.Count != 0)
            {
                log.LogInformation("There are " + emails.Count.ToString() + " to send");
                SendEmail(emails, log);                                 //  If so, send them
                MarkEmailSuccess(emails, log);                          //  And record the results in a data table
            }
        }

        // ============================================================================================================

        /// <summary>
        /// Mark emails as having been processed and the result from SendGrid
        /// </summary>
        /// <param name="emails">List of processed EmailQueue classes</param>
        public void MarkEmailSuccess(List<Classes.EmailQueue> emails, ILogger log)
        {
            try
            {
                string sConn = Environment.GetEnvironmentVariable("SQLConn");
                SqlConnection conn = new SqlConnection(sConn);
                conn.Open();                                            //  Open the database

                foreach (Classes.EmailQueue email in emails)            //  And for each email in the list
                {                                                       //  Update the IsProcessed, ProcessedDate and ProcessNotes columns
                    string sSQL = "update EmailQueueT set IsProcessed=1, ProcessedDate='" + DateTime.Now.ToString() + "', ProcessNotes='" +
                                  (email.Successful ? "SUCCESS" : "ERROR") + "' where EmailQueueID=" + email.EmailQueueID.ToString();
                    SqlCommand cmd = new SqlCommand(sSQL, conn);
                    cmd.ExecuteNonQuery();                              //  Execute it and move to the next email
                    cmd.Dispose();
                }
                conn.Close();
            }
            catch(Exception e)
            {
                log.LogError("Error opening EmailQueueT: " + e.ToString());
            }
        }

        // ============================================================================================================

        /// <summary>
        /// Send any email in the EmailQueue list
        /// </summary>
        /// <param name="emails">List of EmailQueue items to be processed</param>
        public async void SendEmail(List<Classes.EmailQueue> emails, ILogger log)
        {                                                               //  Get the SendGrid password and open a connection
            string sPassword = Environment.GetEnvironmentVariable("Password");
            SendGrid.SendGridClient EmailMessage = new SendGrid.SendGridClient(sPassword);
                                                                        //  Get test mode flag and email address
            string sTestMode = Environment.GetEnvironmentVariable("TestMode");
            string sTestEmail = Environment.GetEnvironmentVariable("TestEmail");
            foreach (Classes.EmailQueue email in emails)                //  Loop through all email in the list
            {
                List<EmailAddress> addresses = new List<EmailAddress>();
                List<EmailAddress> CCs = new List<EmailAddress>();      //  If this is a test setup, use the test email
                if (!string.IsNullOrEmpty(sTestMode) && sTestMode.ToLower().Equals("true"))
                    addresses.Add(new EmailAddress(sTestEmail));
                else
                {
                    if (email.ToEmail.Contains(','))                    //  Otherwise, if there's a comma...
                    {
                        string[] sAddrs = email.ToEmail.Split('.');     //  Do a split on and make them all CC addresses
                        foreach (string sAddr in sAddrs)
                            CCs.Add(new EmailAddress(sAddr));
                    }
                    else                                                //  Otherwise make it the recipient
                        addresses.Add(new EmailAddress(email.ToEmail));
                }
                                                                        //  Create the email message
                EmailAddress from = new EmailAddress(email.FromEmail, email.FromName);
                SendGridMessage msg = MailHelper.CreateSingleEmailToMultipleRecipients(from, addresses, email.Subject, email.Body, email.Body);
                if (CCs.Count != 0)                                     //  Add any CC addresses from above
                    msg.AddCcs(CCs);

                if (!string.IsNullOrEmpty(email.CCEmail))               //  Add CC email if there is one
                    msg.AddCc(new EmailAddress(email.CCEmail));
                if (!string.IsNullOrEmpty(email.BCCEmail))              //  Add BCC email if there is one
                    msg.AddBcc(new EmailAddress(email.BCCEmail));

                foreach (var attachment in email.Attachments)           //  Loop through all attachments
                {                                                       //  Create full filename and then full path to the file
                    var fileName = attachment.FileName + "." + attachment.FileExtension;
                    var filePath = "\\\\" + attachment.ServerName +
                        (!attachment.FileDirectory.StartsWith("\\") ? "\\" : string.Empty) + attachment.FileDirectory +
                        "\\" + fileName;
                                                                        //  Read the file and turn into a Base64 string
                    byte[] bAttachment = System.IO.File.ReadAllBytes(filePath);
                    string sBase64Attachment = Convert.ToBase64String(bAttachment);
                    var ament = new SendGrid.Helpers.Mail.Attachment(); //  Create a SendGrid attachment
                    ament.Type = "application/pdf";                     //  Tell it it's a PDF attachment
                    ament.Disposition = "attachment";
                    ament.Filename = fileName;
                    ament.Content = sBase64Attachment;                  //  Provide Base64 data
                    msg.AddAttachment(ament);                           //  And add the attachment to the message
                }
                try
                {
                    var response = await EmailMessage.SendEmailAsync(msg);  //  Send the email to SendGrid
                    email.Successful = response.IsSuccessStatusCode;    //  And save the status code for later
                }
                catch(Exception e)
                {
                    log.LogError("Error sending email: " + e.ToString());
                }
            }
        }

        // ============================================================================================================

        /// <summary>
        /// Find any email that needs to be sent and return as a List of EmailQueue items
        /// </summary>
        /// <returns>List of EmailQueue items, each item being a single email</returns>
        public List<Classes.EmailQueue> GetSendableEmail(ILogger log)
        {
            List<Classes.EmailQueue> queue = new List<Classes.EmailQueue>();
            try
            {
                string sConn = Environment.GetEnvironmentVariable("SQLConn");
                SqlConnection conn = new SqlConnection(sConn);          //   Open the database and build a query
                conn.Open();
                string sSQL = "select EmailQueueID, FromEmail, FromName, ToEmail, [Subject], IsHTML, Body, SendDate, CCEmail, BCCEmail " +
                              "from EmailQueueT " +
                              "where SendDate <= '" + DateTime.Now.ToString() + "' and IsProcessed=0";
                SqlCommand cmd = new SqlCommand(sSQL, conn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())                                   //  Loop through the query results
                {
                    Classes.EmailQueue item = new Classes.EmailQueue(); //  Create an EmailQueue item for each row
                    item.EmailQueueID = reader.GetInt32(0);
                    item.FromEmail = reader.GetString(1);
                    item.FromName = reader.GetString(2);
                    item.ToEmail = reader.GetString(3);
                    item.Subject = reader.GetString(4);
                    item.IsHTML = reader.GetBoolean(5);
                    item.Body = reader.GetString(6);
                    item.SendDate = reader.GetDateTime(7);
                    item.CCEmail = reader.IsDBNull(8) ? null : reader.GetString(8);
                    item.BCCEmail = reader.IsDBNull(9) ? null : reader.GetString(9);
                    item.Successful = false;
                    queue.Add(item);                                    //  Add the item to the list of items
                }

                foreach (Classes.EmailQueue item in queue)
                    item.Attachments = GetEmailAttachments(conn, item.EmailQueueID, log);

                conn.Close();
            }
            catch(Exception e)
            {
                log.LogError("Error opening EmailQueueT: " + e.ToString());
            }
            return queue;                                               //  Return the list of email items
        }

        // ============================================================================================================

        /// <summary>
        /// Retrieve any attachments for a given EmailQueueID
        /// </summary>
        /// <param name="conn">Open SqlConnection object</param>
        /// <param name="EmailQueueID">Email ID to be checked for attachments</param>
        /// <returns>List of EmailQueueAttachment items, each item being one attachment</returns>
        public List<Classes.EmailQueueAttachment> GetEmailAttachments(SqlConnection conn, int EmailQueueID, ILogger log)
        {
            List<Classes.EmailQueueAttachment> attachments = new List<Classes.EmailQueueAttachment>();
            try
            {                                                           //  Create a SQL query for the Queue ID                                                                       
                string sSQL = "select eqt.EmailQueueAttachmentID, eqt.ApplicationID, eqt.ApplicationServerID, eqt.EmailQueueAttachmentTypeID, " +
                              "eqt.FileDirectory, eqt.FileName, eqt.DateCreated, ast.ServerName, eqat.FileExtension " +
                              "from EmailQueueAttachmentT eqt " +
                              "inner join ApplicationServerT ast on ast.ApplicationServerID=eqt.ApplicationServerID " +
                              "inner join EmailQueueAttachmentTypeT eqat on eqat.EmailQueueAttachmentTypeID=eqt.EmailQueueAttachmentTypeID " +
                              "where eqt.EmailQueueID=" + EmailQueueID.ToString();
                SqlCommand cmd = new SqlCommand(sSQL, conn);
                SqlDataReader reader = cmd.ExecuteReader();             //  Execute the query
                while (reader.Read())                                   //  Read and save each row of data
                {
                    Classes.EmailQueueAttachment attachment = new Classes.EmailQueueAttachment();
                    attachment.EmailQueueAttachmentID = (int)reader.GetInt64(0);
                    attachment.EmailQueueID = EmailQueueID;
                    attachment.ApplicationID = reader.GetInt32(1);
                    attachment.ApplicationServerID = reader.GetInt32(2);
                    attachment.EmailQueueAttachmentTypeID = reader.GetInt32(3);
                    attachment.FileDirectory = reader.GetString(4);
                    attachment.FileName = reader.GetString(5);
                    attachment.DateCreated = reader.GetDateTime(6);
                    attachment.ServerName = reader.GetString(7);
                    attachment.FileExtension = reader.GetString(8);
                    attachments.Add(attachment);
                }
            }
            catch(Exception e)
            {
                log.LogError("Error opening EmailQueueAttachmentT: " + e.ToString());
            }
            return attachments;                                         //  Return the list of attachments
        }
    }
}
