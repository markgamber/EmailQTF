using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailQTF.Classes
{
    public class EmailQueue
    {
        public int EmailQueueID { get; set; }
        public string FromEmail { get; set; }
        public string FromName { get; set; }
        public string ToEmail { get; set; }
        public string Subject { get; set; }
        public bool IsHTML { get; set; }
        public string Body { get; set; }
        public DateTime SendDate { get; set; }
        public string CCEmail { get; set; }
        public string BCCEmail { get; set; }
        public bool Successful { get; set; }
        public List<EmailQueueAttachment> Attachments { get; set; }
    }

    public class EmailQueueAttachment
    {
        public int EmailQueueAttachmentID { get; set; }
        public int EmailQueueID { get; set; }
        public int ApplicationID { get; set; }
        public int ApplicationServerID { get; set; }
        public int EmailQueueAttachmentTypeID { get; set; }
        public string FileDirectory { get; set; }
        public string FileName { get; set; }
        public DateTime DateCreated { get; set; }
        public string ServerName { get; set; }
        public string FileExtension { get; set; }
    }
}
