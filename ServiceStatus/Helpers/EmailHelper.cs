using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace ServiceStatus
{
    public class EmailHelper
    {
        // public void SendEmail(string[] fileContent)
        public void SendEmail(StringBuilder htmlTable)
        {
            #region Email SMTP setup   
            FileHelper fileHelper = new FileHelper();
            var smtpServer = new SmtpClient("smtp-mail.outlook.com");
            smtpServer.Port = 587;
            smtpServer.Credentials = new NetworkCredential("neelkanth03@outlook.com", "Pushpa#86");
            smtpServer.EnableSsl = true;

            var mail = new MailMessage();
            mail.From = new MailAddress("neelkanth03@outlook.com", "EMR Test Automation Team");
            mail.Subject = "Running Services Report";

            #region Getting e-mail list from excel file

            var toEmails = ExcelHelper.xlReadSpecificColoumn("Emails", 3);
            var ccEmail = ExcelHelper.xlReadSpecificColoumn("Emails", 4);
            if (toEmails.Count != 0)
            {
                foreach (var emails in toEmails) { mail.To.Add(new MailAddress(emails)); }
            }
            else
                mail.To.Add(new MailAddress("sumitt77@in.ibm.com"));

            if (ccEmail.Count != 0)
            {
                foreach (var emails in ccEmail) { mail.CC.Add(emails); }
            }
            #endregion

            mail.IsBodyHtml = true;

            string htmlBody = @"<p>Hi all,
                        </br>
                        <br>Please find herewith the Service Running Status as on <span style=background-color:yellow>" + DateTime.Now + "</span> as below: </br></br>"
                        + htmlTable.ToString() +
                        @"<p><br><b>Kind Regards,</b></br>
                        <br><font size='+1'>EMR Test Automation Team</font></br>
                        <br>International Business Machine (IBM)</br>
                        Hinjewadi Phase 2, Pune, INDIA</br></p>";
            mail.Body = htmlBody;
            htmlBody+= "<p>=================================================================================================</p>";
            #endregion
            try
            {
                string resultDir = $@"{fileHelper.GetProjectPath()}\Result\";
                string readContent = string.Empty;
                if (!Directory.Exists(resultDir))
                    Directory.CreateDirectory(resultDir);
                else
                    readContent = File.ReadAllText($@"{resultDir}\result.html");

                File.WriteAllText($@"{resultDir}\result.html", htmlBody + readContent);
            }
            catch (Exception e) { Console.WriteLine(e); }
            #region Email Send
            try
            {
                smtpServer.Send(mail);
                // test.Log(LogStatus.Pass, "Email Send To: " + mail.To + "</br>And as Cc to " + mail.CC);
            }
            catch (SmtpException ex)
            {
                if (ex.StatusCode == SmtpStatusCode.InsufficientStorage)
                {
                    // test.Log(LogStatus.Error, ex.ToString()); test.Log(LogStatus.Info, "Trying to send email again");
                    try
                    {
                        //Send again to ensure this email gets sent
                        smtpServer.Send(mail);
                        //    test.Log(LogStatus.Pass, "Message Send to recipients");
                    }
                    catch (Exception e)
                    {
                        var errorMsg = e.Message;
                        //Email sending failed while trying again
                        //   test.Log(LogStatus.Fail, "Email sending failed!" + "</br>" + ex);
                    }
                }
                else
                {
                    //Handle other SMTP errors here. 
                    // test.Log(LogStatus.Fail, "Email sending failed!" + "</br>" + ex);
                }
                //  }
            }
            #endregion
        }
    }
}
