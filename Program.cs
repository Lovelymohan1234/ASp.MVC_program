using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace BC_Daily_Budget_Utilization_Report
{
    class Program
    {
        public static string schedulerstart, mailtriggerattemp;
        static void Main(string[] args)
        {
            DateTime now = DateTime.Now;
            schedulerstart = now.ToString("T");
            Process aProcess = Process.GetCurrentProcess();
            string aProcName = aProcess.ProcessName;
            if (Process.GetProcessesByName(aProcName).Length > 1)
            {
                Log("System is all ready running..!!!");
                return;
            }
            // Console.WriteLine(DateTime.Now.ToString());
            mailtriggerattemp = "1"; //CH03
            Get_Dailybudgetutilization();
        }
        private static void Get_Dailybudgetutilization()
        {
            bool isError = false;
            try
            {

                SqlConnection sqlConnection;
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["SnD"].ConnectionString);
                SqlCommand command = new SqlCommand("BC_Daily_Budget_Utilization_Rpt", sqlConnection);
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 800000000;
                //command.Parameters.Add("@Id", SqlDbType.VarChar).Value = txtId.Text;
                //command.Parameters.Add("@Name", SqlDbType.DateTime).Value = txtName.Text;
                sqlConnection.Open();
                command.ExecuteNonQuery();
                sqlConnection.Close();

                //*Ch01
                DateTime currentTime = DateTime.Now;
                DateTime x30MinsLater = currentTime.AddMinutes(30);
                DateTime x4hourLater = x30MinsLater.AddHours(4);
                DateTime SLTime = x4hourLater;

                //*Ch01
                string s1Year = SLTime.Year.ToString();
                string s1Month = SLTime.Month.ToString().PadLeft(2, '0');
                string s1Day = SLTime.Day.ToString();
                string s1ErrorTime = s1Day + "-" + s1Month + "-" + s1Year;
                DateTime dt = DateTime.Now;
                string from = System.Configuration.ConfigurationSettings.AppSettings["Mail_From"].ToString(),
                                        to = System.Configuration.ConfigurationSettings.AppSettings["Mail_To"].ToString(),
                                        copy = System.Configuration.ConfigurationSettings.AppSettings["Mail_Copy"].ToString(),

                                    subject = "Daily_Budget_Utilization_Report" + ": " + s1ErrorTime, body, filePath;
                bool isHtmlBody;
                string server = "";
                if (mailtriggerattemp == "1")
                    server = System.Configuration.ConfigurationSettings.AppSettings["MailServer"].ToString();
                else if (mailtriggerattemp == "2")
                    server = "130.24.108.73";
                else if (mailtriggerattemp == "3")
                    server = "130.24.104.70";
                //*CH03
                int mailPort = 25;
                SqlConnection cnn;
                string connectionString = null;
                string sql = null, sqlEX = null;
                connectionString = ConfigurationManager.ConnectionStrings["SnD"].ConnectionString;
                cnn = new SqlConnection(connectionString);
                cnn.Open();
                sql = "select * from BC_Daily_Budget_Utilization_report";
                SqlDataAdapter dscmd = new SqlDataAdapter(sql, cnn);
                DataTable ds_mail = new DataTable();
                dscmd.Fill(ds_mail);
                //string FileName = decimal.Parse(GetDataTable(sql).Rows[0]["SELL_FACTOR1"].ToString());
                string filename = ds_mail.Rows[0].Field<string>(0) + ".csv";
                string returnMsg = "";
                MailMessage mailMsg = new MailMessage();

                //set from Address
                MailAddress mailAddress = new MailAddress(from);
                mailMsg.From = mailAddress;
                //set to Adress
                mailMsg.To.Add(to);
                // Set Message subject
                mailMsg.Subject = subject;
                //set mail cc
                if (copy != "")
                    mailMsg.CC.Add(copy);
                // Set Message Body
                mailMsg.IsBodyHtml = true;
                filePath = System.Configuration.ConfigurationSettings.AppSettings["Attachment_path"].ToString();
                DirectoryInfo dir = new DirectoryInfo(filePath);

                dir.Refresh();
                filePath = filePath + filename;
                if (filePath != "")
                {
                    Attachment attach3 = new Attachment(filePath, "application/vnd.ms-excel");
                    mailMsg.Attachments.Add(attach3);
                }
                string v = subject.Replace("NGDMS Interface - ", "");
                string v1 = v.Replace(" Error", "");
                string htmlBody, htmlBody2, htmlBody3, htmlBody4;
                htmlBody = "<html>Dear Team,<br/><br/></html>";

                htmlBody2 = "<html>Please find the attachment of Daily Budget Utilization Report generated on " + s1Day + "/" + s1Month + "/" + s1Year + ".<br/><br/></html>";
                htmlBody4 = "<html><br/><br/>Thanks & Regards,<br/></html>CSDP Interface Support Team.</html>";
                //<br/><br/>The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain proprietary, confidential or privileged information. <br/>If you are not the intended recipient, you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately and destroy all copies of this message and any attachments.<br/></html>";
                mailMsg.Body = htmlBody + "  " + htmlBody2 + htmlBody4;


                //set message body format
                //mailMsg.IsBodyHtml = isHtmlBody;
                // System.IO.File.WriteAllText(@"c:\abc.xlsx", attachmentStream);
                //byte[] data = GetData(attachmentStream);


                //if (attachmentStream != null)//Define mail attachment.
                //{
                //    ms.Position = 0;//explicitly set the starting position of the MemoryStream
                //Attachment attach = new Attachment(filePath, "application/vnd.ms-excel");
                //mailMsg.Attachments.Add(attach);
                //}


                //set exchange server
                SmtpClient smtpClient = new SmtpClient(server);
                smtpClient.Send(mailMsg);
                mailMsg.Dispose();
                //return returnMsg;


            }
            catch (Exception ex)
            {
                DateTime date = new DateTime();
                Console.WriteLine("SQL Error" + ex.Message.ToString());

                LogException(ex);
                //return 0;
                //CH03
                if (mailtriggerattemp == "1")
                    mailtriggerattemp = "2";
                else if (mailtriggerattemp == "2")
                    mailtriggerattemp = "3";
                isError = true;


            }
            if (isError) Get_Dailybudgetutilization();


        }
        public static void Log(string message)
        {
            StreamWriter streamWriter = null;

            try
            {
                string sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";
                string sPathName = AppDomain.CurrentDomain.BaseDirectory + "\\Dailybudgetutilization";
                string sYear = DateTime.Now.Year.ToString();
                string sMonth = DateTime.Now.Month.ToString();
                string sDay = DateTime.Now.Day.ToString();
                string sErrorTime = sDay + "-" + sMonth + "-" + sYear;
                streamWriter = new StreamWriter(sPathName + sErrorTime + ".txt", true);
                streamWriter.WriteLine(sLogFormat + message);
                streamWriter.Flush();

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                //Console.Read();
            }
            finally
            {
                if (streamWriter != null)
                {
                    streamWriter.Dispose();
                    streamWriter.Close();
                }
            }
        }
        public static void LogException(Exception exception)
        {
            String currenttime = DateTime.Now.ToString();
            try
            {

                string mailSubject = "Daily_Budget_Utilization_Report Exception";
                string mailBody = "Following run time exception threw while running Daily_Budget_Utilization_Report on " + currenttime + ".";
                mailBody += "\n\r";
                mailBody += exception.Message;
                if (exception.StackTrace != null)
                {
                    mailBody += "\n\r";
                    mailBody += "Exception StackTrace As Follows.";
                    mailBody += "\n\r";
                    mailBody += exception.StackTrace;
                }
                mailBody += "\n\r";
                mailBody += "\n\r";
                mailBody += "This is a system generated email.";

                char seperator = Convert.ToChar(System.Configuration.ConfigurationSettings.AppSettings["Seperator"]);
                string toaddress = System.Configuration.ConfigurationSettings.AppSettings["Exceotion_Mail_To"].ToString();
                string copyaddress = System.Configuration.ConfigurationSettings.AppSettings["Exceotion_Mail_Copy"].ToString();
                string mailServer = System.Configuration.ConfigurationSettings.AppSettings["MailServer"].ToString();
                int mailPort = Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["MailPort"]);
                string mailFrom = System.Configuration.ConfigurationSettings.AppSettings["Exceotion_Mail_From"].ToString();
                Send(mailFrom, toaddress.Split(seperator), copyaddress.Split(seperator), mailSubject, mailBody, false, mailServer, mailPort);

            }
            catch (Exception innerException)
            {
                string sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";
                string sPathName = AppDomain.CurrentDomain.BaseDirectory + "\\Daily_Budget_Utilization_Report";
                string sYear = DateTime.Now.Year.ToString();
                string sMonth = DateTime.Now.Month.ToString();
                string sDay = DateTime.Now.Day.ToString();
                string sErrorTime = sDay + "-" + sMonth + "-" + sYear;
                StreamWriter streamWriter = new StreamWriter(sPathName + sErrorTime + ".txt", true);
                streamWriter.WriteLine("Following run time exception threw while running " + currenttime + ".");
                streamWriter.WriteLine(exception.Message);
                if (exception.StackTrace != null)
                {
                    streamWriter.WriteLine(exception.StackTrace);
                }
                streamWriter.WriteLine(" ");
                streamWriter.WriteLine(" ");
                streamWriter.WriteLine(" ");
                streamWriter.WriteLine("==========local exception======");
                streamWriter.WriteLine(innerException.Message);
                streamWriter.Flush();
            }
        }
        public static string Send(string from, string[] to, string[] copy, string subject, string body, bool isHtmlBody, string server, int mailPort)
        {
            string returnMsg = "";
            MailMessage mailMsg = new MailMessage();
            try
            {
                // Set Message subject
                mailMsg.Subject = subject;

                //set from Address
                MailAddress mailAddress = new MailAddress(from);
                mailMsg.From = mailAddress;

                //set to Adress
                foreach (string address in to)
                {
                    if (address != "")
                        mailMsg.To.Add(address);
                }

                //set mail cc
                foreach (string address in to)
                {
                    if (address != "")
                        mailMsg.CC.Add(address);
                }

                // Set Message Body
                mailMsg.Body = body;
                mailMsg.IsBodyHtml = isHtmlBody;


                //set exchange server
                SmtpClient smtpClient = new SmtpClient(server, mailPort);
                //send mail
                smtpClient.Send(mailMsg);
                returnMsg = "sent";

            }
            catch (Exception exception)
            {
                returnMsg = exception.Message.ToString();
            }

            return returnMsg;
        }

    }




}
