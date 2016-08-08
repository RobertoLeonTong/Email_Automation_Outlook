using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IniParser;
using System.Data.SqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailOutlook
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Program n = new Program();
                StringBuilder mailbody = new StringBuilder();
                string Month = DateTime.Now.ToString("MMM");
                //string Day = DateTime.Now.ToString("dd");
                string Day0 = DateTime.Today.ToString("d yyyy");
                int indx = Day0.IndexOf(" ");
                string Dayz = Day0.Substring(0, indx);
                string Year = DateTime.Now.ToString("yyyy");

                string[] Traders = new string[100];
                int Tracker = 0; //Tracks the amounts of emails

                IniParser.FileIniDataParser parser = new FileIniDataParser();
                parser.CommentDelimiter = '#';
                IniData data = parser.LoadFile("config.ini");

                //Extraction of variables/settings from config.ini
                string DBCredentials = data["Config"]["DBCredentials"];
                string EODClients = data["Config"]["table1"];
                string Trades = data["Config"]["table2"];
                string EODEmails = data["Config"]["table3"];
                string Subject = data["Email"]["subject"];
                string MailServer = data["Email"]["mailServer"];
                string From = data["Email"]["from"];

                //Start of program
                Console.WriteLine("-------------------------------------------------");
                Console.WriteLine("Program Start\n");
                Console.WriteLine("Table1: " + EODClients);
                Console.WriteLine("Table2: " + Trades);
                Console.WriteLine("Table3: " + EODEmails);
                Console.WriteLine("\nSubject: " + Subject + "\n");

                //Header
                mailbody.Append("<head>" +
                                 "<style>" +
                                 "th {background-color: #E3290F; font-size: 1em; color: white;font-family:sans-serif;} " +
                                 "table {background: white;width:700px}" +
                                 "table, td, th {border: 2px solid black;border-collapse:collapse;border-width:thin;padding: 1px 5px 1px 5px;font-size:1em;}" +
                                 "th { text-align: left;}" +
                                 "</style></head><body>" +
                                 "<div style='width:500px;color:black;font-family:Arial;text-align:center;'> " +
                                 "<div style='text-align:center;text-decoration:underline;'><h3><em>Daily Clearing Report</em></h3></div>" +
                                 "<h3 style='font-size:0.90em;font-family:Arial;text-align:center;'>Trader Notifications</h3>" +
                                 "<table>" +
                                 "<tr>" +
                                 "<th style='width:100px'><em>Client</em></th>" +
                                 "<th><em>UserID</em></th>" +
                                 "</tr>");


                using (SqlConnection conn = new SqlConnection(DBCredentials))
                {
                    conn.Open(); //Opens Connection

                    SqlCommand command = conn.CreateCommand(); //Creates SQL Query

                    command.CommandText = "Select DISTINCT userId, Client, Description " +
                                           "FROM " + EODClients + " " +
                                           "WHERE userId in (Select DISTINCT userID from " + Trades + " Where tradeDate like '%" + Month + "%" + Dayz +  "%" + Year  + "%')" +
                                           "AND Notification Like '%Trader%'" +
                                           "OR userId in (Select DISTINCT userID from " + Trades + " Where tradeDate like '%"  + Month + "%" + Dayz + "%" + Year + "%')" +
                                           "AND Notification Like '%Both%'";

                    using (SqlDataReader oReader = command.ExecuteReader())  //Reads in first query
                    {
                        while (oReader.Read())
                        {
                            mailbody.AppendLine("<tr><td>" + oReader["Client"].ToString().Replace(" ", string.Empty) + "</td><td>" + oReader["userId"].ToString().Replace(" ", string.Empty) + "</td></tr>");
                        }

                    }

                    //Second Table
                    mailbody.Append("</table><br><br><div style='text-align:center'><h3 style='font-size:0.9em;font-family:Arial'>Co-op EOD Processing</h3></div>" +
                                    "<table>" +
                                    "<tr>" +
                                    "<th style='width:100px'><em>Client</em></th>" +
                                    "<th style='width:100px'><em>UserID</em></th>" +
                                    "<th ><em>Description</em></th>" +
                                    "</tr>");

                    SqlCommand command2 = conn.CreateCommand(); //Creates second SQL query

                    command2.CommandText = "Select DISTINCT userId, Client, Description " +
                                           "FROM " + EODClients + " " +
                                           "WHERE userId in (Select DISTINCT userID from " + Trades + " WHERE tradeDate like '%"  + Month + "%" + Dayz + "%" + Year + "%')" +
                                           "AND Notification Like '%Co-op%'" +
                                           "OR userId in (Select DISTINCT userID from " + Trades + " Where tradeDate like '%" + Month + "%" + Dayz + "%" + Year + "%')" +
                                           "AND Notification Like '%Both%'";

                    using (SqlDataReader oReader = command2.ExecuteReader()) //Reads information from second query
                    {
                        while (oReader.Read())
                        {
                            mailbody.AppendLine("<tr><td>" + oReader["Client"].ToString().Replace(" ", string.Empty) + "</td><td>" + oReader["userId"].ToString().Replace(" ", string.Empty) + "</td><td>" + oReader["Description"].ToString().Trim() + "</td></tr>");
                        }

                    }

                    SqlCommand command1 = conn.CreateCommand(); //Third SQL Query

                    command1.CommandText = "SELECT email FROM " + EODEmails; //Obtains Emails

                    

                    using (SqlDataReader ReadX = command1.ExecuteReader())  //Reads in information from third SQL query
                    {
                        while (ReadX.Read())
                        {
                            Traders[Tracker] = ReadX["email"].ToString(); //Enters Emails
                            Tracker++;
                        }
                    }

                    //Testing Emails
                    for (int i = 0; i < Traders.Length; i++)
                    {
                        if (Traders[i] == null) break;
                            Console.WriteLine("Email " + i + ": " + Traders[i]); //Testing for array content
                    }

                    conn.Close(); //Closes Connection
                }
                //Footer
                mailbody.Append("</table>");

                string ddd = DateTime.Today.ToString("ddd");

                if(ddd == "Fri") mailbody.Append("<p style='font-size: 0.7em'>*Recall whether the TIPVest Pairs Account was used this week.</p>");

                mailbody.Append("</div></body>"); //Ending of HTML - Footer


                n.CreateEmailItem(Subject, Traders, mailbody);

                //Testing Variables
                //Console.Write(mailbody);

                Console.WriteLine("\nProgram Ended"); //Announces ending of program
                Console.WriteLine("-------------------------------------------------\n");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                //Console.WriteLine("Error: " + e.StackTrace);
            }
        }

        private void CreateEmailItem(string subjectEmail, string[] toEmail, StringBuilder bodyEmail) //Creates the email object
        {
            try
            {
                //Building of the email body and structure
                Outlook.Application oApp = new Outlook.Application();
                Outlook.MailItem eMail = (Outlook.MailItem) oApp.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Recipients Recips = (Outlook.Recipients)eMail.Recipients;
                eMail.Subject = subjectEmail;
                eMail.To = toEmail[0];
                for (int i = 1; i < toEmail.Length; i++) { if (toEmail[i] == null) break; eMail.Recipients.Add(toEmail[i]); }
                Recips = null;
                eMail.HTMLBody = bodyEmail.ToString();
                eMail.Send();

                Console.WriteLine("\nMail Sent!");
            }
            catch(Exception e)
            {
                Console.WriteLine("\n" + e.Message);
            }
        }

    }
}
