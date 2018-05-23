using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Xml;

namespace CiscoDataLoad
{
    class CiscoDataLoad
    {        
        public static void InsertIntoCiscoAgentsDaily(string reporturl, string connectionString)
        {
           
            //Source: IT Agent Team Agent Daily Report from CUIC            
            //Destination: CiscoAgentsDaily Table
            string insertQuery = @"INSERT INTO dbo.ciscoagentsdaily(AgentName, Date, LoggedOnTime, TalkTime, CallsAnswered, AgentUserID) 
                                    VALUES(@AgentName, @Date, @Loggedontime, @Talktime, @Callsanswered, @AgentuserID);";

            //Load XML file into report variable
            XmlDocument document = new XmlDocument();
            document.Load(reporturl);
            var report = document.LastChild;

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
            {
                try
                {
                    // define parameters
                    insertCmd.Parameters.Add("@AgentName", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@Date", SqlDbType.Date);
                    insertCmd.Parameters.Add("@Loggedontime", SqlDbType.Float);
                    insertCmd.Parameters.Add("@TalkTime", SqlDbType.Int);
                    insertCmd.Parameters.Add("@Callsanswered", SqlDbType.SmallInt);
                    insertCmd.Parameters.Add("@AgentuserID", SqlDbType.NVarChar, 16);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string agentName = "";
                        DateTime? date = null;
                        double loggedOnTime = 0;
                        int talkTime = 0;
                        int callsAnswered = 0;
                        string agentUserID = "";

                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "AgentName":
                                    agentName = col.InnerText;
                                    break;
                                case "Date":
                                    date = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "LoggedOnTime":
                                    loggedOnTime = Convert.ToDouble(col.InnerText);
                                    break;
                                case "TalkTime":
                                    talkTime = Convert.ToInt32(col.InnerText);
                                    break;
                                case "CallsAnswered":
                                    callsAnswered = Convert.ToInt16(col.InnerText);
                                    break;
                                case "AgentUserID":
                                    agentUserID = col.InnerText;
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@AgentName"].Value = agentName;
                        insertCmd.Parameters["@Date"].Value = date;
                        insertCmd.Parameters["@Loggedontime"].Value = loggedOnTime;
                        insertCmd.Parameters["@TalkTime"].Value = talkTime;
                        insertCmd.Parameters["@CallsAnswered"].Value = callsAnswered;
                        insertCmd.Parameters["@AgentUserID"].Value = agentUserID;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Agent Daily Report Inserted Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table CiscoAgentsDaily\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoCiscoAgentsNotReady(string reporturl, string connectionString)
        {
            //Source: IT Agent Team Agent Daily Not Ready Detail Report from CUIC
            //Destination: CiscoAgentsNotReady Table
            string insertQuery = @"INSERT INTO dbo.ciscoagentsnotready(AgentName, Date, Reason, Duration) 
                                    VALUES(@AgentName, @Date, @Reason, @Duration);";

            //Load XML file into report variable
            XmlDocument document = new XmlDocument();
            document.Load(reporturl);
            var report = document.LastChild;

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
            {
                try
                {
                    // define parameters
                    insertCmd.Parameters.Add("@AgentName", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@Date", SqlDbType.Date);
                    insertCmd.Parameters.Add("@Reason", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@Duration", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string agentName = "";
                        DateTime? date = null;
                        string reason = "";
                        int duration = 0;
                        

                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "AgentName":
                                    agentName = col.InnerText;
                                    break;
                                case "Date":
                                    date = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "Reason":
                                    reason = col.InnerText;
                                    break;
                                case "Duration":
                                    duration = Convert.ToInt32(col.InnerText);
                                    break;                                
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@AgentName"].Value = agentName;
                        insertCmd.Parameters["@Date"].Value = date;
                        insertCmd.Parameters["@Reason"].Value = reason;
                        insertCmd.Parameters["@Duration"].Value = duration;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Daily Agent Not Ready Data Inserted into CiscoAgentsNotReady Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table CiscoAgentsNotReady\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoCiscoAgentsLogIn(string reporturl, string connectionString)
        {
            //Source: IT Agent Login and Logout Report from CUIC
            //Destination: CiscoAgentsLogIn Table
            string insertQuery = @"INSERT INTO dbo.ciscoagentslogin(AgentName, Date, Time, LoggedOnTime) 
                                    VALUES(@AgentName, @Date, @Time, @LoggedOnTime);";

            //Load XML file into report variable
            XmlDocument document = new XmlDocument();
            document.Load(reporturl);
            var report = document.LastChild;

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
            {
                try
                {
                    // define parameters
                    insertCmd.Parameters.Add("@AgentName", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@Date", SqlDbType.Date);
                    insertCmd.Parameters.Add("@Time", SqlDbType.Time, 7);
                    insertCmd.Parameters.Add("@LoggedOnTime", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string agentName = "";
                        DateTime? date = null;
                        TimeSpan? time = null;
                        int loggedonTime = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "AgentName":
                                    agentName = col.InnerText;
                                    break;
                                case "Date":
                                    date = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "Time":
                                    //handle time format from CUIC with H:MM
                                    var timetmp = Convert.ToDateTime(col.InnerText);
                                    time = TimeSpan.Parse(timetmp.ToString("hh:mm"));
                                    break;
                                case "LoginDuration":
                                    loggedonTime = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@AgentName"].Value = agentName;
                        insertCmd.Parameters["@Date"].Value = date;
                        insertCmd.Parameters["@Time"].Value = time;
                        insertCmd.Parameters["@LoggedOnTime"].Value = loggedonTime;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Agents Logged In Logged Out Data Inserted into CiscoAgentsLogIn Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table CiscoAgentsLogIn\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntociscoQueues(string reporturl, string connectionString)
        {
            //Source: IT Call Types Report from CUIC            
            //Destination: CiscoQueues Table
            string insertQuery = @"INSERT INTO dbo.ciscoqueues(EnterpriseName, Date, CallsAnswered, AnswerWaitTime, CallsOffered, HandleTime, TalkTime,
                                   TotalCallsAband, Time, DOW) 
                                    VALUES(@EnterpriseName, @Date, @CallsAnswered, @AnswerWaitTime, @CallsOffered, @HandleTime, @TalkTime, 
                                    @TotalCallsAband, @Time, @DOW);";

            //Load XML file into report variable
            XmlDocument document = new XmlDocument();
            document.Load(reporturl);
            var report = document.LastChild;

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
            {
                try
                {
                    // define parameters
                    insertCmd.Parameters.Add("@EnterpriseName", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@Date", SqlDbType.Date);
                    insertCmd.Parameters.Add("@CallsAnswered", SqlDbType.Int);
                    insertCmd.Parameters.Add("@AnswerWaitTime", SqlDbType.Int);
                    insertCmd.Parameters.Add("@CallsOffered", SqlDbType.Int);
                    insertCmd.Parameters.Add("@HandleTime", SqlDbType.Int);
                    insertCmd.Parameters.Add("@TalkTime", SqlDbType.Int);
                    insertCmd.Parameters.Add("@TotalCallsAband", SqlDbType.Int);
                    insertCmd.Parameters.Add("@Time", SqlDbType.Time);
                    insertCmd.Parameters.Add("@DOW", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string enterpriseName = "";
                        DateTime? date = null;
                        int callsAnswered = 0;
                        int answerWaitTime = 0;
                        int callsOffered = 0;
                        int handleTime = 0;
                        int talkTime = 0;
                        int totalCallsAband = 0;
                        TimeSpan? time = null;
                        int dow = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "EnterpriseName":
                                    enterpriseName = col.InnerText;
                                    break;
                                case "date":
                                    date = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "CallsAnswered":
                                    callsAnswered = Convert.ToInt32(col.InnerText);
                                    break;
                                case "AnswerWaitTime":
                                    answerWaitTime = Convert.ToInt32(col.InnerText);
                                    break;
                                case "CallsOffered":
                                    callsOffered = Convert.ToInt32(col.InnerText);
                                    break;
                                case "HandleTime":
                                handleTime = Convert.ToInt32(col.InnerText);
                                    break;
                                case "TalkTime":
                                    talkTime = Convert.ToInt32(col.InnerText);
                                    break;
                                case "TotalCallsAband":
                                    totalCallsAband = Convert.ToInt32(col.InnerText);
                                    break;
                                case "Time":
                                    //handle time format from CUIC with H:MM
                                    var timetmp = Convert.ToDateTime(col.InnerText);
                                    time = TimeSpan.Parse(timetmp.ToString("hh:mm"));                                  
                                    break;
                                case "DOW":
                                    dow = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@EnterpriseName"].Value = enterpriseName;
                        insertCmd.Parameters["@Date"].Value = date;
                        insertCmd.Parameters["@CallsAnswered"].Value = callsAnswered;
                        insertCmd.Parameters["@AnswerWaitTime"].Value = answerWaitTime;
                        insertCmd.Parameters["@CallsOffered"].Value = callsOffered;
                        insertCmd.Parameters["@HandleTime"].Value = handleTime;
                        insertCmd.Parameters["@TalkTime"].Value = talkTime;
                        insertCmd.Parameters["@TotalCallsAband"].Value = totalCallsAband;
                        insertCmd.Parameters["@Time"].Value = time;
                        insertCmd.Parameters["@DOW"].Value = dow;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("IT Call Types Data Inserted into CiscoQueues Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table CiscoQueues\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoCiscoInboundOutbound(string reporturl, string connectionString)
        {
            //Source: IT Agent Team Agent Skill Group Interval Report from CUIC
            //Destination: CiscoInboundOutbound Table
            string insertQuery = @"INSERT INTO dbo.ciscoinboundoutbound(DateTime, EnterpriseName, CallsAnswered, InboundHandleTime, OutboundCalls, OutboundHandleTime) 
                                    VALUES(@DateTime, @EnterpriseName, @CallsAnswered, @InboundHandleTime, @OutboundCalls, @OutboundHandleTime);";

            //Load XML file into report variable
            XmlDocument document = new XmlDocument();
            document.Load(reporturl);
            var report = document.LastChild;

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
            {
                try
                {
                    // define parameters
                    insertCmd.Parameters.Add("@DateTime", SqlDbType.DateTime);
                    insertCmd.Parameters.Add("@EnterpriseName", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@CallsAnswered", SqlDbType.Int);
                    insertCmd.Parameters.Add("@InboundHandleTime", SqlDbType.Int);
                    insertCmd.Parameters.Add("@OutboundCalls", SqlDbType.Int);
                    insertCmd.Parameters.Add("@OutboundHandleTime", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {                        
                        DateTime? dateTime = null;
                        string enterpriseName = "";
                        int callsAnswered = 0;
                        int inboundHandleTime = 0;
                        int outboundCalls = 0;
                        int outboundHandleTime = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "DateTime":
                                    dateTime = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "EnterpriseName":
                                    enterpriseName = col.InnerText;
                                    break;
                                case "CallsAnswered":
                                    callsAnswered = Convert.ToInt32(col.InnerText);
                                    break;
                                case "InboundHandleTime":
                                    inboundHandleTime = Convert.ToInt32(col.InnerText);
                                    break;
                                case "OutboundCalls":
                                    outboundCalls = Convert.ToInt32(col.InnerText);
                                    break;
                                case "OutboundHandleTime":
                                    outboundHandleTime = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@DateTime"].Value = dateTime;
                        insertCmd.Parameters["@EnterpriseName"].Value = enterpriseName;
                        insertCmd.Parameters["@CallsAnswered"].Value = callsAnswered;
                        insertCmd.Parameters["@InboundHandleTime"].Value = inboundHandleTime;
                        insertCmd.Parameters["@OutboundCalls"].Value = outboundCalls;
                        insertCmd.Parameters["@OutboundHandleTime"].Value = outboundHandleTime;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Team Skill Group Interval Data Inserted into CiscoInboundOutbound Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table CiscoInboundOutbound\r\n" + e.Message);
                }
            }
        }

    }
}
