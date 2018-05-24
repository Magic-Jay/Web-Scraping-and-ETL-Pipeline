using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Xml;

namespace CiscoDataLoad
{
    class Program
    { 
        static void Main(string[] args)
        {
            //sql connection string
            string connectionString = @"server=;database=;integrated security=SSPI;";
            #region US Reports Load

            //Reports
            string ReportUrl1 = "https:";
            string ReportUrl2 = "https:";
            string ReportUrl3 = "https:";
            string ReportUrl4 = "https:";
            string Url5 = "https:";

            //Load US Reports 
            try
            {
                //server uses TLS1.2 secuirty protocal
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                CiscoDataLoad.InsertIntoTable1(ReportUrl1, connectionString);
                CiscoDataLoad.InsertIntoTable2(ReportUrl2, connectionString);
                CiscoDataLoad.InsertIntoTable3(ReportUrl3, connectionString);
                CiscoDataLoad.InsertIntoTable4(ReportUrl4, connectionString);
                CiscoDataLoad.InsertIntoTable5(Url5, connectionString);
            }
            catch (Exception ex)
            {
                LogControl.Write("Reports Inserted failed with error\r\n" + ex.Message);

            }
            
            #endregion


            #region EMEA Reports

            //Reports
            string DailyReportUrl21 = "";
            string ReportUrl22 = "";
            string ReportUrl23 = "";
            string ReportUrl24 = "";

            try
            {
                //ignore unverfied certificate warning from server for reports
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                CiscoDataLoad.InsertIntoTable1(DailyReportUrl21, connectionString);
                CiscoDataLoad.InsertIntoTable2(ReportUrl22, connectionString);
                CiscoDataLoad.InsertIntoTable3(ReportUrl23, connectionString);
                CiscoDataLoad.InsertIntoTable4(ReportUrl24, connectionString);

            }
            catch (Exception ex)
            {
                LogControl.Write("Reports Inserted failed with error\r\n" + ex.Message);                
            }

            #endregion

        }
    }
}
