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
            string connectionString = @"server=mspm1bcl23b1\b1;database=ITReportingD;integrated security=SSPI;";
            #region US Reports Load

            //US Reports
            string USAgentDailyReportUrl = "https://msplucc1c1a.ent.core.medtronic.com:8444/cuicui/permalink/?linkType=xmlType&viewType=Grid&viewId=6EE57DF6100001550003F9B80A308FF1";
            string USAgentDailyNotReadyReportUrl = "https://msplucc1c1a.ent.core.medtronic.com:8444/cuicui/permalink/?linkType=xmlType&viewType=Grid&viewId=6EFCC6FF10000155000402370A308FF1";
            string USAgentLoginLogOutReportUrl = "https://msplucc1c1a.ent.core.medtronic.com:8444/cuicui/permalink/?linkType=xmlType&viewType=Grid&viewId=6F7AC6DA1000015500042FF00A308FF1";
            string USCallTypesReportUrl = "https://msplucc1c1a.ent.core.medtronic.com:8444/cuicui/permalink/?linkType=xmlType&viewType=Grid&viewId=0956F6861000014B5D2C48F40A308FB5";
            string USAgentSkillGroupIntervalUrl = "https://msplucc1c1a.ent.core.medtronic.com:8444/cuicui/permalink/?linkType=xmlType&viewType=Grid&viewId=E17265CE100001600170877B0A308FF1";

            //Load US Reports 
            try
            {
                //CUIC server uses TLS1.2 secuirty protocal
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                CiscoDataLoad.InsertIntoCiscoAgentsDaily(USAgentDailyReportUrl, connectionString);
                CiscoDataLoad.InsertIntoCiscoAgentsNotReady(USAgentDailyNotReadyReportUrl, connectionString);
                CiscoDataLoad.InsertIntoCiscoAgentsLogIn(USAgentLoginLogOutReportUrl, connectionString);
                CiscoDataLoad.InsertIntociscoQueues(USCallTypesReportUrl, connectionString);
                CiscoDataLoad.InsertIntoCiscoInboundOutbound(USAgentSkillGroupIntervalUrl, connectionString);
            }
            catch (Exception ex)
            {
                LogControl.Write("US Reports Inserted failed with error\r\n" + ex.Message);

            }
            
            #endregion


            #region EMEA Reports

            //EMEA Reports
            string EMEAAgentDailyReportUrl = "https://mstlappuic01.mst.medtronic.com:8444/cuic/permalink/PermalinkViewer.htmx?linkType=xmlType&viewType=Grid&viewId=D835ED7110000157004B49000AAF12B5";
            string EMEAAgentDailyNotReadyReportUrl = "https://mstlappuic01.mst.medtronic.com:8444/cuic/permalink/PermalinkViewer.htmx?linkType=xmlType&viewType=Grid&viewId=D8D296B010000157004B66900AAF12B5";
            string EMEAAgentLoginLogOutReportUrl = "https://mstlappuic01.mst.medtronic.com:8444/cuic/permalink/PermalinkViewer.htmx?linkType=xmlType&viewType=Grid&viewId=D8638CB310000157004B52970AAF12B5";
            string EMEACallTypesReportUrl = "https://mstlappuic01.mst.medtronic.com:8444/cuic/permalink/PermalinkViewer.htmx?linkType=xmlType&viewType=Grid&viewId=D8E0683110000157004B68EF0AAF12B5";

            try
            {
                //ignore unverfied certificate warning from server for CUIC EMEA reports
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

                CiscoDataLoad.InsertIntoCiscoAgentsDaily(EMEAAgentDailyReportUrl, connectionString);
                CiscoDataLoad.InsertIntoCiscoAgentsNotReady(EMEAAgentDailyNotReadyReportUrl, connectionString);
                CiscoDataLoad.InsertIntoCiscoAgentsLogIn(EMEAAgentLoginLogOutReportUrl, connectionString);
                CiscoDataLoad.InsertIntociscoQueues(EMEACallTypesReportUrl, connectionString);

            }
            catch (Exception ex)
            {
                LogControl.Write("EMEA Reports Inserted failed with error\r\n" + ex.Message);                
            }

            #endregion

        }
    }
}
