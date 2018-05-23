using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Xml;

namespace CiscoDataLoad
{
    class CiscoDataLoad
    {        
        public static void InsertIntoTable1(string reporturl, string connectionString)
        {
           
            //Source: Report1
            //Destination: Table1
            string insertQuery = @"INSERT INTO dbo.Table1(col1, col2, col3, col4, col5, col6) 
                                    VALUES(@col1, @col2, @col3, @col4, @col5, @col5, @col6);";

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
                    insertCmd.Parameters.Add("@col1", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@col2", SqlDbType.Date);
                    insertCmd.Parameters.Add("@col3", SqlDbType.Float);
                    insertCmd.Parameters.Add("@col4", SqlDbType.Int);
                    insertCmd.Parameters.Add("@col5", SqlDbType.SmallInt);
                    insertCmd.Parameters.Add("@col6", SqlDbType.NVarChar, 16);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string col1 = "";
                        DateTime? col2 = null;
                        double col3 = 0;
                        int col4 = 0;
                        int col5 = 0;
                        string col6 = "";

                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "col1":
                                    col1 = col.InnerText;
                                    break;
                                case "col2":
                                    col2 = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "col3":
                                    col3 = Convert.ToDouble(col.InnerText);
                                    break;
                                case "col4":
                                    col4 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "col5":
                                    col5 = Convert.ToInt16(col.InnerText);
                                    break;
                                case "col6":
                                    col6 = col.InnerText;
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@col1"].Value = col1;
                        insertCmd.Parameters["@col2"].Value = col2;
                        insertCmd.Parameters["@col3"].Value = col3;
                        insertCmd.Parameters["@col4"].Value = col4;
                        insertCmd.Parameters["@col5"].Value = col5;
                        insertCmd.Parameters["@col6"].Value = col6;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write(" Report Inserted Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table1\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoTable2(string reporturl, string connectionString)
        {
            //Source: Report2
            //Destination: Table2
            string insertQuery = @"INSERT INTO dbo.table2(col1, col2, col3, col4) 
                                    VALUES(@col1, @col2, @col3, @col4);";

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
                    insertCmd.Parameters.Add("@col1", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@col2", SqlDbType.Date);
                    insertCmd.Parameters.Add("@col3", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@col4", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string col1 = "";
                        DateTime? col2 = null;
                        string col3 = "";
                        int col4 = 0;
                        

                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "col1":
                                    col1 = col.InnerText;
                                    break;
                                case "col2":
                                    col2 = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "col3":
                                    col3 = col.InnerText;
                                    break;
                                case "col4":
                                    col4 = Convert.ToInt32(col.InnerText);
                                    break;                                
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@col1"].Value = col1;
                        insertCmd.Parameters["@col2"].Value = col2;
                        insertCmd.Parameters["@col3"].Value = col3;
                        insertCmd.Parameters["@col4"].Value = col4;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("nserted into Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table\r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoTable3(string reporturl, string connectionString)
        {
            //Source: Report 3
            //Destination:  Table 3
            string insertQuery = @"INSERT INTO dbo.table3(c1, c2, c3, c4) 
                                    VALUES(@c1, @c2, @c3, @c4);";

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
                    insertCmd.Parameters.Add("@c1", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@c2", SqlDbType.Date);
                    insertCmd.Parameters.Add("@c3", SqlDbType.Time, 7);
                    insertCmd.Parameters.Add("@c4", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string c1 = "";
                        DateTime? c2 = null;
                        TimeSpan? c3 = null;
                        int c4 = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "c1":
                                    c1 = col.InnerText;
                                    break;
                                case "c2":
                                    c2 = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "c3":
                                    //handle time format from with H:MM
                                    var timetmp = Convert.ToDateTime(col.InnerText);
                                    c3 = TimeSpan.Parse(timetmp.ToString("hh:mm"));
                                    break;
                                case "c4":
                                    c4 = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@c1"].Value = c1;
                        insertCmd.Parameters["@c2"].Value = c2;
                        insertCmd.Parameters["@c3"].Value = c3;
                        insertCmd.Parameters["@c4"].Value = c4;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Data Inserted into Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table \r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoTable4(string reporturl, string connectionString)
        {
            //Source: Report4
            //Destination: Table 4
            string insertQuery = @"INSERT INTO dbo.table4(c1, c2, c3, c4, c5, c6, c7,
                                   c8, c9, c10) 
                                    VALUES(@c1, @c2, @c3, @c4, @c5, @c6, @c7, 
                                    @c8, @c9, @c10);";

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
                    insertCmd.Parameters.Add("@c1", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@c2", SqlDbType.Date);
                    insertCmd.Parameters.Add("@c3", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c4", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c5", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c6", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c7", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c8", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c9", SqlDbType.Time);
                    insertCmd.Parameters.Add("@c10", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {
                        string c1 = "";
                        DateTime? c2 = null;
                        int c3 = 0;
                        int c4 = 0;
                        int c5 = 0;
                        int c6 = 0;
                        int c7 = 0;
                        int c8 = 0;
                        TimeSpan? c9 = null;
                        int c10 = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "c1":
                                    c1 = col.InnerText;
                                    break;
                                case "c2":
                                    c2 = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "c3":
                                    c3 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c4":
                                    c4 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c5":
                                    c5 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c6":
                                c6 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c7":
                                    c7 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c8":
                                    c8 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c9":
                                    //handle time format  with H:MM
                                    var timetmp = Convert.ToDateTime(col.InnerText);
                                    c9 = TimeSpan.Parse(timetmp.ToString("hh:mm"));                                  
                                    break;
                                case "c10":
                                    c10 = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@c1"].Value = c1;
                        insertCmd.Parameters["@c2"].Value = c2;
                        insertCmd.Parameters["@c3"].Value = c3;
                        insertCmd.Parameters["@c4"].Value = c4;
                        insertCmd.Parameters["@c5"].Value = c5;
                        insertCmd.Parameters["@c6"].Value = c6;
                        insertCmd.Parameters["@c7"].Value = c7;
                        insertCmd.Parameters["@c8"].Value = c8;
                        insertCmd.Parameters["@c9"].Value = c9;
                        insertCmd.Parameters["@c10"].Value = c10;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Data Inserted into Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table \r\n" + e.Message);
                }
            }
        }

        public static void InsertIntoTable5(string reporturl, string connectionString)
        {
            //Source: Report 5
            //Destination:  Table 5
            string insertQuery = @"INSERT INTO dbo.table5(c1, c2, c3, c4, c5, c6) 
                                    VALUES(@c1, @c2, @c3, @c4, @c5, @c6);";

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
                    insertCmd.Parameters.Add("@c1", SqlDbType.DateTime);
                    insertCmd.Parameters.Add("@c2", SqlDbType.VarChar, 64);
                    insertCmd.Parameters.Add("@c3", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c4", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c5", SqlDbType.Int);
                    insertCmd.Parameters.Add("@c6", SqlDbType.Int);

                    // open connection, loop over file 'rows' and 'column' tags, execute query to insert data one row at a time
                    conn.Open();

                    //parse XML file to extract row/col values
                    foreach (XmlNode row in report.ChildNodes)
                    {                        
                        DateTime? c1 = null;
                        string c2 = "";
                        int c3 = 0;
                        int c4 = 0;
                        int c5 = 0;
                        int c6 = 0;


                        foreach (XmlElement col in row)
                        {
                            switch (col.Attributes["name"].Value)
                            {
                                case "c1":
                                    c1 = Convert.ToDateTime(col.InnerText);
                                    break;
                                case "c2":
                                    c2 = col.InnerText;
                                    break;
                                case "c3":
                                    c3 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c4":
                                    c4 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c5":
                                    c5 = Convert.ToInt32(col.InnerText);
                                    break;
                                case "c6":
                                    c6 = Convert.ToInt32(col.InnerText);
                                    break;
                                default:
                                    break;
                            }
                        }

                        insertCmd.Parameters["@c1"].Value = c1;
                        insertCmd.Parameters["@c2"].Value = c2;
                        insertCmd.Parameters["@c3"].Value = c3;
                        insertCmd.Parameters["@c4"].Value = c4;
                        insertCmd.Parameters["@c5"].Value = c5;
                        insertCmd.Parameters["@c6"].Value = c6;
                        insertCmd.ExecuteNonQuery();

                        LogControl.Write("Data Inserted into  Table Successfully");

                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    LogControl.Write("Error at Insert into Table \r\n" + e.Message);
                }
            }
        }

    }
}
