using System;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using EASendMail; //add EASendMail namespace (This needs the license code)



//=====================[Read This]===========================
//    if you are reading this, it means that you have got into my source code. Not hard to do.
//
//    But this is my code, and mine alone.
//    It has been developed by me, is my intellectual property and is what I use to feed my family.
//    So what you are doing is stealing from me, and worse, you are stealing from my family.

//    So here's the thing.
//
//    If I catch you, and believe me this wont be hard, I will fucking crucify you, the company you work for, and if you are a contractor,
//    I will take everything I can from you.
//
//    Take this seriously.
//
//    Trimble is a $2bn US company with 6000 employees. They tried to screw me.
//    In 2020 I took them to court.
//    Me, on my own.
//    I won.
//    6 Figure $ settlement thank you.
//    You willing to take that chance? You got that sort of money to defend your stealing code?
//
//    People think I am this pleasant smiling guy from South Africa, always willing to help.
//    Yeah yeah whatever, until someone crosses me.
//    Then my retribution knows no bounds, and I will aggressively and actively persue and destroy anyone that touches me or my family.

//    Dont steal from me. Dont even contemplate such a thing.
//    Close this code and delete it from your computer, your memory stick, your mind, your scraps of notes, wherever you have uploaded it,
//    and go have a coffee.
//    Forget you even saw this.

//    The alternative is not worth it.
//
//     Trust me on that one....... as I said....I will fucking crucify you.
//     Stone fucking dead.
//
//=======================================================================================================================================
//
// To add this to a project..
// Right Click References
//  add the reference
//  add Reference.. browse to GNALibrary from the bin/Release location or use the Recent option
// include after static void Main(): 
//      gnaToolbox gna = new gnaToolbox();
//
// The versions of the packages must match or you get an error message that the version could not be found.
// Update so that the packages all match
//
// Remember to recompile gnaToolbox using Release after any change
//


namespace GNAlibrary
{
    public class gnaToolbox
    {

#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416


        public double add(double num_one, double num_two)
        {
            return num_one + num_two;
        }


        //===================================================================

        public SqlCommand SqlCommand { get; private set; } = null!;


        //============= General Messages ========

        public void WelcomeMessage(string strMessage)
        {
            Console.Title = strMessage;

            Console.WriteLine(" ");
            Console.WriteLine("GNA Geomatics software");
            Console.WriteLine("Julian Gray");
            Console.WriteLine("+49 176 7299 7904");
            Console.WriteLine("gna.geomatics@gmail.com");
            Console.WriteLine(" ");
            Console.WriteLine("Do not steal my software..");
            Console.WriteLine("--------------------------");
            Console.WriteLine(" ");
            Console.WriteLine("");
        }



        public void HelloWorld()
        {
            Console.WriteLine("GNAlibrary.gnaToolbox : Hello world");
            Console.ReadKey();
        }

        public string Test(string strText)
        {
            strText += " : data received, processed and returned from GNA module";
            return strText;
        }

        public void GNAlibraryVersion()
        {
            Console.WriteLine("version 2019.04.19.1");
            Console.ReadKey();
        }

        public void testDBconnection(string strDBconnection)
        {
            // 
            // Purpose
            //  To test a DB Connection
            // Useage
            //  gna.testDBconnection(strDBconnection);
            // Output
            //  Successful or Failed with error message
            //
            // If you suddenly start having failed connections, check that the versions of the SQL packages match between the GNAlibrary and the calling software
            // If they dont, then update so that they both match
            //

            //instantiate and open connection

            Console.WriteLine("Module: gna.testDBconnection");

            SqlConnection conn = new SqlConnection(strDBconnection);

            try
            {
                conn.Open();
                Console.WriteLine("DB Connection Successful!");
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                Console.WriteLine("DB Connection Failed: ");
                Console.WriteLine("    Check the connection string");
                Console.WriteLine("    Check the name of the DB in the connection string, sometimes T4D other times TPP..");
                Console.WriteLine("    Check the password, TPP: Tr1mbl3 and T4D: Tr1mbl3!..");
                Console.WriteLine("");
                Console.WriteLine("Error message:");
                Console.WriteLine(ex);
                Console.ReadLine();
            }

            finally
            {
                conn.Dispose();
                conn.Close();
                Console.WriteLine("DB Connection tested, and successful");
                Console.WriteLine("");
            }

        }



        //=================[ Registry Methods ]=======================================================================================================





        public string checkSoftwareReferenceDate(string strProject, string strEmailLogin, string strEmailPassword, string strSendEmail)
        {


            string strStatus = "empty";

            //
            // Check the number of days remaining on the license
            // The start date and validity period must be set using the software "SetSoftwareKey"
            //

            // This module must be at the start of all software modules
            // use:
            //
            // string strSendEmail = "No";
            // string strSoftwareKey = gna.checkSoftwareReferenceDate(strProject, strEmailLogin, strEmailPassword, strSendEmail);
            // if (strSoftwareKey == "expired") goto TheEnd;
            //
            // TheEnd:
            //  Console.WriteLine("Done");
            //
            // the answer is either "valid" or "expired"
            //  strSendEmail = "Yes" when CheckSoftwareKey is set as a scheduled task 
            //  strSendEmail is set to "No" when used inside the software
            //

            //Console.WriteLine("checkSoftwareReferenceDate");
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Diebus");
            string strValidityPeriod = key.GetValue("TempusValide", "No Value").ToString();
            key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Portunus");
            string strReferenceDate = key.GetValue("Clavis", "No Value").ToString();

            Console.WriteLine("Licence period: " + strValidityPeriod + " days");
            Console.WriteLine("Reference date: " + strReferenceDate);

            DateTime InstallDate = DateTime.ParseExact(strReferenceDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.Today;

            TimeSpan interval = (TodayDate - InstallDate);
            int iRemainingDays = Convert.ToInt16(strValidityPeriod) - interval.Days;

            Console.WriteLine("Remaining days: " + iRemainingDays.ToString() + " days");

            if (iRemainingDays < 0)
            {
                strStatus = "expired";

                Console.WriteLine(" ");
                Console.WriteLine("The software license has expired.");
                Console.WriteLine("Please contact Julian Gray on +49 176 7299 7904 to reactivate");
                Console.WriteLine(" ");
                Console.ReadLine();
            }
            else
            {
                strStatus = "valid";
            }

            if ((iRemainingDays < 4) & (strSendEmail == "Yes"))
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");
                    {
                        oMail.From = "gna.geomatics@gmail.com";
                        oMail.To = new AddressCollection("gna.geomatics@gmail.com");
                    };

                    // Set email subject
                    oMail.Subject = "Software license about to expire (" + strProject + ")";
                    // Set email body
                    oMail.TextBody = "No of days remaining: " + iRemainingDays.ToString();

                    // SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                    Console.WriteLine("Advisory email issued");

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Email transmission failed..");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine("");
                    Console.WriteLine(ep.Message);
                    Console.ReadKey();
                }
            }

            return strStatus;
        }

        //=================[ email methods ]=============================================================================

        public void emailExcelWorkbook(string strEmailFrom, string strEmailRecipients, string strEmailLogin, string strEmailPassword, string strReportHeader, string strExcelWorkbook)
        {

            try
            {
                SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3")
                {

                    // Set sender email address, please change it to yours
                    //oMail.From = "julian.gray@korecgroup.com";

                    From = strEmailFrom,

                    //oMail.To = new AddressCollection("Test1<test@adminsystem.com>, Test2<test2@adminsystem.com>");
                    To = new AddressCollection(strEmailRecipients)
                };

                // Generate the time stamp
                DateTime now = DateTime.Now;
                string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");

                // Set email subject
                oMail.Subject = strReportHeader + " (" + strDateTime + ")";
                // Set email body
                oMail.TextBody = "This is an automated monitoring report issued by the monitoring system. Please do not reply to this email.";
                // Add the attachment
                oMail.AddAttachment(strExcelWorkbook);

                // SMTP server address
                SmtpServer oServer = new SmtpServer("smtp.gmail.com")
                {
                    // User and password for ESMTP authentication
                    // oServer.User = "gna.monitoringsoftware@gmail.com";
                    // oServer.Password = "Tr1mbl3!";

                    User = strEmailLogin,
                    Password = strEmailPassword,

                    // Most modern SMTP servers require SSL/TLS connection now.
                    // ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
                    ConnectType = SmtpConnectType.ConnectTryTLS,

                    // Set the SMTP server 587 port
                    Port = 587
                };

                SmtpClient oSmtp = new SmtpClient();
                oSmtp.SendMail(oServer, oMail);

            }
            catch (Exception ep)
            {
                Console.WriteLine("Email transmission failed..");
                Console.WriteLine(strEmailLogin);
                Console.WriteLine(strEmailPassword);
                Console.WriteLine(strEmailFrom);
                Console.WriteLine(strEmailRecipients);
                Console.WriteLine("");
                Console.WriteLine(ep.Message);
                Console.ReadKey();
            }

        }

        //=================[ Date Time methods ]======================================================================================================

        public Tuple<string, string> generateTimeBlockStartEnd(int iOffsetFromNow, int iTimeBlockSize, int iTimeBlockNumber)
        {
            // Purpose: 
            //      To compute the start and end times of the time block used to extract data from the DB
            // Input: 
            //      OffsetFromNow, TimeBlockSize, TimeBlockNumber (hours,integer,integer)
            // Output:
            //      BlockStartTime, BlockEndTime
            // Use:
            //      App.config gives the initial time offset, block size and number of blocks
            //      For time block number from 0 to number of blocks
            //          extract the start and end times
            //              var answer = gna.TimeBlockStartEnd(2, 4, 0);
            //              string strTimeBlockStartUTC = answer.Item1;
            //              string strTimeBlockEndUTC = answer.Item2;
            //          extract the data from the DB within those start and end times
            //          update the worksheet

            iOffsetFromNow = iOffsetFromNow + ((iTimeBlockNumber - 1) * iTimeBlockSize);
            double dblStartTimeOffset = -1.0 * (Convert.ToDouble(iOffsetFromNow));
            double dblEndTimeOffset = dblStartTimeOffset - (Convert.ToDouble(iTimeBlockSize));

            string strTimeBlockStartUTC = " '" + DateTime.UtcNow.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm") + "' ";
            string strTimeBlockEndUTC = " '" + DateTime.UtcNow.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm") + "' ";
            return new Tuple<string, string>(strTimeBlockStartUTC, strTimeBlockEndUTC);
        }

        public string[,] generateTimeBlockArray(int iTimeBlockSize, int iNoOfDays)
        {
            //
            // Purpose:
            //      To create an array of start/end time blocks between Now and back iNoOfDays
            // Input:
            //      No of days back from today (Rolling window)
            //      Time block size (hours)
            //      No of time blocks
            // Output:
            //      Returns strTimeBlockArray [counter, start time, end time] [iCounter,0-1]
            // Useage:
            //      int iBlockSize = Convert.ToInt32(strBlockSize);
            //      int iNoOfDays = Convert.ToInt16(strRollingTimeWindowDays); ;
            //      string[,] strTimeBlockArray = gna.generateTimeBlockArray(iBlockSize, iNoOfDays);
            //      string[iCounter,0] : strCounter
            //      string[iCounter,1] : strStartTime
            //      string[iCounter,2] : strEndTime
            // Comment:
            //      Last time = "NoMore"
            //

            //double dblTimeOffset = Convert.ToDouble(iOffsetFromNow * (-1.0));
            double dblTimeBlockSize = Convert.ToDouble(iTimeBlockSize);
            //DateTime dtUTCnow = DateTime.UtcNow;
            //DateTime dtEndTime = DateTime.UtcNow;
            //DateTime dtStartTime = DateTime.UtcNow;

            string[,] strTimeBlockArray = new string[2000, 3];

            // Compute the start date/time (local time)
            DateTime dtToday = DateTime.Today;
            double dblDaysOffset = Convert.ToDouble(iNoOfDays * (-1.0));
            double dblHoursOffset = 100;
            DateTime dtStartDate = dtToday.AddDays(dblDaysOffset);
            DateTime dtNow = DateTime.Now;
            DateTime dtStartTime = dtStartDate;
            DateTime dtEndTime = dtStartTime.AddHours(dblTimeBlockSize);

            int iCounter = 0;

            do
            {
                dblHoursOffset = (dtNow - dtEndTime).TotalHours;
                string strStartTime = "'" + dtStartTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                string strEndTime = "'" + dtEndTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";

                strTimeBlockArray[iCounter, 0] = Convert.ToString(iCounter);
                strTimeBlockArray[iCounter, 1] = strStartTime;
                strTimeBlockArray[iCounter, 2] = strEndTime;
                iCounter++;

                dtStartTime = dtEndTime;
                dtEndTime = dtStartTime.AddHours(dblTimeBlockSize);
                dblHoursOffset = (dtNow - dtEndTime).TotalHours;

            } while (dblHoursOffset >= 0);

            strTimeBlockArray[iCounter, 0] = "NoMore";
            strTimeBlockArray[iCounter, 1] = "2000-01-01 00:00:00";
            strTimeBlockArray[iCounter, 2] = "2000-01-01 01:00:00";
            return strTimeBlockArray;
        }


        public double UTCtimeOffset()
        {
            //
            // Purpose:
            //      To generate the UTC time offset for the server
            // Input:
            //      Nothing
            // Output:
            //      The difference between UTC and now in double hours
            // Useage:
            //      double dblUTCoffset = gna.UTCtimeOffset();
            //

            double dblUTCtimeOffset = 0;

            DateTime dtUTCnow = DateTime.UtcNow;
            DateTime dtNow = DateTime.Now;
            dblUTCtimeOffset = Convert.ToDouble((dtUTCnow - dtNow).Hours);

            return dblUTCtimeOffset;
        }

        //=================[ Database methods ]=============================================================================================



        public string getProjectID(string strDBconnection, string strProjectTitle)
        {
            // Purpose:
            //      To determine the project ID from TMCMonitoringProjects
            // Input:
            //      Receives Project Title
            // Output:
            //      Returns Project ID
            // Useage:
            //      string strProjectID = gna.getProjectID(strDBconnection, strProjectTitle);
            //

            string strProjectID = "";
            Int16 iCounter = 0;

            // Connection and Reader declared outside the try block
            using (SqlConnection conn = new SqlConnection(strDBconnection))
            {

                //instantiate and open connection
                //conn = new SqlConnection(strDBconnection);
                conn.Open();

                try
                {
                    // define the SQL query
                    string SQLaction = @"
                    SELECT ID, ProjectTitle  
                    FROM dbo.TMCMonitoringProjects 
                    WHERE TMCMonitoringProjects.ProjectTitle = @ProjectName
                    AND TMCMonitoringProjects.IsDeleted = 0
                    ";
                    SqlCommand cmd = new SqlCommand(SQLaction, conn);

                    // define the parameter used in the command object and add to the command
                    cmd.Parameters.Add(new SqlParameter("@ProjectName", strProjectTitle));

                    // Define the data reader
                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // get the values
                    while (dataReader.Read())
                    {
                        int iProjectID = (Int32)dataReader["ID"];
                        strProjectID = Convert.ToString(iProjectID);
                        iCounter++;
                    }


                    // Close the dataReader
                    if (dataReader != null)
                    {
                        dataReader.Close();
                    }

                }

                catch (System.Data.SqlClient.SqlException ex)
                {
                    Console.WriteLine("getPointCoordinates: DB Connection Failed when retrieving Project ID => Project name not correct : ");
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }

                finally
                {
                    conn.Dispose();
                    conn.Close();
                }
            }

            if (iCounter == 0) { strProjectID = "Missing"; }

            return strProjectID;

        }

        public string getLocationID(string strDBconnection, string strProjectID, string strPointName)
        {
            // Purpose:
            //      To determine the Location ID  
            // Assumption:
            //      The point name =  the location name
            // Output:
            //      Returns location ID
            // Useage:
            //      string strLocationID = gna.getLocationID(strDBconnection, strProjectID, strPointName);
            //

            string strLocationID = "Missing";

            // Connection and Reader declared outside the try block
            SqlConnection conn = new SqlConnection(strDBconnection);
            conn.Open();

            try
            {
                // define the SQL query
                string SQLaction = @"
                SELECT [ID]
                FROM [dbo].[TMCLocation]
                WHERE [ProjectID]= @ProjectID
                AND [Name]=@PointName
                ";

                SqlCommand cmd = new SqlCommand(SQLaction, conn);

                // define the parameter used in the command object and add to the command
                cmd.Parameters.Add(new SqlParameter("@ProjectID", strProjectID));
                cmd.Parameters.Add(new SqlParameter("@PointName", strPointName));

                // Define the data reader
                SqlDataReader dataReader = cmd.ExecuteReader();

                // get the values
                while (dataReader.Read())
                {
                    int iLocationID = (Int32)dataReader["ID"];
                    strLocationID = Convert.ToString(iLocationID);
                }


                // Close the dataReader
                if (dataReader != null)
                {
                    dataReader.Close();
                }

            }

            catch (System.Data.SqlClient.SqlException ex)
            {
                Console.WriteLine("getLocationID: ");
                Console.WriteLine("");
                Console.WriteLine("Suggection: Check the project name in the config file");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
            }

            finally
            {
                conn.Dispose();
                conn.Close();
            }

            return strLocationID;

        }







        public string[,] getPointDeltas(string strDBconnection, string strProjectTitle, string strTimeBlockStart, string strTimeBlockEnd, string[,] strNamesID)
        {
            //
            // Purpose:
            //      To extract the mean dN,dE,dH,dR,dT from dbo.TMTPosition_Terrestrial table for the time block strTimeBlockStart to strTimeBlockEnd
            // Input:
            //      Receives 
            //          array of point names & sensorID generated by getPointID(from DB) or readPointNamesSensorID(from Reference worksheet)
            //          the Project Title from the config file
            //          the start and end time blocks
            // Output:
            //      Returns array [PointName,dN,dE,dH, dR, dT, number of points used to compute mean]   [0,1,2,3,4,5,6]
            //      strPointDeltas[iCounter, 0] = strPointName;
            //      strPointDeltas[iCounter, 1] = MeandN
            //      strPointDeltas[iCounter, 2] = MeandE
            //      strPointDeltas[iCounter, 3] = MeandH
            //      strPointDeltas[iCounter, 4] = MeandR
            //      strPointDeltas[iCounter, 5] = MeandT
            //      strPointDeltas[iCounter, 6] = ObservationCounter = "-99" id there are no observations
            //
            //
            //
            // Useage:
            //      string[,] strPointDeltas = gna.getPointDeltas(strDBconnection, strProjectTitle, strTimeBlockStart, strTimeBlockEnd, strNamesID);
            // Comment:
            //      If missing then deltas are 0,0,0,-99
            //      last point in list = "NoMore"
            //

            string[,] strDeltas = new string[2000, 7];
            string strPointName = "";
            string strPointID = "";
            int iCounter = 0;
            double dbldN = 0.0;
            double dbldE = 0.0;
            double dbldH = 0.0;
            double dbldR = 0.0;
            double dbldT = 0.0;
            double dblMeandN = 0.0;
            double dblMeandE = 0.0;
            double dblMeandH = 0.0;
            double dblMeandR = 0.0;
            double dblMeandT = 0.0;
            int iObservationCounter = 0;

            // Select the block of observations for the point within the Time Block: between strRefBlockStart and strRefBlockEnd
            // generate the mean dN, dE, dH, dR, dT

            do
            {
                strPointName = strNamesID[iCounter, 0];
                strPointID = strNamesID[iCounter, 1];

                SqlConnection conn = new SqlConnection(strDBconnection);
                conn.Open();

                try
                {
                    if (strPointID == "Missing") goto ComputeMeans;

                    string SQLaction = @"SELECT * FROM dbo.TMTPosition_Terrestrial " +
                        "WHERE SensorID = @SensorID " +
                        "AND EndTimeUTC BETWEEN " + strTimeBlockStart +
                        "AND " + strTimeBlockEnd;

                    string strTemp = SQLaction;
                    SqlCommand cmd = new SqlCommand(SQLaction, conn);

                    // define the parameter used in the command object and add to the command
                    cmd.Parameters.Add(new SqlParameter("@SensorID", strPointID));

                    // Define the data reader
                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // Now read through the results and generate a mean value
                    dblMeandN = 0.0;
                    dblMeandE = 0.0;
                    dblMeandH = 0.0;
                    dblMeandR = 0.0;
                    dblMeandT = 0.0;

                    iObservationCounter = 0;

                    while (dataReader.Read())
                    {
                        dbldN = Math.Round(Convert.ToDouble(dataReader["dN"]), 4);
                        dbldE = Math.Round(Convert.ToDouble(dataReader["dE"]), 4);
                        dbldH = Math.Round(Convert.ToDouble(dataReader["dH"]), 4);
                        dbldR = Math.Round(Convert.ToDouble(dataReader["dR"]), 4);
                        dbldT = Math.Round(Convert.ToDouble(dataReader["dT"]), 4);

                        //Console.WriteLine(Convert.ToString(dbldN) + "  " + Convert.ToString(dbldE) + "  " + Convert.ToString(dbldH));

                        dblMeandN += dbldN;
                        dblMeandE += dbldE;
                        dblMeandH += dbldH;
                        dblMeandR += dbldR;
                        dblMeandT += dbldT;
                        iObservationCounter++;
                    }

                    // Close the dataReader
                    if (dataReader != null)
                    {
                        dataReader.Close();
                    }
                }

                catch (System.Data.SqlClient.SqlException ex)
                {
                    Console.WriteLine("getPointDeltas: DB Connection Failed: ");
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }

                finally
                {
                    conn.Dispose();
                    conn.Close();
                }

ComputeMeans:

                if ((strPointID != "Missing") && (iObservationCounter > 0))
                {
                    // Compute the mean dN, dE, dH
                    dblMeandN = Math.Round((dblMeandN / iObservationCounter), 4);
                    dblMeandE = Math.Round((dblMeandE / iObservationCounter), 4);
                    dblMeandH = Math.Round((dblMeandH / iObservationCounter), 4);
                    dblMeandR = Math.Round((dblMeandR / iObservationCounter), 4);
                    dblMeandT = Math.Round((dblMeandT / iObservationCounter), 4);
                }
                else
                {
                    // allocate false values
                    dblMeandN = 0.0;
                    dblMeandE = 0.0;
                    dblMeandH = 0.0;
                    dblMeandR = 0.0;
                    dblMeandT = 0.0;
                    iObservationCounter = -99;
                }

                //Console.WriteLine("Mean");
                //Console.WriteLine(Convert.ToString(dblMeandN) + "  " + Convert.ToString(dblMeandE) + "  " + Convert.ToString(dblMeandH));


                //Insert the data into the data arrays
                strDeltas[iCounter, 0] = strPointName;
                strDeltas[iCounter, 1] = Convert.ToString(dblMeandN);
                strDeltas[iCounter, 2] = Convert.ToString(dblMeandE);
                strDeltas[iCounter, 3] = Convert.ToString(dblMeandH);
                strDeltas[iCounter, 4] = Convert.ToString(dblMeandR);
                strDeltas[iCounter, 5] = Convert.ToString(dblMeandT);
                strDeltas[iCounter, 6] = Convert.ToString(iObservationCounter);
                iCounter++;
            } while (strPointName != "NoMore");

            strDeltas[iCounter, 0] = "NoMore";
            strDeltas[iCounter, 1] = "999";
            strDeltas[iCounter, 2] = "999";
            strDeltas[iCounter, 3] = "999";
            strDeltas[iCounter, 4] = "999";
            strDeltas[iCounter, 5] = "999";
            strDeltas[iCounter, 6] = "0";

            return strDeltas;
        }

        public string[,] getPointDh(string strDBconnection, string strProjectTitle, string strTimeBlockStart, string strTimeBlockEnd, string[,] strNamesID)
        {

            string[,] strDeltas = new string[2000, 5];
            string strPointName = "";
            string strPointID = "";
            int iCounter = 0;
            double dbldN = 0.0;
            double dbldE = 0.0;
            double dbldH = 0.0;
            double dblMeandN = 0.0;
            double dblMeandE = 0.0;
            double dblMeandH = 0.0;
            int iObservationCounter = 0;

            // Select the block of observations for the point within the Time Block: between strRefBlockStart and strRefBlockEnd
            // generate the mean dN, dE, dH

            do
            {
                strPointName = strNamesID[iCounter, 0];
                strPointID = strNamesID[iCounter, 1];

                SqlConnection conn = new SqlConnection(strDBconnection);
                conn.Open();

                try
                {
                    if (strPointID == "Missing") goto ComputeMeans;

                    string SQLaction = @"SELECT * FROM dbo.TMTPosition_Terrestrial " +
                        "WHERE SensorID = @SensorID " +
                        "AND EndTimeUTC BETWEEN " + strTimeBlockStart +
                        "AND " + strTimeBlockEnd;

                    //                string SQLaction = @"SELECT * FROM dbo.TMTPosition_Terrestrial " +
                    //"WHERE SensorID = @SensorID " +
                    //"AND EndTimeUTC BETWEEN " + strRefBlockStart +
                    //"AND " + strRefBlockEnd;


                    string strTemp = SQLaction;
                    SqlCommand cmd = new SqlCommand(SQLaction, conn);

                    // define the parameter used in the command object and add to the command
                    cmd.Parameters.Add(new SqlParameter("@SensorID", strPointID));

                    // Define the data reader
                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // Now read through the results and generate a mean value
                    dblMeandN = 0.0;
                    dblMeandE = 0.0;
                    dblMeandH = 0.0;
                    iObservationCounter = 0;

                    while (dataReader.Read())
                    {
                        dbldN = Math.Round(Convert.ToDouble(dataReader["dN"]), 4);
                        dbldE = Math.Round(Convert.ToDouble(dataReader["dE"]), 4);
                        dbldH = Math.Round(Convert.ToDouble(dataReader["dH"]), 4);
                        dblMeandN += dbldN;
                        dblMeandE += dbldE;
                        dblMeandH += dbldH;
                        iObservationCounter++;
                    }

                    // Close the dataReader
                    if (dataReader != null)
                    {
                        dataReader.Close();
                    }
                }

                catch (System.Data.SqlClient.SqlException ex)
                {
                    Console.WriteLine("getPointDeltas: DB Connection Failed: ");
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }

                finally
                {
                    conn.Dispose();
                    conn.Close();
                }

ComputeMeans:

                if ((strPointID != "Missing") && (iObservationCounter > 0))
                {
                    // Compute the mean dN, dE, dH
                    dblMeandN = Math.Round((dblMeandN / iObservationCounter), 4);
                    dblMeandE = Math.Round((dblMeandE / iObservationCounter), 4);
                    dblMeandH = Math.Round((dblMeandH / iObservationCounter), 4);
                }
                else
                {
                    // allocate false values
                    dblMeandN = 0.0;
                    dblMeandE = 0.0;
                    dblMeandH = 0.0;
                    iObservationCounter = -99;
                }

                //Insert the data into the data arrays
                strDeltas[iCounter, 0] = strPointName;
                strDeltas[iCounter, 1] = Convert.ToString(dblMeandN);
                strDeltas[iCounter, 2] = Convert.ToString(dblMeandE);
                strDeltas[iCounter, 3] = Convert.ToString(dblMeandH);
                strDeltas[iCounter, 4] = Convert.ToString(iObservationCounter);
                iCounter++;
            } while (strPointName != "NoMore");

            strDeltas[iCounter, 0] = "NoMore";
            strDeltas[iCounter, 1] = "999";
            strDeltas[iCounter, 2] = "999";
            strDeltas[iCounter, 3] = "999";
            strDeltas[iCounter, 4] = "0";

            return strDeltas;
        }

        public Tuple<double, double> extractAverageDistance(string strDBconnection, string strTimeBlockStart, string strTimeBlockEnd, string strSensorName)
        {
            //
            // Purpose:
            //      To extract the average raw distance from dbo.TMTDistance table for the time block strTimeBlockStart to strTimeBlockEnd
            // Input:
            //      Receives 
            //          the target name
            //          the start and end time blocks
            // Output:
            //      Returns tuple<double,double> <average distance, no of distances used> 
            // Useage:
            //              var answer = gna.extractAverageDistance(strDBconnection, strTimeBlockStart, strTimeBlockEnd, strSensorName);
            //              double dblAverageDistance = answer.Item1;
            //              double dblNoOfElements = answer.Item2;  // this is actually an integer
            // Comment:
            //      If missing then 
            //          average distance = 0.0
            //          no of elements = -99     
            //

            double dblDistanceCounter = 0.0;
            double dblAverageDistance = 0.0;
            double dblRawDistance = 0.0;

            // Connection and Reader declared outside the try block

            //instantiate and open connection
            SqlConnection conn = new SqlConnection(strDBconnection);
            conn.Open();

            try
            {

                string SQLaction = @"
                    SELECT 
                    [TMTDistance].[RawDistance],
                    [TMTDistance].[EndTimeUTC]
                    FROM [TMTDistance]
                    INNER JOIN [TMCSensor]
                    ON [TMTDistance].[SensorID] = [TMCSensor].[ID]
                    AND [TMCSensor].[Name]= @sensorName
                    AND [TMTDistance].[IsOutlier] = 0
                    AND EndTimeUTC BETWEEN " + strTimeBlockStart + " AND " + strTimeBlockEnd;

                SqlCommand cmd = new SqlCommand(SQLaction, conn);

                // define the parameter used in the command object and add to the command
                // I use this form as I am including a Unicode character in the Select statement

                var par = new SqlParameter("@sensorName", System.Data.SqlDbType.NVarChar)
                {
                    Value = strSensorName
                };
                cmd.Parameters.Add(par);

                // Define the data reader
                SqlDataReader dataReader = cmd.ExecuteReader();

                // get the values
                while (dataReader.Read())
                {
                    dblRawDistance = (double)dataReader["RawDistance"];
                    dblAverageDistance = dblAverageDistance + dblRawDistance;
                    dblDistanceCounter++;
                }

                // Close the dataReader
                if (dataReader != null)
                {
                    dataReader.Close();
                }
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                Console.WriteLine("extractAverageDistance: DB Connection Failed: ");
                Console.WriteLine(ex);
                Console.ReadKey();
            }

            finally
            {
                conn.Dispose();
                conn.Close();
            }

            if (dblDistanceCounter > 0)
            {
                dblAverageDistance = Math.Round((dblAverageDistance / dblDistanceCounter), 4);
            }
            else
            {
                dblAverageDistance = 0.0;
                dblDistanceCounter = -99.0;
            }

            // Console.WriteLine("extractAverageDistance: "+strSensorName+" - "+ Convert.ToString(dblAverageDistance));

            return new Tuple<double, double>(dblAverageDistance, dblDistanceCounter);

        }


        //=================[ Excel methods ]=============================================================================================

        public void checkWorksheetExists(string strExcelWorkbookFullPath, string strWorksheet)
        {
            // To check if the workbook exists
            // To check if the worksheet exists
            // If not, then it adds the worksheet

            try {
                FileInfo testFile = new FileInfo(strExcelWorkbookFullPath);
            }
            catch
            {
                Console.WriteLine(strExcelWorkbookFullPath + " does not exist");
                Console.ReadLine();
            }



            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the workbook if it does not exist
                try
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(strWorksheet);
                    worksheet.Calculate();
                    package.Save();
                }
                catch
                {
                    // it exists so do nothing
                }

            }


        }


        public string[] getPointNames(string strExcelWorkbookFullPath, string strReferenceWorksheet, string strFirstDataRow)
        {
            //
            // Purpose:
            //      To read the point names from the reference worksheet
            // Input:
            //      Full path and name of workbook
            //      Name of reference worksheet
            //      First row of point names
            // Output:
            //      Returns array [PointNames]
            // Useage:
            //      string[] strPointNames = gna.getPointNames(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow);
            //



            // open the existing workbook
            var fiExcelSpreadsheet = new FileInfo(strExcelWorkbookFullPath);

            // navigate through the worksheets
            using (var package = new ExcelPackage(fiExcelSpreadsheet))
            {

                string strActiveWorksheet = strReferenceWorksheet;
                var worksheet = package.Workbook.Worksheets[strActiveWorksheet];

                int iRow = Convert.ToInt16(strFirstDataRow);
                int iCol = 2;
                string[] strPointName = new string[2000];
                int i = 0;
                string strName;

                do
                {
                    // Read in the point names
                    strPointName[i] = Convert.ToString(worksheet.Cells[iRow, iCol].Value);
                    strName = strPointName[i];
                    iRow++;
                    i++;
                } while (strName != "");

                strPointName[i - 1] = "NoMore";

                return strPointName;

            }

        }



        public void writeEndTimeStamp(string strExcelWorkbookFullPath, string strReferenceWorksheet, string strTimeStamp, string strFirstDataRow, string[,] strNamesID)
        {
            //
            // Purpose:
            //      To write the timestamp & point ID to the reference worksheet in the master worksheet for the coordinate export function
            // Input:
            //      The master workbook starts with the point names
            //      The reference coordinates are extracted and added to the worksheet
            //      The last time of the block used to create the reference values gets appended with this routine as the start time
            //      The worksheet then gets used when generating the export data
            // Output:
            //      Reference worksheet ready for use in the coordinate export software
            // Useage:
            //      gna.writeEndTimeStamp(strMasterWorkbookFullPath, strReferenceWorksheet, strRefBlockEnd, strFirstDataRow);
            //

            // Write the reference coordinates to the workbook
            FileInfo masterWorkbook = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(masterWorkbook))
            {

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];
                // Find the last data row
                int iLastRow = Convert.ToInt16(strFirstDataRow);
                string strPointName = "blank";
                do
                {
                    // Read in the point names
                    strPointName = Convert.ToString(namedWorksheet.Cells[iLastRow, 2].Value);
                    iLastRow++;
                } while (strPointName != "");

                iLastRow = iLastRow - 1;

                int iFirstRow = Convert.ToInt16(strFirstDataRow);

                // Now write the time stamp to column 6 
                for (int iRow = iFirstRow; iRow < iLastRow; iRow++)
                {
                    int i = iRow - iFirstRow;
                    namedWorksheet.Cells[iRow, 6].Value = strTimeStamp;
                }
                //Set the column width & alignment
                double columnWidth = 20;
                namedWorksheet.Column(6).Width = columnWidth;
                namedWorksheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedWorksheet.Calculate();
                package.Save();
            }

        }





        public void writeReferenceCoordinates(string strExcelWorkbookFullPath, string strReferenceWorksheet, string[,] strReferenceCoordinates, string strFirstDataRow, string strCoordinateOrder)
        {
            //
            // Purpose:
            //      To write the reference coordinates to the reference worksheet
            // Input:
            //      Full path and name of workbook
            //      Name of reference worksheet
            //      Array of reference coordinates
            //      First data row
            // Output:
            //      --
            // Useage:
            //      gna.writeReferenceCoordinates(strExcelWorkbookFullPath, strReferenceWorksheet, strReferenceCoordinates, strFirstDataRow, strCoordinateOrder);
            //


            string strName = "";
            int iCounter = 0;

            // Write the reference coordinates to the workbook
            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];

                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);

                do
                {
                    strName = (strReferenceCoordinates[iCounter, 0]).Trim();

                    if (strName == "NoMore") goto nextAction;
                    if (strName == "") goto nextAction;

                    namedWorksheet.Cells[iRow, 1].Value = iCounter;
                    namedWorksheet.Cells[iRow, 2].Value = strReferenceCoordinates[iCounter, 0];

                    if (strCoordinateOrder == "ENH") 
                    {
                    namedWorksheet.Cells[iRow, 3].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 2]);
                    namedWorksheet.Cells[iRow, 4].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 1]);
                    namedWorksheet.Cells[iRow, 5].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 3]);
                    }
                    else
                    {
                        namedWorksheet.Cells[iRow, 3].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 1]);
                        namedWorksheet.Cells[iRow, 4].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 2]);
                        namedWorksheet.Cells[iRow, 5].Value = Convert.ToDouble(strReferenceCoordinates[iCounter, 3]);
                    }
                    namedWorksheet.Cells[iRow, 7].Value = "Monitoring Prism";
                    iCounter++;
                    iRow++;
                } while (strName != "NoMore");
nextAction:

                namedWorksheet.Calculate();
                package.Save();
                formatSpecificCells(strExcelWorkbookFullPath, strReferenceWorksheet, 2, 2, 2, iRow - 1);
            }

            return;
        }

        public void writeReferenceDeltas(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, string[,] strPointDeltas, string strTimeBlockStart, string strTimeBlockEnd, string strCoordinateOrder)
        {
            //
            // Purpose:
            //      To write the reference deltas to the master workbook, creating the spreadsheets as needed
            // Input:
            //      As in link
            //      Name of reference worksheet
            //      First row of point names
            // Output:
            //      a workbook that can then be used as a master workbook. 
            // Useage:
            //      gna.writeReferenceDeltas(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strPointDeltas, strRefBlockStart, strRefBlockEnd);
            //

            int iColumn = 6;
            int iLastRow = 0;

            writeDeltas(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strCoordinateOrder);
            Console.WriteLine("Reference Deltas written");

            string strRequiredDelta = "";

            if (strCoordinateOrder=="ENH") 
            {
                iColumn = 3;
                strRequiredDelta = "dE";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dE written");
                strRequiredDelta = "dN";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dN written");
                strRequiredDelta = "dH";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dH written");
            }
            else
            {
                iColumn = 3;
                strRequiredDelta = "dN";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dN written");
                strRequiredDelta = "dE";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dE written");
                strRequiredDelta = "dH";
                iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
                Console.WriteLine("dH written");
            }

            strWorksheet = "d2D";
            iLastRow = writeD2D(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd);
            Console.WriteLine("d2D written");

            strWorksheet = "d(Radial)";
            strRequiredDelta = "dR";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dR written");
            strWorksheet = "d(Tangential)";
            strRequiredDelta = "dT";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dT written");
        }



        public void writeTimeBlockDeltas(string strExcelWorkbookFullPath, string strFirstDataRow, string strColumn, string[,] strPointDeltas, string strTimeBlockStart, string strTimeBlockEnd, string strCoordinateOrder)
        {
            //
            // Purpose:
            //      To write the deltas for a time block to the daily workbook, creating the spreadsheets as needed
            // Input:
            //      As in link
            //      Name of reference worksheet
            //      First row of point names
            // Output:
            //      Updated daily workbook
            // Useage:
            //      gna.writeTimeBlockDeltas(strExcelWorkbookFullPath, strFirstDataRow, strColumn, strPointDeltas, strRefBlockStart, strRefBlockEnd);
            //

            int iColumn = Convert.ToInt32(strColumn);
            int iLastRow = 0;

            string strRequiredDelta = "dN";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dN written");

            strRequiredDelta = "dE";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dE written");

            strRequiredDelta = "dH";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dH written");

            string strWorksheet = "d2D";
            iLastRow = writeD2D(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd);
            Console.WriteLine("d2D written");

            strWorksheet = "d(Radial)";
            strRequiredDelta = "dR";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dR written");

            strWorksheet = "d(Tangential)";
            strRequiredDelta = "dT";
            iLastRow = writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta, strCoordinateOrder);
            Console.WriteLine("dT written");
        }










        public void writeDeltas_NotUsed(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, string[,] strPointDeltas, int iColumn, string strTimeBlockStart, string strTimeBlockEnd, string strCoordinateOrder)
        {


            //
            // Purpose:
            //      To write the deltas to a defined worksheet
            // Deltas
            //      strPointDeltas[iCounter, 0] = strPointName;
            //      strPointDeltas[iCounter, 1] = MeandN
            //      strPointDeltas[iCounter, 2] = MeandE
            //      strPointDeltas[iCounter, 3] = MeandH
            //      strPointDeltas[iCounter, 4] = MeandR
            //      strPointDeltas[iCounter, 5] = MeandT
            //      strPointDeltas[iCounter, 6] = ObservationCounter = "-99" id there are no observations
            //      Full path and name of workbook
            //      Name of worksheet
            //      First data row
            //      Array of deltas
            //      First data column
            //      Time block start
            //      Time block end
            // Output:
            //      --
            // Useage:
            //      gna.writeDeltas(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd);
            // Comment
            //      This can write the deltas to any column - to make provision for a rolling time window of multiple time blocks
            //      The column number gets generated at the same time as the time block (generateTimeBlockStartEnd)
            //

            string strName = "";
            int iCounter = 0;

            // Write the Deltas to the workbook
            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                // add a new worksheet to the workbook if it does not exist
                try
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(strWorksheet);
                }
                catch
                {
                    // it exists so do nothing
                }
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];

                //formatFirstRow(strExcelWorkbookFullPath, strWorksheet);

                //using (var range = namedWorksheet.Cells["A2:B500, L2:L2500"])
                //{
                //    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //}

                //using (var range = namedWorksheet.Cells["M2:M2500"])
                //{
                //    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //}

                //using (var range = namedWorksheet.Cells["C2:T2500"])
                //{
                //    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}

                ////Set the column width
                //double columnWidth = 20;
                //namedWorksheet.Column(1).Width = columnWidth;
                //namedWorksheet.Column(iColumn).Width = columnWidth;
                //namedWorksheet.Column(iColumn + 1).Width = columnWidth;
                //namedWorksheet.Column(iColumn + 2).Width = columnWidth;
                //columnWidth = 20;

                ////Set the column Headers
                //if (strWorksheet == "Reference")
                //{
                //    namedWorksheet.Cells[1, 6].Value = "dE(now)";
                //    namedWorksheet.Cells[1, 7].Value = "dN(now)";
                //    namedWorksheet.Cells[1, 8].Value = "dH(now)";
                //    namedWorksheet.Cells[1, 9].Value = "E(now)";
                //    namedWorksheet.Cells[1, 10].Value = "N(now)";
                //    namedWorksheet.Cells[1, 11].Value = "H(now)";
                //    namedWorksheet.Cells[1, 13].Value = "Reading Count";
                //}
                //else
                //{
                //    namedWorksheet.Cells[1, iColumn].Value = "dE";
                //    namedWorksheet.Cells[1, iColumn + 1].Value = "dN";
                //    namedWorksheet.Cells[1, iColumn + 2].Value = "dH";
                //    namedWorksheet.Cells[1, iColumn + 7].Value = "Reading Count";
                //}

                ////Format number with 3 decimal places
                //namedWorksheet.Cells["A2:A2500"].Style.Numberformat.Format = "0";
                //namedWorksheet.Cells["C2:K2500"].Style.Numberformat.Format = "0.000";
                //namedWorksheet.Column(iColumn + 7).Style.Numberformat.Format = "0";

                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);

                do
                {
                    strName = strPointDeltas[iCounter, 0];
                    string strStatus = strPointDeltas[iCounter, 6];

                    if (strName != "NoMore")
                    {
                        if (strStatus == "-99")
                        {
                            namedWorksheet.Cells[iRow, 2].Value = strPointDeltas[iCounter, 0];
                            namedWorksheet.Cells[iRow, iColumn].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 1].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 2].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 7].Value = "Missing";
                            //namedWorksheet.Cells[iRow, iColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //namedWorksheet.Cells[iRow, iColumn].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                            //namedWorksheet.Cells[iRow, iColumn + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //namedWorksheet.Cells[iRow, iColumn + 1].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                            //namedWorksheet.Cells[iRow, iColumn + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //namedWorksheet.Cells[iRow, iColumn + 2].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                            namedWorksheet.Cells[iRow, iColumn].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 1].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 2].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 7].Style.Font.Italic = true;
                        }
                        else
                        {
                            namedWorksheet.Cells[iRow, iColumn].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 1].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 2].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 7].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, 2].Value = strPointDeltas[iCounter, 0]; //Name
                            namedWorksheet.Cells[iRow, iColumn].Value = Convert.ToDouble(strPointDeltas[iCounter, 2]); //dN
                            namedWorksheet.Cells[iRow, iColumn + 1].Value = Convert.ToDouble(strPointDeltas[iCounter, 1]); //dE
                            namedWorksheet.Cells[iRow, iColumn + 2].Value = Convert.ToDouble(strPointDeltas[iCounter, 3]); //dH
                            namedWorksheet.Cells[iRow, iColumn + 7].Value = Convert.ToInt16(strPointDeltas[iCounter, 6]);

                            if (strWorksheet == "Reference")
                            {
                                string strFormula = "=(C" + Convert.ToString(iRow) + ")+(F" + Convert.ToString(iRow) + ")";
                                namedWorksheet.Cells[iRow, 9].Formula = strFormula;
                                strFormula = "=(D" + Convert.ToString(iRow) + ")+(G" + Convert.ToString(iRow) + ")";
                                namedWorksheet.Cells[iRow, 10].Formula = strFormula;
                                strFormula = "=(E" + Convert.ToString(iRow) + ")+(H" + Convert.ToString(iRow) + ")";
                                namedWorksheet.Cells[iRow, 11].Formula = strFormula;
                            }
                        }
                        iCounter++;
                        iRow++;
                    }

                } while (strName != "NoMore");

                namedWorksheet.Cells[iRow + 1, iColumn + 1].Value = strTimeBlockStart;
                namedWorksheet.Cells[iRow + 2, iColumn + 1].Value = strTimeBlockEnd;

                namedWorksheet.Calculate();
                package.Save();

                if (strWorksheet == "Reference")
                {
                    drawBox(strExcelWorkbookFullPath, strWorksheet, 9, 11, 2, iRow - 1);
                }


            }

            return;
        }

        public int writeD2D(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, string[,] strPointDeltas, int iColumn, string strTimeBlockStart, string strTimeBlockEnd)
        {
            //
            // Purpose:
            //      To write the mean dS to the worksheet for a time series
            // Input:
            //      Full path and name of workbook
            //      Name of worksheet
            //      First data row
            //      Array of deltas
            //      First data column
            //      Time block start
            //      Time block end
            // Output:
            //      --
            // Useage:
            //      gna.writeDh(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd);
            // Comment
            //      This can write the dH to any column - to make provision for a rolling time window of multiple time blocks
            //      The column number gets generated at the same time as the time block (generateTimeBlockStartEnd)
            //


            string strName = "";
            int iCounter = 0;
            double dE, dN, dS;

            checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);
            //formatFirstRow(strExcelWorkbookFullPath, strWorksheet);

            // Write the Deltas to the workbook
            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                // add a new worksheet to the workbook if it does not exist
                try
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(strWorksheet);
                }
                catch
                {
                    // it exists so do nothing
                }
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];

                //formatFirstRow(strExcelWorkbookFullPath, strWorksheet);

                //Set the column width
                double columnWidth = 20;
                namedWorksheet.DefaultColWidth = columnWidth;

                //Set the column Headers
                namedWorksheet.Cells[1, iColumn].Value = "d(Hor)";

                //Format number with 3 decimal places
                namedWorksheet.Cells["A2:A2500"].Style.Numberformat.Format = "0";
                namedWorksheet.Cells["C2:K2500"].Style.Numberformat.Format = "0.000";

                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);
                do
                {
                    strName = strPointDeltas[iCounter, 0];
                    if (strName == "NoMore") goto nextAction;

                    string strStatus = strPointDeltas[iCounter, 6];
                    if (strStatus == "-99")
                    {
                        namedWorksheet.Cells[iRow, iColumn].Value = "Missing";
                    }
                    else
                    {
                        namedWorksheet.Cells[iRow, 1].Value = iCounter;
                        namedWorksheet.Cells[iRow, 2].Value = strName;
                        dE = Convert.ToDouble(strPointDeltas[iCounter, 2]);
                        dN = Convert.ToDouble(strPointDeltas[iCounter, 1]);
                        dS = Math.Round((Math.Pow((dE * dE) + (dN * dN), 0.5)), 3);
                        namedWorksheet.Cells[iRow, iColumn].Value = dS;
                    }

                    iCounter++;
                    iRow++;
                } while (strName != "NoMore");

nextAction:
                int iLastRow = iRow;
                namedWorksheet.Cells[iRow + 1, iColumn].Value = strTimeBlockStart;
                namedWorksheet.Cells[iRow + 2, iColumn].Value = strTimeBlockEnd;
                namedWorksheet.Calculate();
                package.Save();
                return iLastRow;

            }
        }

        public string[] getROarray(string strExcelWorkbookFullPath, string strCalibrationWorksheet, string strFirstDataRow)
        {

            //
            // Purpose:
            //      extract the RO details ("ATS1 → CP02") from the calibration spreadsheet
            // Input:
            //      Full path and name of workbook
            //      Name of calibration worksheet
            //      First data row in the calibration worksheet
            // Output:
            //      strROdistances["ATS1 → CP02;reference distance"]
            // Useage:
            //      string[] strROdistances = gna.getROarray(strExcelWorkbookFullPath, strCalibrationWorksheet, strFirstOutputRow);
            //

            int iStart = Convert.ToInt32(strFirstDataRow);
            int iRow;
            int iCol = 1;
            int iROCounter = -1;
            String[] strROdistances = new String[50];

            FileInfo excelWorkbook = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strCalibrationWorksheet];
                // Read in the RO reference distance names, starting at the first data row.
                for (iRow = iStart; iRow < 50; iRow++)
                {
                    string strText = Convert.ToString(namedWorksheet.Cells[iRow, iCol].Value);
                    string strReferenceDistance = Convert.ToString(namedWorksheet.Cells[iRow, iCol + 1].Value);
                    if (strText == "") goto nextStep;
                    iROCounter++;
                    strROdistances[iROCounter] = strText + ";" + strReferenceDistance;
                }
            }
nextStep:
            iROCounter++;
            strROdistances[iROCounter] = "NoMore";
            return strROdistances;
        }

        public void populateCalibrationWorksheet(string strDBconnection, string strTimeBlockStart, string strTimeBlockEnd, string strExcelWorkbookFullPath, string strCalibrationWorksheet, string strFirstOutputRow, string strDistanceColumn)
        {
            //
            // Purpose:
            //      To populate the calibration worksheet a time block
            // Input:
            //      Full path and name of workbook
            //      Name of worksheet
            //      First data row
            //      Array of deltas
            //      First data column
            //      Time block start
            //      Time block end
            // Output:
            //      a populated Calibration worksheet
            // Useage:
            //      gna.populateCalibrationWorksheet(strDBconnection, strTimeBlockStart, strTimeBlockEnd, strExcelWorkbookFullPath, strCalibrationWorksheet, strFirstOutputRow, strDistanceColumn)
            //

            String[] strROmeanDistances = new String[50];
            String[] strRO1 = new String[50];

            string[] strROdistances = getROarray(strExcelWorkbookFullPath, strCalibrationWorksheet, strFirstOutputRow); // a string containing Descriptor;Distance
            int iFirstRow = Convert.ToInt16(strFirstOutputRow);
            int iColumn = Convert.ToInt16(strDistanceColumn);
            int iCounter = 0;
            string strROtarget = "";
            do
            {
                strROtarget = strROdistances[iCounter];
                strRO1 = strROtarget.Split(';');
                strROtarget = strRO1[0];

                // Now extract the average distance between ATS & RO for the time block 
                var answer = extractAverageDistance(strDBconnection, strTimeBlockStart, strTimeBlockEnd, strROtarget);
                double dblAverageDistance = answer.Item1;
                strROmeanDistances[iCounter] = strROtarget + ";" + Convert.ToString(dblAverageDistance);
                iCounter++;
            } while (strROdistances[iCounter] != "NoMore");

            strROmeanDistances[iCounter] = "NoMore";

            // Write the observations to the  working copy of the workbook
            FileInfo workingWorkbook = new FileInfo(@strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(workingWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strCalibrationWorksheet];

                // Format the cells
                namedWorksheet.Cells["E7:E8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedWorksheet.Cells["E7:E8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedWorksheet.Cells["B7:D50"].Style.Numberformat.Format = "0.000";

                // First write the data/time and the data time window
                namedWorksheet.Cells[7, 5].Value = strTimeBlockStart;
                namedWorksheet.Cells[8, 5].Value = strTimeBlockEnd;

                iCounter = 0;
                strROtarget = strROmeanDistances[iCounter];
                int iRow = Convert.ToInt16(strFirstOutputRow);

                do
                {
                    // Write the meaned distances
                    strROtarget = strROmeanDistances[iCounter];

                    if (strROtarget != "blank")
                    {
                        // Now apply conditional formatting

                        strRO1 = strROtarget.Split(';');
                        strROtarget = strRO1[0];
                        string strAverageDistance = strRO1[1];
                        namedWorksheet.Cells[iRow, iColumn].Value = Convert.ToDouble(strAverageDistance);
                        string strFormula = "=(C" + Convert.ToString(iRow) + ")-(B" + Convert.ToString(iRow) + ")";
                        namedWorksheet.Cells[iRow, 4].Formula = strFormula;
                        iCounter++;
                        iRow++;
                    }
                } while (strROmeanDistances[iCounter] != "NoMore");

                namedWorksheet.Calculate();
                package.Save();
            }

        }

        public int writeSpecificDelta(string strExcelWorkbookFullPath, string strFirstDataRow, string[,] strPointDeltas, int iColumn, string strTimeBlockStart, string strTimeBlockEnd, string strRequiredDelta, string strCoordinateOrder)
        {
            //
            // Purpose:
            //      To write a specific mean delta strRequiredDelta: (dN, dE, dHt, dR, dT) to a specific worksheet in a specific column for a time window, 
            //      less the reference value
            // Output:
            //      iLastRow - to be used for specific formatting
            // Useage:
            //      gna.writeSpecificDelta(strExcelWorkbookFullPath, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strRequiredDelta);
            // Comment
            //      This can write the dH to any column - to make provision for a rolling time window of multiple time blocks
            //      The column number gets generated at the same time as the time block (generateTimeBlockStartEnd)
            //

//*************   THIS NEEDS UPDATING TO ADDRESS COORDINATE ORDER *********


            string strName = "";
            int iCounter;
            int iElement = 0;
            string strSpecificWorksheet = "";
            string strHeader = "";

            // Identify the required Delta
            //    0 Target
            //    1 dN
            //    2 dE
            //    3 dHt
            //    4 dR
            //    5 dT
            //    6 Counter

            if (strCoordinateOrder == "ENH")
            {
                switch (strRequiredDelta)
                {
                    case "dN":
                        iElement = 2;
                        strSpecificWorksheet = "dN";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dE":
                        iElement = 1;
                        strSpecificWorksheet = "dE";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dH":
                        iElement = 3;
                        strSpecificWorksheet = "dH";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dR":
                        iElement = 4;
                        strSpecificWorksheet = "d(Radial)";
                        strHeader = "d(Radial)";
                        break;
                    case "dT":
                        iElement = 5;
                        strSpecificWorksheet = "d(Tangential)";
                        strHeader = "d(Tangential)";
                        break;
                }

            }
            else {
                switch (strRequiredDelta)
                {
                    case "dN":
                        iElement = 1;
                        strSpecificWorksheet = "dN";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dE":
                        iElement = 2;
                        strSpecificWorksheet = "dE";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dH":
                        iElement = 3;
                        strSpecificWorksheet = "dH";
                        strHeader = strSpecificWorksheet;
                        break;
                    case "dR":
                        iElement = 4;
                        strSpecificWorksheet = "d(Radial)";
                        strHeader = "d(Radial)";
                        break;
                    case "dT":
                        iElement = 5;
                        strSpecificWorksheet = "d(Tangential)";
                        strHeader = "d(Tangential)";
                        break;
                }


            }

            checkWorksheetExists(strExcelWorkbookFullPath, strSpecificWorksheet);
            //formatFirstRow(strExcelWorkbookFullPath, strSpecificWorksheet);

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strSpecificWorksheet];

                //Set the column width
                double columnWidth = 20;
                namedWorksheet.DefaultColWidth = columnWidth;

                //Set the column Header
                namedWorksheet.Cells[1, iColumn].Value = strHeader;

                //Format number with 3 decimal places
                namedWorksheet.Cells["A2:A2500"].Style.Numberformat.Format = "0";
                namedWorksheet.Cells["C2:K2500"].Style.Numberformat.Format = "0.000";
                namedWorksheet.Cells["C2:K2500"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);
                do
                {
                    strName = strPointDeltas[iCounter, 0];

                    if (strName == "NoMore") goto NextAction;
                    namedWorksheet.Cells[iRow, 1].Value = iCounter;
                    namedWorksheet.Cells[iRow, 2].Value = strName;

                    string strFormula = "=" + strPointDeltas[iCounter, iElement] + "-(C" + Convert.ToString(iRow) + ")";
                    namedWorksheet.Cells[iRow, iColumn].Formula = strFormula;

                    string strStatus = strPointDeltas[iCounter, 6];
                    if (strStatus == "-99")
                    {
                        namedWorksheet.Cells[iRow, iColumn].Value = "Missing";
                    }

                    iCounter++;
                    iRow++;
                } while (strName != "NoMore");
NextAction:
                int iLastRow = iRow;
                namedWorksheet.Cells[iRow + 1, iColumn].Value = strTimeBlockStart;
                namedWorksheet.Cells[iRow + 2, iColumn].Value = strTimeBlockEnd;

                namedWorksheet.Calculate();
                package.Save();
                return iLastRow;
            }
        }

        public void formatSpecificCells(string strExcelWorkbookFullPath, string strWorksheet, int iLeftCol, int iRightCol, int iTopRow, int iBottomRow)
        {
            //
            // Purpose:
            //      To format a block of cells in a specified worksheet
            // Useage:
            //      gna.formatSpecificCells(strExcelWorkbookFullPath, strWorksheet, iLeftCol, iRightCol, iTopRow, int iBottomRow);
            //

            // Define the workbook
            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the workbook if it does not exist
                checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];

                using (var range = namedWorksheet.Cells[iTopRow, iLeftCol, iBottomRow, iRightCol])
                {
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    range.Style.Font.Bold = false;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    range.Style.Font.Color.SetAuto();
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                namedWorksheet.Calculate();
                package.Save();
            }
        }

        public void formatFirstRowCrdExport(string strExcelWorkbookFullPath, string strWorksheet, string strCoordinateOrder)
        {
            checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
                using (var range = namedWorksheet.Cells["A1:ZZ1"])
                {
                    double rowHeight = 35;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                    //range.Style.Font.Color.SetAuto();
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.WrapText = true;
                    namedWorksheet.Row(1).Height = rowHeight;
                    namedWorksheet.Row(1).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                // Insert the Column headers

                if (strCoordinateOrder == "ENH")
                {
                    namedWorksheet.Cells[1, 1].Value = "SensorID";
                    namedWorksheet.Cells[1, 2].Value = "Target";
                    namedWorksheet.Cells[1, 3].Value = "E(ref)";
                    namedWorksheet.Cells[1, 4].Value = "N(ref)";
                    namedWorksheet.Cells[1, 5].Value = "H(ref)";
                    namedWorksheet.Cells[1, 6].Value = "LastReading";
                    namedWorksheet.Cells[1, 7].Value = "Type";
                }
                else
                {
                    namedWorksheet.Cells[1, 1].Value = "SensorID";
                    namedWorksheet.Cells[1, 2].Value = "Target";
                    namedWorksheet.Cells[1, 3].Value = "N(ref)";
                    namedWorksheet.Cells[1, 4].Value = "E(ref)";
                    namedWorksheet.Cells[1, 5].Value = "H(ref)";
                    namedWorksheet.Cells[1, 6].Value = "LastReading";
                    namedWorksheet.Cells[1, 7].Value = "Type";
                }

                namedWorksheet.Calculate();
                package.Save();
            }
        }









        //public void formatFirstRow(string strExcelWorkbookFullPath, string strWorksheet)
        //{
        //    checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

        //    FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

        //    using (ExcelPackage package = new ExcelPackage(newFile))
        //    {
        //        ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
        //        using (var range = namedWorksheet.Cells["A1:ZZ1"])
        //        {
        //            double rowHeight = 35;
        //            range.Style.Font.Bold = true;
        //            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //            range.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
        //            //range.Style.Font.Color.SetAuto();
        //            range.Style.Font.Color.SetColor(Color.White);
        //            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            range.Style.WrapText = true;
        //            namedWorksheet.Row(1).Height = rowHeight;
        //            namedWorksheet.Row(1).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //        }

        //        // Insert the Column headers
        //        namedWorksheet.Cells[1, 1].Value = "ID";
        //        namedWorksheet.Cells[1, 2].Value = "Prism";
        //        namedWorksheet.Cells[1, 3].Value = "E(reference)";
        //        namedWorksheet.Cells[1, 4].Value = "N(reference)";
        //        namedWorksheet.Cells[1, 5].Value = "H(reference)";
        //        namedWorksheet.Cells[1, 6].Value = "dE (Now)";
        //        namedWorksheet.Cells[1, 7].Value = "dN (Now)";
        //        namedWorksheet.Cells[1, 8].Value = "dH (Now)";
        //        namedWorksheet.Cells[1, 9].Value = "E(Now)";
        //        namedWorksheet.Cells[1, 10].Value = "N(Now)";
        //        namedWorksheet.Cells[1, 11].Value = "H(Now)";
        //        namedWorksheet.Cells[1, 12].Value = "Name";
        //        namedWorksheet.Cells[1, 13].Value = "Reading Count";
        //        namedWorksheet.Cells[1, 14].Value = "Prism Offset";
        //        namedWorksheet.Cells[1, 15].Value = "Top of Rail";
        //        namedWorksheet.Cells[1, 16].Value = " ";
        //        namedWorksheet.Cells[1, 17].Value = "E(Ref)";
        //        namedWorksheet.Cells[1, 18].Value = "N(Ref)";
        //        namedWorksheet.Cells[1, 19].Value = "H(Ref)";
        //        namedWorksheet.Cells[1, 20].Value = "dS (Base)";
        //        namedWorksheet.Cells[1, 21].Value = "dH (Base)";
        //        namedWorksheet.Cells[1, 22].Value = "dS (Corr)";
        //        namedWorksheet.Cells[1, 23].Value = "dH (Corr)";
        //        namedWorksheet.Cells[1, 24].Value = "dS (Now)";
        //        namedWorksheet.Cells[1, 25].Value = "dH (Now)";
        //        namedWorksheet.Cells[1, 26].Value = "Rail Number";

        //        package.Save();
        //    }

        //    return;
        //}

        public void drawBox(string strExcelWorkbookFullPath, string strWorksheet, int iLeftCol, int iRightCol, int iTopRow, int iBottomRow)
        {
            //
            // Purpose:
            //      To draw a box around a block of cells in a specified worksheet
            // Useage:
            //      gna.drawBox(strExcelWorkbookFullPath, strWorksheet, iLeftCol, iRightCol, iTopRow, iBottomRow);
            //

            checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];

                using (var range = namedWorksheet.Cells[iTopRow, iLeftCol, iBottomRow, iRightCol])
                {

                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                namedWorksheet.Calculate();
                package.Save();
            }

            return;

        }






        //public void worksheetHousekeeping(string strExcelWorkbookFullPath, string strReferenceWorksheet, string strFirstDataRow, string[] strPointNames)
        //{
        //    //
        //    // Purpose:
        //    //      To do formatting etc to the worksheets that has nowhere else to be put
        //    // Input:
        //    //      Full path and name of workbook
        //    //      Name of worksheet
        //    //      First data row
        //    // Output:
        //    //      Correctly formatted headers
        //    //      Correctly formatted ID, Point name and reference columns
        //    // Useage:
        //    //      worksheetHousekeeping(strExcelWorkbookFullPath,strWorksheet,strFirstDataRow);
        //    // Comment
        //    //      This forms part of the prepareReferenceWorksheet module.
        //    //

        //    string strName = "";
        //    int iCounter = 0;

        //    int iTopRow = Convert.ToInt32(strFirstDataRow);

        //    // determine how many points there are
        //    do
        //    {
        //        strName = strPointNames[iCounter];
        //        iCounter++;
        //    } while (strName != "NoMore");

        //    int iBottomRow = iTopRow + iCounter - 2;

        //    string[] strColumnHeader = new string[7];

        //    strColumnHeader[1] = "dN(Ref)";
        //    strColumnHeader[2] = "dE(Ref)";
        //    strColumnHeader[3] = "dH(Ref)";
        //    strColumnHeader[4] = "d2D(Ref)";
        //    strColumnHeader[5] = "dR(Ref)";
        //    strColumnHeader[6] = "dT(Ref)";

        //    string[] strWorksheet = new string[7];
        //    strWorksheet[1] = "dN";
        //    strWorksheet[2] = "dE";
        //    strWorksheet[3] = "dH";
        //    strWorksheet[4] = "d2D";
        //    strWorksheet[5] = "d(Radial)";
        //    strWorksheet[6] = "d(Tangential)";

        //    FileInfo workingWorkbook = new FileInfo(@strExcelWorkbookFullPath);

        //    iCounter = 1;

        //    using (ExcelPackage package = new ExcelPackage(workingWorkbook))
        //    {

        //        do
        //        {
        //            string workingWorksheet = strWorksheet[iCounter];
        //            ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[workingWorksheet];
        //            //Set the column Headers
        //            namedWorksheet.Cells[1, 1].Value = "ID";
        //            namedWorksheet.Cells[1, 2].Value = "Name";
        //            namedWorksheet.Cells[1, 3].Value = strColumnHeader[iCounter];
        //            // Set the Column alignment
        //            using (var range = namedWorksheet.Cells["A1:C2500"])
        //            {
        //                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            }
        //            // Box columns
        //            using (var range = namedWorksheet.Cells[iTopRow, 3, iBottomRow, 3])
        //            {
        //                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        //                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //            }
        //            using (var range = namedWorksheet.Cells[iTopRow, 2, iBottomRow, 2])
        //            {
        //                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
        //                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //            }

        //            iCounter++;
        //        } while (iCounter < 7);

        //        package.Save();
        //    }
        //}

        //=================[ General methods ]=============================================================================================

        public int countDeltas(string[] strPointNames)
        {
            //
            // Purpose:
            //      To count the number of targets = number of deltas
            // Output:
            //      iCounter = number of targets, starting at 1
            // Useage:
            //
            //      iCounter= gna.countDeltas(strPointNames);
            //


            int iCounter = -1;
            string strPointName = "";
            do
            {
                iCounter++;
                strPointName = strPointNames[iCounter];
                Console.WriteLine(Convert.ToString(iCounter) + " : " + strPointName);

            } while (strPointName != "NoMore");

            return iCounter;

        }

        public void prepareReferenceWorkbook(string strDBconnection, string strExcelWorkbookFullPath, string strReferenceWorksheet, string strCalibrationWorksheet, string strFirstDataRow, string strProjectTitle, string strRefBlockStart, string strRefBlockEnd, string strCoordinateOrder)
        {
            //
            // Purpose:
            //      To prepare a reference workbook
            // Input:
            //      the basic workbook must contain the following worksheets:
            //          Reference worksheet: containing a list of all the targets
            //          Calibration worksheet: containing the calibration distances to be created.
            // Output:
            //      The reference workbook, with the various reference values populated
            // Useage:
            //      gna.prepareReferenceWorkbook(strDBconnection, strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strProjectTitle, strRefBlockStart, strRefBlockEnd);
            // Comment:
            //      The flag PrepareReferenceMeasurements = "Yes" then the reference values are extracted and populated
            //      After this has been done once, this flag must be set to "No" and this step is skipped
            //

            Console.WriteLine("Master workbook being prepared...");
            // Format the reference worksheet header line
            //formatFirstRow(strExcelWorkbookFullPath, strReferenceWorksheet);
            // Read in the point names from the reference worksheet.
            Console.WriteLine("Extracting Point Names....");
            string[] strPointNames = getPointNames(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow);
            // Read the PointID for each point.
            Console.WriteLine("Extracting Point ID....");
            string[,] strNamesID = getPointID(strDBconnection, strProjectTitle, strPointNames);
            // Get the reference coordinates
            Console.WriteLine("Extracting Reference Coordinates....");
            string[,] strReferenceCoordinates = getReferenceCoordinates(strDBconnection, strProjectTitle, strNamesID);
            // Write the reference coordinates to the reference worksheet
            Console.WriteLine("Writing Reference Coordinates....");
            writeReferenceCoordinates(strExcelWorkbookFullPath, strReferenceWorksheet, strReferenceCoordinates, strFirstDataRow, strCoordinateOrder);
            // Get the reference Deltas
            Console.WriteLine("Extracting deltas....");
            string[,] strPointDeltas = getPointDeltas(strDBconnection, strProjectTitle, strRefBlockStart, strRefBlockEnd, strNamesID);
            // Write the reference deltas to the various worksheets
            Console.WriteLine("Writing deltas....");
            writeReferenceDeltas(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strPointDeltas, strRefBlockStart, strRefBlockEnd, strCoordinateOrder);
            // Populate the calibration worksheet
            Console.WriteLine("Prepare the Calibration worksheet....");
            populateCalibrationWorksheet(strDBconnection, strRefBlockStart, strRefBlockEnd, strExcelWorkbookFullPath, strCalibrationWorksheet, "8", "2");
            //Console.WriteLine("Finalise Housekeeping....");
            // Finalise housekeeping
            //worksheetHousekeeping(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strPointNames);
            Console.WriteLine(" ");
        }


        //================[ Email Methods ]===============================================================================================

        public void sendEmail(string strEmailLogin, string strEmailPassword, string strEmailFrom, string strEmailRecipients, string strSubjectLine, string strEmailBody)
        {
            //
            // Purpose:
            //      To send a simple email with no attachment
            // Useage:
            //      gna.sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strEmailRecipients, strSubjectLine,strEmailBody);
            //
            // Information
            //      https://www.emailarchitect.net/easendmail/kb/csharp.aspx?cat=3
            //

            try
            {
                SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");

                // SMTP server address
                SmtpServer oServer = new SmtpServer("smtp.gmail.com");
                oServer.User = strEmailLogin;
                oServer.Password = strEmailPassword;
                oServer.ConnectType = SmtpConnectType.ConnectTryTLS;
                oServer.Port = 587;

                //Set sender email address, please change it to yours
                oMail.From = strEmailFrom;
                oMail.To = new AddressCollection(strEmailRecipients);
                oMail.Subject = strSubjectLine;
                oMail.TextBody = strEmailBody;

                //oMail.AddAttachment(strExcelWorkingFile);

                SmtpClient oSmtp = new SmtpClient();
                oSmtp.SendMail(oServer, oMail);
                Console.WriteLine("email sent...");
            }
            catch (Exception ep)
            {
                Console.WriteLine("Failed to send email:");
                Console.WriteLine("");
                Console.WriteLine(strEmailLogin);
                Console.WriteLine(strEmailPassword);
                Console.WriteLine("");
                Console.WriteLine("Possible issues:");
                Console.WriteLine("1. Two step verification must be activated");
                Console.WriteLine("2. Use the correct Apps password");
                Console.WriteLine("3. See the T4D Cookbook for step by step details");
                Console.WriteLine("");
                Console.WriteLine(ep.Message);
                Console.ReadKey();
            }
        }

        //================[ Alarm Methods ]===============================================================================================

        public void updateAlarmFile(string alarmFile, string strAlarmState)
        {
            // Update the Alarm file
            File.Delete(alarmFile);
            using (StreamWriter writetext = new StreamWriter(alarmFile, false))
            {
                writetext.WriteLine(strAlarmState);
                writetext.Close();
            }
        }


        public void noDataAlarm(string strSettopM1, string strProjectName, string strCurrentGKAfolder, string strNoDataInterval, string strEmailLogin, string strEmailPassword, string strEmailFrom, string strNoDataAlarmRecipients, string testEmail)
        {
            //
            // Purpose:
            //      To check the time of the last gka file that was processed and if it was too long ago, fire off a no data alarm. 
            //      This is repeated for 3 alarms 
            //      When a gka file is processed the alarm state gets reset to 0 = OK. 
            // Useage:
            //      gka.noDataAlarm(strProjectName, strCurrentGKAfolder, strNoDataInterval, strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, testEmail);
            // Comment
            //      if testEmail is "Yes" then a test email is sent. The alarm is not evaluated
            //




            string[] element = new string[35];
            string[] element4 = new string[35];

            //string strSettopM1 = getSettopNumber(strCurrentGKAfolder);
            string strFileDateTime = getFileTimestamp(strCurrentGKAfolder);

            DateTime FileDate = DateTime.ParseExact(strFileDateTime, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.UtcNow;

            int hours = (((TodayDate - FileDate).Days) * 24) + (TodayDate - FileDate).Hours;
            int iNoDataInterval = Convert.ToInt16(strNoDataInterval);



            string strAlarmState = "OK";
            string strSubjectLine = strProjectName + ": No Data Alarm (";
            string strEmailBody = "";
            string strSystemState = "";

            // create the alarm file if missing
            string alarmFile = strCurrentGKAfolder + "AlarmFile.txt";
            if (!File.Exists(alarmFile))
            {
                using (StreamWriter writetext = new StreamWriter(alarmFile, false))
                {
                    writetext.WriteLine("0");
                    writetext.Close();
                }
            }

            strSystemState = testEmail;

            // the special case where I am testing the alarm email function
            if (testEmail == "Yes")
            {
                Console.WriteLine("Send test email");

                strSubjectLine = "Test eMail (" + strSettopM1 + ")";
                strEmailBody = "This is a test email from the NoDataAlarm function";
                sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, strSubjectLine, strEmailBody);
                Console.WriteLine("A test email has been sent: the data has not been evaluated for " + strSettopM1);
                Console.ReadKey();
                goto TheEnd;
            }



            strAlarmState = File.ReadAllText(alarmFile).TrimEnd();

            // The system is in "No Alarm" state
            if ((hours <= iNoDataInterval) && (strAlarmState == "0"))
            {
                Console.WriteLine("State: No Alarm");
                strSystemState = "No Alarm";
                strAlarmState = "0";
                strSubjectLine = strProjectName + " (" + strSettopM1 + ") " + ": data being received";
                strEmailBody = "Alarm state OK";

                goto TheEnd;
            }

            // The system is in "Alarm" state, data is now being received, and system is reset to "No Alarm" state
            if ((hours <= iNoDataInterval) && ((strAlarmState == "1") || (strAlarmState == "2") || (strAlarmState == "3")))
            {


                strSystemState = "Alarm Reset";

                Console.WriteLine("State: Alarm reset");

                strAlarmState = "0";
                strSubjectLine = strProjectName + " Alarm reset (" + strSettopM1 + "): data being received";
                strEmailBody = "Alarm state has changed to OK";

                sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, strSubjectLine, strEmailBody);
                updateAlarmFile(alarmFile, strAlarmState);
                updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);
                Console.WriteLine(strSubjectLine);
                goto TheEnd;
            }

            // The system is placed "Alarm" state,and the current state is elevated
            if (hours > iNoDataInterval)
            {
                Console.WriteLine("State: Alarm");

                strSystemState = "Alarm";
                strEmailBody = "No data files received for the past " + hours.ToString() + " hrs in " + strCurrentGKAfolder;
                strSubjectLine = "No data alarm: " + strProjectName + " - alarm(";





                switch (strAlarmState)
                {
                    case "0":
                        strAlarmState = "1";
                        strSubjectLine = strSubjectLine + strAlarmState + ")";
                        strEmailBody = strEmailBody + "\r\n" + "Settop: " + strSettopM1;
                        sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, strSubjectLine, strEmailBody);
                        updateAlarmFile(alarmFile, strAlarmState);
                        updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);
                        Console.WriteLine(strSubjectLine);
                        break;
                    case "1":
                        strAlarmState = "2";
                        strSubjectLine = strSubjectLine + strAlarmState + ")";
                        strEmailBody = strEmailBody + "\r\n" + "Settop: " + strSettopM1;
                        sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, strSubjectLine, strEmailBody);
                        updateAlarmFile(alarmFile, strAlarmState);
                        updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);
                        Console.WriteLine(strSubjectLine);
                        break;
                    case "2":
                        strAlarmState = "3";
                        strSubjectLine = strSubjectLine + strAlarmState + ")";
                        strEmailBody = strEmailBody + "\r\n" + "Settop: " + strSettopM1;
                        sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, strSubjectLine, strEmailBody);
                        updateAlarmFile(alarmFile, strAlarmState);
                        updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);
                        Console.WriteLine(strSubjectLine);
                        break;
                    case "3":
                        strSubjectLine = "Do nothing";
                        strAlarmState = "3";
                        Console.WriteLine("Alarm state still active: " + strSettopM1);
                        break;
                    default:
                        strSubjectLine = "Do nothing";
                        strAlarmState = "3";
                        Console.WriteLine("Alarm state still active: " + strSettopM1);
                        break;
                }
            }

TheEnd:
            Console.WriteLine("NoData alarm evaluated");
            return;
        }


        public void systemStatusEmail(string strProjectTitle, string[] strSettop, string[] strGKAfolder, string strEmailLogin, string strEmailPassword, string strEmailFrom, string strSystemStatusRecipients)
        {
            //
            // This is a daily email that reports, in 1 email, the state of every Settop
            // It also verifies that the alarm function is working
            // The log file for each Settop is updated to confirm that the system status email was sent
            //

            string strFolder = "";
            string strSettopNumber = "";
            string strAlarmState = "";
            string[] strStatus = new string[7];
            int iSettopCounter = 0;
            string strSubjectLine = strProjectTitle + ": System OK";

            for (int i = 1; i < 11; i++)
            {

                strFolder = strGKAfolder[i];
                strSettopNumber = strSettop[i];

                if (strFolder == "None") break;
                string alarmFile = strFolder + "AlarmFile.txt";
                strAlarmState = File.ReadAllText(alarmFile).TrimEnd();
                if (strAlarmState == "0")
                {
                    strAlarmState = "OK";
                    strSubjectLine = strProjectTitle + ": System OK";
                }
                else
                {
                    strAlarmState = "Alarm";
                    strSubjectLine = strProjectTitle + ": System in Alarm state";
                }
                strStatus[i] = strSettopNumber + ": " + strAlarmState;
                iSettopCounter = i;
                updateAlarmLogFile(strGKAfolder[i], "System status email sent");
            }

            string strEmailBody = "\r\n";

            for (int i = 1; i <= iSettopCounter; i++)
            {
                strEmailBody = strEmailBody + strStatus[i] + "\r\n";
            }
            sendEmail(strEmailLogin, strEmailPassword, strEmailFrom, strSystemStatusRecipients, strSubjectLine, strEmailBody);
            Console.WriteLine("Status email sent");
            //Console.ReadKey();
        }

        public void updateAlarmLogFile(string strCurrentGKAfolder, string strSubjectLine)
        {
            //
            // to append the status line of every email to the Alarm log
            //  useage: updateAlarmLogFile(strCurrentGKAfolder, strSubjectLine);

            // create the alarm Log file if missing
            string alarmLog = strCurrentGKAfolder + "AlarmLog.txt";
            if (!File.Exists(alarmLog))
            {
                using (StreamWriter writetext = new StreamWriter(alarmLog, false))
                {
                    writetext.WriteLine("Alarm Log File (UTC time)");
                    writetext.Close();
                }
            }

            // Get the date/time stamp
            string strNow = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") + " : ";

            if (strSubjectLine != "Do nothing")
            {
                using (StreamWriter writetext = File.AppendText(alarmLog))
                {
                    writetext.WriteLine(strNow + strSubjectLine);
                    writetext.Close();
                }
            }
        }

        //================[ General Methods ]===============================================================================================

        public string getSettopNumber(string strInputString)
        {
            // 
            // Receive a string (normally a file location etc) and search for the settop serial number that is part of that string
            // it works for M1.. and for m1..
            // Return the Settop serial number
            //

            string strSettopSerialNumber = "";

            string pattern = @"M1";

            Match m = Regex.Match(strInputString, pattern, RegexOptions.IgnoreCase);
            if (m.Success)
            {
                // Console.WriteLine("Found '{0}' at position {1}.", m.Value, m.Index);
                strSettopSerialNumber = strInputString.Substring(m.Index, 9);
            }
            else
            {
                // Console.WriteLine("Not found '{0}'", pattern);
                strSettopSerialNumber = "M1xxxxx";
            }
            return strSettopSerialNumber;
        }


        public string getFileTimestamp(string strCurrentGKAfolder)
        {

            // this returns the file timestamp from the Settop latestFile file

            string strLatestFile = strCurrentGKAfolder + "latestFile";

            if (!File.Exists(strLatestFile))
                throw new Exception("File not found: " + strLatestFile);

            string strFileTimestamp;
            using (FileStream fs = new FileStream(strLatestFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (StreamReader sr = new StreamReader(fs))
                {
                    strFileTimestamp = sr.ReadToEnd().Trim(new char[] { '\r', '\n', ' ' });
                }
            }

            // Now reformat this 
            string YYYY = strFileTimestamp.Substring(0, 4);
            string MM = strFileTimestamp.Substring(5, 2);
            string dd = strFileTimestamp.Substring(8, 2);
            string HH = strFileTimestamp.Substring(11, 2);
            string mm = strFileTimestamp.Substring(17, 2);
            string ss = strFileTimestamp.Substring(20, 2);
            strFileTimestamp = YYYY + MM + dd + HH + mm + ss;
            return strFileTimestamp;
        }


        public string formatTimestamp(string strTimestamp)
        {
            // 
            // Receives 19/03/2021 11:43:17
            // Returns 2021-03-19 11:43:17
            // 
            // strTimeStamp = gna.formatTimestamp(strTimeStamp);
            //

            strTimestamp = strTimestamp.Trim();

            // Now reformat this 
            string YYYY = strTimestamp.Substring(6, 4);
            string MM = strTimestamp.Substring(3, 2);
            string dd = strTimestamp.Substring(0, 2);
            string time = strTimestamp.Substring(11, 8);
            strTimestamp = YYYY + "-" + MM + "-" + dd + " " + time;
            return strTimestamp;
        }



        public void writeHistoricData(string strExcelMasterFile, string strReferenceWorksheet, string strProjectDetailsWorksheet, string strFirstDataRow, string strFirstOutputRow, string strTimeBlockEnd, string[,] strPointDeltas)
        {
            //
            // Purpose:
            //      To read the data column from the project details worksheet
            //      To increment this value by 1 and to write it back to the worksheet
            //      To write to the Historic dH and Historic dS worksheets
            // Input:
            //      the strDeltas array must be available from strPointDeltas = gna.getPointDeltas()
            // Output:
            //      Nothing
            // Useage:
            // gna.writeHistoricData(strExcelMasterFile, strReferenceWorksheet, strProjectDetailsWorksheet, strFirstDataRow, strFirstOutputRow, strTimeBlockEnd, strPointDeltas);
            //

            // open the existing workbook
            var fiExcelMasterWorkbook = new FileInfo(strExcelMasterFile);
            using var package = new ExcelPackage(fiExcelMasterWorkbook);
            string strActiveWorksheet = strProjectDetailsWorksheet;
            var worksheet = package.Workbook.Worksheets[strActiveWorksheet];

            string strDataColumn = "blank";
            // read the active column in which the latest readings must be written
            strDataColumn = Convert.ToString(worksheet.Cells[1, 2].Value);
            int iDataColumn = Convert.ToInt16(worksheet.Cells[1, 2].Value) + 1;
            // Write this to the Project Details worksheet
            if (iDataColumn < 3)
            {
                iDataColumn = 3;
                strDataColumn = "3";
            }
            worksheet.Cells[1, 2].Value = iDataColumn;
            if (iDataColumn > 3) iDataColumn--;

            // Count the number of readings to be inserted
            string strName = "";
            int iCounter = 0;
            do
            {
                strName = strPointDeltas[iCounter, 0];
                iCounter++;
            } while (strName != "NoMore");

            int iNoOfReadings = iCounter - 2;

            // Make the reference worksheet active
            strActiveWorksheet = strReferenceWorksheet;
            worksheet = package.Workbook.Worksheets[strActiveWorksheet];
            // read in column T = current dS (col 20)
            // read in column U = current dH (col 21)
            double[] dblCurrentDs = new double[2500];
            double[] dblCurrentDh = new double[2500];

            iCounter = 0;
            int iRow = Convert.ToInt16(strFirstDataRow);
            do
            {
                dblCurrentDs[iCounter] = Convert.ToDouble(worksheet.Cells[iRow, 20].Value);
                dblCurrentDh[iCounter] = Convert.ToDouble(worksheet.Cells[iRow, 21].Value);
                iCounter++;
                iRow++;
            } while (iCounter < iNoOfReadings);

            // make the Historic dH worksheet active and write the dH values in the (incremented) column
            iRow = Convert.ToInt16(strFirstOutputRow);
            iCounter = 0;
            strActiveWorksheet = "Historic dH";
            worksheet = package.Workbook.Worksheets[strActiveWorksheet];

            // Write the reference date
            worksheet.Cells[6, iDataColumn].Value = strTimeBlockEnd;

            do
            {
                worksheet.Cells[iRow, iDataColumn].Value = dblCurrentDh[iCounter];
                iCounter++;
                iRow++;
            } while (iCounter < iNoOfReadings);


            // Repeat for the Historic dS worksheet
            iRow = Convert.ToInt16(strFirstOutputRow);
            iCounter = 0;
            strActiveWorksheet = "Historic dS";
            worksheet = package.Workbook.Worksheets[strActiveWorksheet];
            // Write the reference date
            worksheet.Cells[6, iDataColumn].Value = strTimeBlockEnd;
            strName = "";
            do
            {
                strName = strPointDeltas[iCounter, 0];
                worksheet.Cells[iRow, iDataColumn].Value = dblCurrentDs[iCounter];
                iCounter++;
                iRow++;
            } while (iCounter < iNoOfReadings);

            worksheet.Calculate();
            package.Save();

        }


        public void deleteMissingData(string strExcelWorkingFile, string strReferenceWorksheet, string strFirstOutputRow)
        {
            // walk through column M : Reading Count, looking for "Missing" and insert Missing in various columns
            // Then if applicable repeat the process but now interpolate top of rail values

            FileInfo excelWorkbook = new FileInfo(strExcelWorkingFile);

            using (ExcelPackage package = new ExcelPackage(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];
                // Read in the Reading Count, starting at the first data row.
                string strEndOfData = "Continue";

                int iRow = Convert.ToInt16(strFirstOutputRow);
                do
                {
                    string strText = Convert.ToString(namedWorksheet.Cells[iRow, 13].Value);
                    if (strText == "") strEndOfData = "Stop";

                    if (strText == "Missing")
                    {
                        Console.WriteLine("Missing: reset values");
                        namedWorksheet.Cells[iRow, 11].Value = "Missing";
                        namedWorksheet.Cells[iRow, 15].Value = "Missing";
                        namedWorksheet.Cells[iRow, 24].Value = "Missing";
                        namedWorksheet.Cells[iRow, 25].Value = "Missing";
                    }
                    iRow++;
                } while (strEndOfData == "Continue");

                namedWorksheet.Calculate();
                package.Save();
            }
        }

        

        public string[,] getPointID(string strDBconnection, string strProjectTitle, string[] strPointNames)
        {
            //
            // Purpose:
            //      To extract the point ID from TMCLocation table
            // Input:
            //      Receives array of point names extracted from the workbook
            // Output:
            //      Returns array [PointName,ID]
            // Useage:
            //      string[,] strNamesID = gna.getPointID(strDBconnection, strProjectTitle, strPointNames);
            // Comment:
            //      If missing then ID="Missing"
            //      last point in list = "NoMore"
            //

            string[,] strPointID = new string[2000, 2];
            int i = 0;
            string strPointName = "";
            int iObservationCounter = 0;
            int[] intLocationID = new int[100];
            string strLocationID = "";
            int iCounter = 0;

            // get the project ID and then use this in the selection
            string strProjectID = getProjectID(strDBconnection, strProjectTitle);

            // Connection and Reader declared outside the loop
            SqlConnection conn = new SqlConnection(strDBconnection);
            conn.Open();

            do
            {
                try
                {
                    strPointID[iCounter, 0] = "empty";

                    // define the SQL query
                    string SQLaction = @"
                    SELECT LocationID  
                    FROM dbo.TMCSensor 
                    WHERE TMCSensor.Name = @Name
                    AND TMCSensor.IsEnabled = 1
                    AND TMCSensor.IsDeleted = 0
                    ";

                    SqlCommand cmd = new SqlCommand(SQLaction, conn);

                    strPointName = strPointNames[iCounter];

                    // define the parameter used in the command object and add to the command
                    cmd.Parameters.Add(new SqlParameter("@Name", strPointName));

                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // Write the results (if there are any) to an array

                    iObservationCounter = 0;
                    if (dataReader != null)
                    {

                        while (dataReader.Read())
                        {
                            intLocationID[iObservationCounter] = Convert.ToInt16(dataReader["LocationID"]);
                            // Console.WriteLine(iObservationCounter + " LocationID: " + intLocationID[iObservationCounter]);
                            iObservationCounter++;
                        }

                        dataReader.Close();

                        // The array of Location IDs has been extracted.
                        // Now run through the array until the correct location is found

                        // define the SQL query

                        string SQLaction2 = @"
                        SELECT ID  
                        FROM dbo.TMCLocation 
                        WHERE TMCLocation.ProjectID = @ProjectID
                        AND TMCLocation.ID = @LocationID
                        ";

                        i = 0;
                        iObservationCounter--;
                        for (i = 0; i <= iObservationCounter; i++)
                        {
                            strLocationID = Convert.ToString(intLocationID[i]);
                            //Console.WriteLine("**"+strLocationID+"**");
                            // define the parameter used in the command object and add to the command
                            SqlCommand cmd2 = new SqlCommand(SQLaction2, conn);
                            cmd2.Parameters.Add(new SqlParameter("@ProjectID", strProjectID));
                            cmd2.Parameters.Add(new SqlParameter("@LocationID", strLocationID));

                            SqlDataReader dataReader2 = cmd2.ExecuteReader();

                            if (dataReader2.HasRows)
                            {
                                while (dataReader2.Read())
                                {
                                    strPointID[iCounter, 0] = strPointName;
                                    strPointID[iCounter, 1] = strLocationID;
                                }

                                //Console.WriteLine(iCounter+": "+strPointID[iCounter, 0] + " " + strPointID[iCounter, 1]);
                                //iCounter++;
                            }
                            else
                            {
                                strPointID[iCounter, 0] = strPointName;
                                strPointID[iCounter, 1] = "Missing";
                            }
                            dataReader2.Close();
                        }
                    }
                    else
                    {
                        strPointID[iCounter, 0] = strPointName;
                        strPointID[iCounter, 1] = "Missing";
                        iObservationCounter++;
                    }
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    Console.WriteLine("getPointID: ");
                    Console.WriteLine(strPointName);
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }
                finally
                {
                    if (strPointID[iCounter, 0] == "empty")
                    {
                        strPointID[iCounter, 0] = strPointName;
                        strPointID[iCounter, 1] = "Missing";
                    }
                }


                iCounter++;

            } while (strPointName != "NoMore");

            conn.Dispose();
            conn.Close();

            iCounter -= 1;
            strPointID[iCounter, 0] = "NoMore";
            strPointID[iCounter, 1] = "0";

            return strPointID;
        }

        public string[,] getReferenceCoordinates(string strDBconnection, string strProjectTitle, string[,] strNamesID)
        {

            //
            // Purpose:
            //      To extract the reference coordinates from TMCLocation table
            // Input:
            //      Receives array of point names & ID generated by getPointID, and the Project Title
            // Output:
            //      Returns array [PointName,N,E,Ht]   [0,1,2,3]
            // Useage:
            //      string[,] strReferenceCoordinates = gna.getReferenceCoordinates(strDBconnection, strProjectTitle, strNamesID);
            // Comment:
            //      If missing then coordinates are 999.999,999.999,999.999
            //      last point in list = "NoMore"
            //      Apply Displacements does not change these coordinates.
            //

            string[,] strReferenceCoordinates = new string[2000, 4];
            string strPointName = "";
            int[] intLocationID = new int[100];
            string strLocationID = "";

            // Set counters
            int iCounter = 0;

            // get the project ID
            string strProjectID = getProjectID(strDBconnection, strProjectTitle);

            // Connection and Reader declared outside the try block
            SqlConnection conn = new SqlConnection(strDBconnection);
            conn.Open();


            int i = 0;

            //Console.WriteLine("getReferenceCoordinates");

            do
            {

                strPointName = strNamesID[i, 0];
                strLocationID = strNamesID[i, 1];

                if (strLocationID == "Missing")
                {
                    strReferenceCoordinates[iCounter, 0] = strPointName;
                    strReferenceCoordinates[iCounter, 1] = "0";
                    strReferenceCoordinates[iCounter, 2] = "0";
                    strReferenceCoordinates[iCounter, 3] = "0";
                }
                else
                {
                    // define the SQL query
                    string SQLaction = @"
                    SELECT Northing, Easting, Elevation  
                    FROM dbo.TMCLocation 
                    WHERE TMCLocation.ID = @LocationID
                    ";

                    // define the parameter used in the command object and add to the command
                    SqlCommand cmd = new SqlCommand(SQLaction, conn);
                    cmd.Parameters.Add(new SqlParameter("@LocationID", strLocationID));

                    // Execute
                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // Assign
                    if (dataReader != null)
                    {

                        while (dataReader.Read())
                        {
                            strReferenceCoordinates[iCounter, 0] = strPointName;
                            strReferenceCoordinates[iCounter, 1] = Convert.ToString(dataReader["Northing"]);
                            strReferenceCoordinates[iCounter, 2] = Convert.ToString(dataReader["Easting"]);
                            strReferenceCoordinates[iCounter, 3] = Convert.ToString(dataReader["Elevation"]);
                            //Console.WriteLine(iCounter + "  " + strReferenceCoordinates[iCounter, 0] + "  " + strReferenceCoordinates[iCounter, 1] + "  " + strReferenceCoordinates[iCounter, 2] + "  " + strReferenceCoordinates[iCounter, 3]);
                        }

                    }

                    dataReader.Close();
                }

                iCounter++;
                i++;

            } while (strPointName != "NoMore");


            iCounter--;
            strReferenceCoordinates[iCounter, 0] = strPointName;
            strReferenceCoordinates[iCounter, 1] = "999";
            strReferenceCoordinates[iCounter, 2] = "999";
            strReferenceCoordinates[iCounter, 3] = "999";

            conn.Dispose();
            conn.Close();


            //Console.WriteLine("Extracted Coordinates");

            //for (i = 0; i <= iCounter; i++)
            //{
            //    Console.WriteLine(i + "  " + strReferenceCoordinates[i, 0] + "  " + strReferenceCoordinates[i, 1] + "  " + strReferenceCoordinates[i, 2] + "  " + strReferenceCoordinates[i, 3]);
            //}

            //Console.ReadKey();




            return strReferenceCoordinates;
        }


        public string[,] getSensorID(string strDBconnection, string strProjectTitle, string[,] strNamesID)
        {
            //
            // Purpose:
            //      To extract the sensor ID from TMCSensor table
            // Input:
            //      Receives array of strNamesID[point names, locationID] from getPointID(), and the project title
            // Output:
            //      Returns array [PointName,SensorID]
            // Useage:
            //      string[,] strSensorID = gna.getSensorID(strDBconnection, strProjectTitle, strNamesID);
            // Comment:
            //      If missing then ID="Missing"
            //      last point in list = "NoMore"
            //

            string[,] strSensorID = new string[2000, 2];
            string strPointName = "";
            int[] intLocationID = new int[100];
            string strLocationID = "";
            int iCounter = 0;

            // Connection and Reader declared outside the loop
            SqlConnection conn = new SqlConnection(strDBconnection);
            conn.Open();

            do
            {
                try
                {
                    // define the SQL query
                    string SQLaction = @"
                    SELECT ID  
                    FROM dbo.TMCSensor 
                    WHERE TMCSensor.Name = @Name
                    AND TMCSensor.LocationID = @LocationID
                    AND TMCSensor.IsEnabled = 1
                    AND TMCSensor.IsDeleted = 0
                    ";

                    SqlCommand cmd = new SqlCommand(SQLaction, conn);

                    strPointName = strNamesID[iCounter, 0];
                    strLocationID = strNamesID[iCounter, 1];

                    if (strLocationID == "Missing")
                    {
                        strSensorID[iCounter, 0] = strPointName;
                        strSensorID[iCounter, 1] = "Missing";
                        goto NextPoint;
                    }

                    // define the parameter used in the command object and add to the command
                    cmd.Parameters.Add(new SqlParameter("@Name", strPointName));
                    cmd.Parameters.Add(new SqlParameter("@LocationID", strLocationID));

                    SqlDataReader dataReader = cmd.ExecuteReader();

                    // Append the results to the output array

                    if (dataReader != null)
                    {
                        while (dataReader.Read())
                        {
                            strSensorID[iCounter, 1] = Convert.ToString(dataReader["ID"]);
                            strSensorID[iCounter, 0] = strPointName;
                        }

                        dataReader.Close();

                    }
                    else
                    {
                        strSensorID[iCounter, 0] = strPointName;
                        strSensorID[iCounter, 1] = "Missing";
                    }
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    Console.WriteLine("getSensorID: ");
                    Console.WriteLine(strPointName);
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }
NextPoint:
                iCounter++;

            } while (strPointName != "NoMore");

            conn.Dispose();
            conn.Close();

            iCounter -= 1;
            strSensorID[iCounter, 0] = "NoMore";
            strSensorID[iCounter, 1] = "0";

            return strSensorID;
        }

        public void writeSensorID(string strMasterWorkbookFullPath, string strReferenceWorksheet, string[,] strSensorID, string strFirstDataRow)
        {
            //
            // Purpose:
            //      To write the Sensor ID to the reference worksheet
            // Input:
            //      Full path and name of workbook
            //      Name of reference worksheet
            //      Array of Point Name and Sensor ID
            //      First data row
            // Output:
            //      --
            // Useage:
            //      gna.writeSensorID(strMasterWorkbookFullPath, strReferenceWorksheet, strSensorID, strFirstDataRow);
            //

            string strName = "";
            string strID = "";
            int iCounter = 0;

            // Write the reference coordinates to the workbook
            FileInfo newFile = new FileInfo(strMasterWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {

                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];

                //Set the column Header
                namedWorksheet.Cells[1, 1].Value = "SensorID";
                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);

                do
                {
                    strName = (strSensorID[iCounter, 0]).Trim();
                    strID = (strSensorID[iCounter, 1]).Trim();

                    if (strName == "NoMore") goto nextAction;
                    if (strName == "") goto nextAction;

                    //Console.WriteLine(strName + " " + strID);

                    if (strID != "Missing")
                    {
                        namedWorksheet.Cells[iRow, 1].Value = Convert.ToInt16(strSensorID[iCounter, 1]);
                    }
                    else
                    {
                        namedWorksheet.Cells[iRow, 1].Value = strID;
                    }

                    // apply basic format to this row
                    //int iRowBottom = iRow + 1;
                    using (var range = namedWorksheet.Cells[iRow, 1, iRow, 20])
                    {
                        range.Style.Font.Bold = false;
                        range.Style.Font.Color.SetAuto();
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        range.Style.Font.Size = 10;
                    }

                    iCounter++;
                    iRow++;
                } while (strName != "NoMore");
nextAction:
                namedWorksheet.Calculate();
                package.Save();
            }

            return;
        }

        public void formatFirstRowDisplacementReport(string strExcelWorkbookFullPath, string strWorksheet, string strCoordinateOrder)
        {
            checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
                using (var range = namedWorksheet.Cells["A1:ZZ1"])
                {
                    double rowHeight = 35;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.WrapText = true;
                    namedWorksheet.Row(1).Height = rowHeight;
                    namedWorksheet.Row(1).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                // Insert the Column headers

                if (strCoordinateOrder == "ENH")
                {
                    namedWorksheet.Cells[1, 1].Value = "SensorID";
                    namedWorksheet.Cells[1, 2].Value = "Target";
                    namedWorksheet.Cells[1, 3].Value = "E(base)";
                    namedWorksheet.Cells[1, 4].Value = "N(base)";
                    namedWorksheet.Cells[1, 5].Value = "H(base)";
                    namedWorksheet.Cells[1, 6].Value = "dE (mean)";
                    namedWorksheet.Cells[1, 7].Value = "dN (mean)";
                    namedWorksheet.Cells[1, 8].Value = "dH (mean)";
                    namedWorksheet.Cells[1, 9].Value = "E";
                    namedWorksheet.Cells[1, 10].Value = "N";
                    namedWorksheet.Cells[1, 11].Value = "H";
                    namedWorksheet.Cells[1, 12].Value = "Name";
                    namedWorksheet.Cells[1, 13].Value = "Reading Count";
                }
                else
                {
                    namedWorksheet.Cells[1, 1].Value = "SensorID";
                    namedWorksheet.Cells[1, 2].Value = "Target";
                    namedWorksheet.Cells[1, 3].Value = "N(base)";
                    namedWorksheet.Cells[1, 4].Value = "E(base)";
                    namedWorksheet.Cells[1, 5].Value = "H(base)";
                    namedWorksheet.Cells[1, 6].Value = "dN (mean)";
                    namedWorksheet.Cells[1, 7].Value = "dE (mean)";
                    namedWorksheet.Cells[1, 8].Value = "dH (mean)";
                    namedWorksheet.Cells[1, 9].Value = "N";
                    namedWorksheet.Cells[1, 10].Value = "E";
                    namedWorksheet.Cells[1, 11].Value = "H";
                    namedWorksheet.Cells[1, 12].Value = "Name";
                    namedWorksheet.Cells[1, 13].Value = "Reading Count";
                }

                namedWorksheet.Calculate();
                package.Save();
            }

            return;
        }

        public string[,] readPointNamesSensorID(string strExcelWorkbookFullPath, string strReferenceWorksheet, string strFirstDataRow)
        {
            //
            // Purpose:
            //      To read the sensor ID and point names from the reference worksheet
            // Input:
            //      Full path and name of workbook
            //      Name of reference worksheet
            //      First row of point names
            // Output:
            //      Returns array [sensorID,PointNames]
            // Useage:
            //      string[,] strSensorIDPointNames = gna.readPointNamesSensorID(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow);
            //



            // open the existing workbook
            var fiExcelSpreadsheet = new FileInfo(strExcelWorkbookFullPath);

            // navigate through the worksheets
            using (var package = new ExcelPackage(fiExcelSpreadsheet))
            {

                string strActiveWorksheet = strReferenceWorksheet;
                var worksheet = package.Workbook.Worksheets[strActiveWorksheet];

                int iRow = Convert.ToInt16(strFirstDataRow);
                string[,] strSensorIDPointNames = new string[2000, 2]; // 
                int i = 0;
                string strName;

                do
                {
                    // Read in the point names & sensorID
                    strSensorIDPointNames[i, 0] = Convert.ToString(worksheet.Cells[iRow, 2].Value);   // PointName
                    strSensorIDPointNames[i, 1] = Convert.ToString(worksheet.Cells[iRow, 1].Value);  // SensorID

                    strName = strSensorIDPointNames[i, 0];
                    iRow++;
                    i++;
                } while (strName != "");

                i--;
                strSensorIDPointNames[i, 0] = "NoMore";
                strSensorIDPointNames[i, 1] = "0";

                return strSensorIDPointNames;

            }

        }


        public void lockWorksheet(string strExcelWorkbookFullPath, string[] strWorksheets)
        {

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);
            

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                string strActiveWorksheet = "";
                int i = 0;
                do
                {
                    strActiveWorksheet = strWorksheets[i];

                    try
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[strActiveWorksheet];
                        worksheet.Cells[1, 1, 2000, 50].Style.Locked = true;
                        worksheet.Cells[1, 1, 2000, 50].Style.Hidden = true;
                        worksheet.Protection.SetPassword("GNAprotection2022");
                        worksheet.Protection.IsProtected = true;
                    }
                    catch
                    {
                        // the worksheet does not exist so do nothing
                        Console.WriteLine("worksheet missing");
                    }

                    i++;
                    strActiveWorksheet = strWorksheets[i];
                } while (strActiveWorksheet != "blank");

                
                package.Save();
            }

            return;
        }


        public void unlocklockWorksheet(string strExcelWorkbookFullPath, string[] strWorksheets)
        {

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);


            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                string strActiveWorksheet = "";
                int i = 0;
                do
                {
                    strActiveWorksheet = strWorksheets[i];
                    ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strActiveWorksheet];
                    namedWorksheet.Cells[1, 1, 2000, 50].Style.Locked = false;
                    namedWorksheet.Cells[1, 1, 2000, 50].Style.Hidden = false;
                    namedWorksheet.Protection.SetPassword("GNAprotection2022");
                    namedWorksheet.Protection.IsProtected = true;
                    i++;
                    strActiveWorksheet = strWorksheets[i];
                } while (strActiveWorksheet != "blank");

                package.Save();
            }

            return;
        }







        public void writeDeltas(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, string[,] strPointDeltas, int iColumn, string strTimeBlockStart, string strTimeBlockEnd, string strCoordinateOrder)
        {


            //
            // Purpose:
            //      To write the deltas to a defined worksheet
            // Deltas
            //      strPointDeltas[iCounter, 0] = strPointName;
            //      strPointDeltas[iCounter, 1] = MeandN
            //      strPointDeltas[iCounter, 2] = MeandE
            //      strPointDeltas[iCounter, 3] = MeandH
            //      strPointDeltas[iCounter, 4] = MeandR
            //      strPointDeltas[iCounter, 5] = MeandT
            //      strPointDeltas[iCounter, 6] = ObservationCounter = "-99" id there are no observations
            //      Full path and name of workbook
            //      Name of worksheet
            //      First data row
            //      Array of deltas
            //      First data column
            //      Time block start
            //      Time block end
            // Output:
            //      --
            // Useage:
            //      gna.writeDeltas(strExcelWorkbookFullPath, strWorksheet, strFirstDataRow, strPointDeltas, iColumn, strTimeBlockStart, strTimeBlockEnd, strCoordinateOrder);
            // Comment
            //      This can write the deltas to any column - to make provision for a rolling time window of multiple time blocks
            //      The column number gets generated at the same time as the time block (generateTimeBlockStartEnd)
            //

            string strName = "";
            int iCounter = 0;

            double dbldN;
            double dbldE;
            double dbldH;
            int iNumberOfObs;

            // Write the Deltas to the workbook
            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);
            ExcelPackage package = new ExcelPackage(newFile);
            //using (ExcelPackage package = new ExcelPackage(newFile))

            using (package)
            {

                // add a new worksheet to the workbook if it does not exist
                try
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(strWorksheet);
                }
                catch
                {
                    // it exists so do nothing
                }
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];

                iCounter = 0;
                int iRow = Convert.ToInt16(strFirstDataRow);

                do
                {

                    string strStatus = strPointDeltas[iCounter, 6];   // Reading Count

                    strName = strPointDeltas[iCounter, 0];
                    dbldN = Convert.ToDouble(strPointDeltas[iCounter, 1]);
                    dbldE = Convert.ToDouble(strPointDeltas[iCounter, 2]);
                    dbldH = Convert.ToDouble(strPointDeltas[iCounter, 3]);
                    iNumberOfObs = Convert.ToInt16(strPointDeltas[iCounter, 6]);

                    //Console.WriteLine(strPointDeltas[iCounter, 0] + "  " + strPointDeltas[iCounter, 1] + "  " + strPointDeltas[iCounter, 2] + "  " + strPointDeltas[iCounter, 3] + "  " + strPointDeltas[iCounter, 6]);
                    //Console.ReadKey();


                    if (strName != "NoMore")
                    {
                        if (strStatus == "-99")
                        {
                            namedWorksheet.Cells[iRow, 2].Value = strName;
                            namedWorksheet.Cells[iRow, iColumn].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 1].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 2].Value = 0.00;
                            namedWorksheet.Cells[iRow, iColumn + 9].Value = "Missing";
                            namedWorksheet.Cells[iRow, iColumn].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 1].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 2].Style.Font.Italic = true;
                            namedWorksheet.Cells[iRow, iColumn + 9].Style.Font.Italic = true;
                        }
                        else
                        {
                            namedWorksheet.Cells[iRow, iColumn].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 1].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 2].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, iColumn + 9].Style.Font.Italic = false;
                            namedWorksheet.Cells[iRow, 2].Value = strName; //Name

                            if (strCoordinateOrder == "ENH")
                            {
                                namedWorksheet.Cells[iRow, iColumn].Value = dbldE;
                                namedWorksheet.Cells[iRow, iColumn + 1].Value = dbldN;
                                namedWorksheet.Cells[iRow, iColumn + 2].Value = dbldH;
                                namedWorksheet.Cells[iRow, iColumn + 9].Value = iNumberOfObs;
                            }
                            else
                            {
                                namedWorksheet.Cells[iRow, iColumn].Value = dbldN;
                                namedWorksheet.Cells[iRow, iColumn + 1].Value = dbldE;
                                namedWorksheet.Cells[iRow, iColumn + 2].Value = dbldH;
                                namedWorksheet.Cells[iRow, iColumn + 9].Value = iNumberOfObs;
                            }

                        }
                        iCounter++;
                        iRow++;
                    }

                } while (strName != "NoMore");

                namedWorksheet.Cells[iRow + 1, iColumn + 1].Value = strTimeBlockStart;
                namedWorksheet.Cells[iRow + 2, iColumn + 1].Value = strTimeBlockEnd;


                if (strWorksheet == "Reference")
                {
                    drawBox(strExcelWorkbookFullPath, strWorksheet, 9, 11, 2, iRow - 1);
                    drawBox(strExcelWorkbookFullPath, strWorksheet, 18, 18, 2, iRow - 1);
                }

                // force computation of worksheet before saving
                namedWorksheet.Calculate();
                package.Save();
            }

            
            return;
        }

        public void formatFirstRowTrackGeometryReport(string strExcelWorkbookFullPath, string strWorksheet, string strCoordinateOrder)
        {
            checkWorksheetExists(strExcelWorkbookFullPath, strWorksheet);

            FileInfo newFile = new FileInfo(strExcelWorkbookFullPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
                using (var range = namedWorksheet.Cells["A1:ZZ1"])
                {
                    double rowHeight = 35;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                    //range.Style.Font.Color.SetAuto();
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.WrapText = true;
                    namedWorksheet.Row(1).Height = rowHeight;
                    namedWorksheet.Row(1).Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    namedWorksheet.Row(1).Style.Border.Bottom.Color.SetColor(Color.Yellow);   
                }

                if (strCoordinateOrder == "ENH") 
                {
                // Insert the Column headers
                namedWorksheet.Cells[1, 1].Value = "SensorID";
                namedWorksheet.Cells[1, 2].Value = "Prism";
                namedWorksheet.Cells[1, 3].Value = "E(CrdList)";
                namedWorksheet.Cells[1, 4].Value = "N(CrdList)";
                namedWorksheet.Cells[1, 5].Value = "H(CrdList)";
                namedWorksheet.Cells[1, 6].Value = "dE (current)";
                namedWorksheet.Cells[1, 7].Value = "dN (current)";
                namedWorksheet.Cells[1, 8].Value = "dH (current)";
                namedWorksheet.Cells[1, 9].Value = "dE (constant)";
                namedWorksheet.Cells[1, 10].Value = "dN (constant)";
                namedWorksheet.Cells[1, 11].Value = "dH (constant)";
                namedWorksheet.Cells[1, 12].Value = "E (current)";
                namedWorksheet.Cells[1, 13].Value = "N (current)";
                namedWorksheet.Cells[1, 14].Value = "H (current)";
                namedWorksheet.Cells[1, 15].Value = "Reading Count";
                namedWorksheet.Cells[1, 16].Value = "New Name";
                namedWorksheet.Cells[1, 17].Value = "Prism Offset";
                namedWorksheet.Cells[1, 18].Value = "Top of Rail";
                namedWorksheet.Cells[1, 19].Value = " ";
                namedWorksheet.Cells[1, 20].Value = "E (ref)";
                namedWorksheet.Cells[1, 21].Value = "N (ref)";
                namedWorksheet.Cells[1, 22].Value = "H (ref)";
                namedWorksheet.Cells[1, 23].Value = "dS (ref)";
                namedWorksheet.Cells[1, 24].Value = "dH (ref)";
                namedWorksheet.Cells[1, 25].Value = "dS (corr)";
                namedWorksheet.Cells[1, 26].Value = "dH (corr)";
                namedWorksheet.Cells[1, 27].Value = "dS (current)";
                namedWorksheet.Cells[1, 28].Value = "dH (current)";
                }
                else
                {
                    // Insert the Column headers
                    namedWorksheet.Cells[1, 1].Value = "SensorID";
                    namedWorksheet.Cells[1, 2].Value = "Prism";
                    namedWorksheet.Cells[1, 3].Value = "N(CrdList)";
                    namedWorksheet.Cells[1, 4].Value = "E(CrdList)";
                    namedWorksheet.Cells[1, 5].Value = "H(CrdList)";
                    namedWorksheet.Cells[1, 6].Value = "dN (current)";
                    namedWorksheet.Cells[1, 7].Value = "dE (current)";
                    namedWorksheet.Cells[1, 8].Value = "dH (current)";
                    namedWorksheet.Cells[1, 9].Value = "dN (constant)";
                    namedWorksheet.Cells[1, 10].Value = "dE (constant)";
                    namedWorksheet.Cells[1, 11].Value = "dH (constant)";
                    namedWorksheet.Cells[1, 12].Value = "N (current)";
                    namedWorksheet.Cells[1, 13].Value = "E (current)";
                    namedWorksheet.Cells[1, 14].Value = "H (current)";
                    namedWorksheet.Cells[1, 15].Value = "Reading Count";
                    namedWorksheet.Cells[1, 16].Value = "New Name";
                    namedWorksheet.Cells[1, 17].Value = "Prism Offset";
                    namedWorksheet.Cells[1, 18].Value = "Top of Rail";
                    namedWorksheet.Cells[1, 19].Value = " ";
                    namedWorksheet.Cells[1, 20].Value = "N (ref)";
                    namedWorksheet.Cells[1, 21].Value = "E (ref)";
                    namedWorksheet.Cells[1, 22].Value = "H (ref)";
                    namedWorksheet.Cells[1, 23].Value = "dS (ref)";
                    namedWorksheet.Cells[1, 24].Value = "dH (ref)";
                    namedWorksheet.Cells[1, 25].Value = "dS (corr)";
                    namedWorksheet.Cells[1, 26].Value = "dH (corr)";
                    namedWorksheet.Cells[1, 27].Value = "dS (current)";
                    namedWorksheet.Cells[1, 28].Value = "dH (current)";
                }

                namedWorksheet.Calculate();
                package.Save();
            }

            return;
        }



        public void TGR_interpolateMissingData(string strExcelWorkingFile, string strReferenceWorksheet, string strFirstOutputRow)
        {
            // To interpolate missing Top of Rail Elevations on the track geometry report
            // Useage:
            // gna.TGR_interpolateMissingData(strExcelWorkingFile, strReferenceWorksheet, strFirstOutputRow);
            // 
            //

            int iStart = 0;
            int iEnd = 0;
            double dblToRstart = 0.0;
            double dblToRend = 0.0;
            double dblDh = 0.0;

            FileInfo excelWorkbook = new FileInfo(strExcelWorkingFile);

            using (ExcelPackage package = new ExcelPackage(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];
                namedWorksheet.Calculate();
                package.Save();

                string strText = "Continue";
                int iRow = Convert.ToInt16(strFirstOutputRow);
                int j = iRow;

                //
                // Reading count is column 15
                // if the data is missing the cell contains "Missing"
                // "Top of Rail" data will be interpolated (col 18)
                // If the first or last prism is missing, data is not interpolated - the interpolated data must be bracketed
                //

                // Locate the first prism that has observed top of Rail data

                do
                {

                    strText = Convert.ToString(namedWorksheet.Cells[iRow, 15].Value);

                    do
                    {
                        iRow++;
                        strText = Convert.ToString(namedWorksheet.Cells[iRow, 15].Value);  // Reading Count: status or count
                        if (strText == "") goto DataCompleted;
                        if (strText == "Missing") goto InterpolateStart;
                    } while (strText != "Missing");
InterpolateStart:
                    iStart = iRow - 1;

                    // iStart: The row number before the block of missing data begins
                    // dblToRstart: The last observed Top of Rail value;

                    // Now step through the missing values until the first observed Top of Rail reading
                    string strRailBracket;
                    do
                    {
                        strText = Convert.ToString(namedWorksheet.Cells[iRow, 15].Value);
                        strRailBracket= Convert.ToString(namedWorksheet.Cells[iRow, 30].Value);  // Rail Start, Rail, Rail End

                        if (strText == "") {
                            int iPreviousRow = iRow - 1;
                            if ((Convert.ToString(namedWorksheet.Cells[iPreviousRow, 15].Value)) == "Missing")
                            {
                                goto CarryOn;
                            }
                            goto DataCompleted;   // opt out at end of data
                        }   
                        iRow++;
                    } while ((strText == "Missing") && (strRailBracket == "Rail"));
CarryOn:
                    iRow--;
                    iEnd = iRow;

                    //Missing data lies between :  iStart +  iEnd
                    //Top of Rail Heights: from dblToRstart to dblToRend

                    string strRailElement = Convert.ToString(namedWorksheet.Cells[iStart, 30].Value);    //Rail Start, Rail, Rail End

                    // set start value iStart and dblToRstart

                    //Console.WriteLine("before correction");
                    //Console.WriteLine("iStart: " + iStart + " " + Convert.ToString(namedWorksheet.Cells[iStart, 30].Value));
                    //Console.WriteLine("iend: " + iEnd + " " + Convert.ToString(namedWorksheet.Cells[iEnd, 30].Value));
                    //Console.WriteLine("");



                    if (strRailElement == "Rail End")
                    {
                        goto NextLine;
                    }
                    else if (strRailElement == "Rail Start")
                    {
                        dblToRstart = Convert.ToDouble(namedWorksheet.Cells[iStart, 31].Value);
                    }
                    else
                    {
                        dblToRstart = Convert.ToDouble(namedWorksheet.Cells[iStart, 18].Value);
                    }

                    strRailElement = Convert.ToString(namedWorksheet.Cells[iEnd, 30].Value);    //Rail Start, Rail, Rail End

                    if (strRailElement == "Rail Start")
                    {
                        iEnd--;
                        dblToRend = Convert.ToDouble(namedWorksheet.Cells[iEnd, 31].Value);
                    }
                    else if (strRailElement == "Rail Start")
                    {
                        dblToRend = Convert.ToDouble(namedWorksheet.Cells[iEnd, 31].Value);
                    }
                    else
                    {
                        dblToRend = Convert.ToDouble(namedWorksheet.Cells[iEnd, 18].Value);
                    }

                    // Populate the Reference worksheet with interpolated data

                    int iSegments = iEnd - iStart;
                    dblDh = Math.Round(((dblToRend - dblToRstart) / iSegments),4);
                    dblToRstart = Math.Round(dblToRstart, 4);
                    dblToRend = Math.Round(dblToRend, 4);


                    //Console.WriteLine("After correction");
                    //Console.WriteLine("iStart: " + iStart + " " + dblToRstart);
                    //Console.WriteLine("iEnd: " + iEnd + " " + dblToRend);
                    //Console.WriteLine("Segments: " + iSegments);
                    //Console.WriteLine("dH: " + dblDh);
                    //Console.ReadLine();

                    int i = 1;
                    double dblInterpolatedToR = Math.Round(dblToRstart,4);

                    //Console.WriteLine("start " + dblInterpolatedToR);

                    iRow = iStart + 1;
                    do
                    {
                        dblInterpolatedToR = Math.Round((dblInterpolatedToR + dblDh),4);

                        //Console.WriteLine(dblInterpolatedToR);

                        namedWorksheet.Cells[iRow, 18].Value = dblInterpolatedToR;
                        namedWorksheet.Cells[iRow, 18].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        namedWorksheet.Cells[iRow, 14].Style.Fill.BackgroundColor.SetColor(Color.Coral);
                        namedWorksheet.Cells[iRow, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        namedWorksheet.Cells[iRow, 18].Style.Fill.BackgroundColor.SetColor(Color.Coral);
                        namedWorksheet.Cells[iRow, 24].Style.Fill.BackgroundColor.SetColor(Color.Coral);
                        namedWorksheet.Cells[iRow, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        namedWorksheet.Cells[iRow, 28].Style.Fill.BackgroundColor.SetColor(Color.Coral);

                        namedWorksheet.Cells[iRow, 28].Value = "Interpolated";

                        string strFormula = "=(R" + Convert.ToString(iRow) + ")-(Q" + Convert.ToString(iRow) + ")";
                        namedWorksheet.Cells[iRow, 14].Formula = strFormula;

                        namedWorksheet.Cells[iRow, 14].Style.Numberformat.Format = "0.000";
                        namedWorksheet.Cells[iRow, 14].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 15].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 16].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 18].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 24].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 28].Style.Font.Italic = true;
                        namedWorksheet.Cells[iRow, 29].Style.Font.Italic = true;

                        i++;
                        iRow++;
                    } while (i < iSegments);

NextLine:
                    strText = Convert.ToString(namedWorksheet.Cells[iRow, 15].Value);

                } while (strText != "");  

DataCompleted:
                namedWorksheet.Calculate();
                package.Save();

            }
        }













    }
        
}