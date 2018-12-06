using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;
using System.Timers;

using System.Data;
using System.Data.SqlClient;

using System.Configuration;

//using OfficeOpenXml;

using System.Collections.Generic;

using FileHelpers;
//using FileHelpers.RunTime;

using CyberSource.Clients;
//using CyberSource.Clients.SoapWebReference;
using CyberSource.Clients.SoapServiceReference;

namespace ARC_CallProcessing_CyberSource
{
    /// <summary>
    /// Processing Call Center (LiveOps) Files
    /// </summary>
    public partial class Service1 : ServiceBase
    {
        int row = 0;
        int col = 0;
        /// <summary>
        /// Step 1:	Open .imp file
        /// Step 2:	Process Records in .imp file
        /// Step 3:	Create .exp file save and close
        /// Step 4:	Close .imp file
        /// Step 5:	Rename .imp file to .dne file
        /// Step 6:	Loop for next .imp file
        /// </summary>
        #region Log/Run Cycle Settings
        int EmailCycle = 20; // Hours
        int RunCycle = 3; // Minimum Hours to sleep before running again
        bool logToggle = false;
        bool runToggle = false; // This will ensure we run if the  RunCycle is not met yet (we started the service within les than that)
        bool runLoop = false;
        bool running = false;
        bool logSleep = false;
        bool oDebug = Convert.ToBoolean(ConfigurationSettings.AppSettings["modeDebug"].ToString()); // Hard Coded Run Once
        bool runOnStart = Convert.ToBoolean(ConfigurationSettings.AppSettings["runOnStart"].ToString()); // Hard Coded Run Once
        DateTime dtStart = DateTime.UtcNow;
        DateTime dtLoop = DateTime.UtcNow;
        DateTime dtEmail = DateTime.UtcNow;
        DateTime dtLog = DateTime.UtcNow;

        #endregion Log/Run Cycle Settings
        #region Hard Coded and Soft Coded initial values
        private System.Timers.Timer timer = null;
        string serviceName = ConfigurationSettings.AppSettings["serviceName"];
        string logging = ConfigurationSettings.AppSettings["logging"].ToUpper(); //Login Type

        string servicepath = ConfigurationSettings.AppSettings["servicePath"]; //If File - Path
        string logfilepath = ConfigurationSettings.AppSettings["servicePath"] + ConfigurationSettings.AppSettings["logfilePath"]; //If File - Path
        string logfilename = ConfigurationSettings.AppSettings["logfileName"]; //If File - Name
        long logfilemaxsize; //Will clear the file once maxsize is reached
        int logInterval = 60;
        int cntrLoop = 0;
        int cntrFiles = 0;
        //int dailyLoop = 0;

        //String flResponseHTML = "";

        int epochID = (int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds;

        String lnBreak = "--------------------------------------------------------------------";

        #endregion Hard Coded and Soft Coded initial values
        #region Seald Classe(s)
        /// <summary>
        /// This class is used to populate the SQL Stored Procedure:
        /// lms_send_leads_task_attempt_vestigi
        /// We use a class to verify that we are sending the proper formats to SQL
        /// </summary>
        /// <summary>
        /// This class is used to populate the SQL Stored Procedure:
        /// lms_send_leads_task_attempt_vestigi
        /// We use a class to verify that we are sending the proper formats to SQL
        /// </summary>
        Process_Count pCounts = new Process_Count();
        public sealed class Process_Count
        {
            public int Loops;
            public int Errors;
            public int Records;
            public int RecordsErr;
            public int getError;
            public DateTime Start = DateTime.UtcNow;
        }
        Error_Message errMsg = new Error_Message();
        public sealed class Error_Message
        {
            public int Count;
            public String Error;
            public String Message;
            public String Source;
            public String StackTrace;
            public DateTime TimeStamp;
        }
        #endregion
        public Service1()
        {
            InitializeComponent();
            string servicepollinterval = ConfigurationSettings.AppSettings["servicepollinterval"];
            double interval = 30000;
            try
            {
                interval = Convert.ToDouble(servicepollinterval);
            }
            catch (Exception) { }

            string logfileinterval = ConfigurationSettings.AppSettings["logfileinterval"];
            try
            {
                logInterval = Convert.ToInt32(logfileinterval);
            }
            catch (Exception) { }

            timer = new System.Timers.Timer(interval);
            timer.Elapsed += new ElapsedEventHandler(this.ServiceTimer_Tick);

            string logfilemaxsizestr = ConfigurationSettings.AppSettings["logfilemaxsize"];
            logfilemaxsize = 10000;
            try
            {
                logfilemaxsize = Convert.ToInt64(logfilemaxsizestr);
            }
            catch (Exception) { }
        }
        /// <summary> 
        /// Start: Service started - Email Sent
        /// Stop: Service stoped - Email Sent - research "Safe" stop
        /// Pause: Service stoped - research "Safe" stop
        /// Continue: Service stoped - research "Safe" stop
        /// </summary> 
        protected override void OnStart(string[] args)
        {
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Start();
            String msg = String.Format("{0} [{1}]", "Service Started", dtStart);
            msg += "|" + Connection.GetConnectionType();
            Log(msg, "standard");
            //SendStandardEmail("Service Was Started", "Started", "none");
        }
        protected override void OnStop()
        {
            timer.AutoReset = false;
            timer.Enabled = false;
            //SendStandardEmail("Service Was Stopped", "Stopped", "none");
            String msg = String.Format("{0} [{1}]", "Service Stopped", DateTime.UtcNow);
            Log(msg, "standard");
        }
        protected override void OnPause()
        {
            this.timer.Stop();
            //SendStandardEmail("Service Was Paused", "Paused", "none");
            Log("Service Paused", "standard");
        }
        protected override void OnContinue()
        {
            this.timer.Start();
            Log("Service Continued", "standard");
            //SendStandardEmail("Service Was Continued", "Continued", "none");
        }
        /// <summary>
        /// Main loop, we need to ensure this is not running...?
        /// </summary>
        private void ServiceTimer_Tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            this.timer.Stop();
            Loop_Start();
            //SendStandardEmail("Service Loop", "Loop", "none");
            //Log("Service Loop", "standard");
            this.timer.Start();
        }
        /// <summary>
        /// Verify this is not already running
        /// Add a counter
        /// Send an email once a [period] (?day?)
        /// 
        /// Processing a DNC File against Visions
        /// 1. Check for File(s)
        /// 2. Update Lead in DE
        /// 3. Send Update to Visions
        /// 4. Next Record
        /// 5. Send Count/Processing Email
        /// </summary>
        public void Loop_Start()
        {
            if (!running)
            {
                // Run @ 1 AM Daily
                // Sleep for 22 Hours
                // Log Daily
                // If DateTime >= 1 AM && LastRun > 5 Hours ? Run : Do Not Run
                if (pCounts.Start < dtEmail) { pCounts.Start = dtEmail; }
                pCounts.Loops++;
                cntrLoop++; // Loops
                // Run When We Start if oDebug
                if (oDebug)
                {
                    // Run if we are in deBug mode when we start
                    if (dtStart == dtLoop && runLoop == false) { runLoop = true; Log_Running("Loop-001"); }
                }
                else if ((DateTime.UtcNow.ToString("MM/dd/yyyy HH") == DateTime.Parse("5/17/2012 11:00").ToString("MM/dd/yyyy HH")))
                {
                    // Run for a hard coded date (remove?)
                    Log_Running("Loop-009");
                    runLoop = true;
                }
                else if (runOnStart)
                {
                    // Run on Start if desired
                    Log_Running("Loop-010");
                    runLoop = true;
                    runOnStart = false;
                }
                else
                {
                    // Run if it is 1 AM and we have slept for 22 hours
                    // Or we have not run yet
                    if ((DateTime.UtcNow - dtLoop).TotalHours >= RunCycle || !runToggle)
                    {
                        string runHour = ConfigurationSettings.AppSettings["servicesingleruntime"];
                        //(DateTime.UtcNow.Hour == DateTime.Parse("01:00:00").Hour).ToString();
                        if (DateTime.UtcNow.Hour == DateTime.Parse("05:00:00").Hour)
                        {
                            Log_Running("Loop-005:00");
                            runLoop = true;
                            runToggle = true;
                        }
                        else if (DateTime.UtcNow.Hour == DateTime.Parse("10:00:00").Hour)
                        {
                            Log_Running("Loop-010:00");
                            runLoop = true;
                            runToggle = true;
                        }
                        else
                        {
                            // If we got here via runToggle, do not sleep log it?
                            // if (runToggle) { Log_Running("Sleep-000"); }
                        }
                        // If it's X Hour
                        // And we have not run Today
                        // Run
                        if (DateTime.UtcNow.Hour == DateTime.Parse(runHour).Hour
                            && dtLoop.ToString("MM/dd/yyyy") != DateTime.UtcNow.ToString("MM/dd/yyyy")
                            )
                        {
                            Log_Running("Loop-002");
                            runLoop = true;
                            runToggle = true;
                        }
                    }
                }
                //runLoop = true; // Hard Code Run
                if (runLoop)
                {
                    Processing_Start();
                    dtLoop = DateTime.UtcNow;
                    if (pCounts.Records > 0)
                    {
                        //Notification_Check("File Created", "EmailGroup1");
                    }
                    else
                    {
                        //Notification_Check("No Records to Process", "EmailGroup1");
                    }
                    Log_Running("Loop-003-End of Loop");
                    runLoop = false;
                }
                else
                {
                    if (logSleep) { Log_Running("Sleep-001"); }
                }
            }
            else
            {
                //This is already running....?
                Log_Running("Error");
                cntrLoop++;
            }

            string servicepollsleep = ConfigurationSettings.AppSettings["servicepollsleep"];
            int sleep = 30000; //1000 = 1 second (30,000)
            if (pCounts.Records >= 0)
            {
                try
                {
                    sleep = Convert.ToInt32(servicepollsleep);
                }
                catch (Exception) { }
            }
            // sleep = 30000; // Sleep for X seconds that remain in the hour + 5 mins.
            // Next cycle
            try
            {
                sleep = Convert.ToInt32(((DateTime.Parse(DateTime.UtcNow.AddHours(1).ToString("MM/dd/yyyy HH:25:00")) - DateTime.UtcNow).TotalMilliseconds));
            }
            catch { }
            //if (oDebug) { Log_Running("Sleep for " + (sleep).ToString()); }
            //else { Log("Sleep at " + DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss") + " for " + (sleep).ToString(), "sleep"); }
            Log("Sleep at " + DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss") + " for " + SecondsTo(sleep), "sleep");
            System.Threading.Thread.Sleep(sleep); //DeBug
        }
        protected String SecondsTo(Double Seconds)
        {
            Seconds = Seconds / 1000;
            TimeSpan time = TimeSpan.FromSeconds(Seconds);

            String rtrn = String.Format("{0}:{1}:{2}",
                Math.Floor(time.TotalHours).ToString().PadLeft(2, '0'),
                time.Minutes.ToString().PadLeft(2, '0'),
                time.Seconds.ToString().PadLeft(2, '0'));

            return rtrn;
        }

        /// <summary>
        /// Describe the process...
        /// </summary>
        protected void Processing_Start()
        {
            #region Processing Dir of Files - Try
            try
            {
                // Check for files
                // Insert the Call Record
                // Insert other table records
                // Determine if we need to process credit card
                // Finalize Record
                // Close File
                // Move File
                // Back to 1

                // Live
                #region Check for files
                string inputPath = ConfigurationManager.AppSettings["INPUT_PATH"];
                string outputPath = ConfigurationManager.AppSettings["OUTPUT_PATH"];

                DirectoryInfo dir = new DirectoryInfo(inputPath);
                string fileSearchString = ConfigurationManager.AppSettings["fileSearchString"]; //"LO_NBC_SANDY*.txt" LO_OK130525051945.txt
                FileInfo[] files = dir.GetFiles(fileSearchString, SearchOption.TopDirectoryOnly);
                
                string fName = String.Empty;
                #endregion Check for files
                DateTime startTime = DateTime.UtcNow;
                if (files.Length > 0)
                {
                    foreach (FileInfo file in files)
                    {
                        #region FileInfo file in files
                        FileHelperEngine engine = new FileHelperEngine(typeof(ARC_CallImport));
                        ARC_CallImport[] batchRecords = (ARC_CallImport[])engine.ReadFile(file.FullName);
                        int tRecords = engine.TotalRecords;
                        //fName = Path.GetFileNameWithoutExtension(file.FullName);
                        fName = Path.GetFileName(file.FullName);
                        Log("Processing File: " + fName + " [" + tRecords.ToString() + "]", "standard");
                        if (tRecords > 0)
                        {
                        }
                        else
                        {
                        }
                        #region SqlConnection
                        string sqlCon = "ARC_Production";
                        if (oDebug)
                        {
                            sqlCon = "ARC_Stage";
                        }
                        using (SqlConnection con = new SqlConnection(Connection.GetConnectionString(sqlCon, "")))
                        {
                            Donation_Open_Database(con);
                            #region Process Each Record in the File
                            if (engine.TotalRecords > 0)
                            {
                                
                                foreach (ARC_CallImport rcrd in batchRecords)
                                {
                                    pCounts.Records++;
                                    string logRecord = "";
                                    string rcrdStatus = string.Empty;
                                    try
                                    {
                                        #region Process Record
                                        // Start the log record
                                        logRecord = rcrd.confirmation;
                                        logRecord += "," + rcrd.callcenter;
                                        logRecord += "," + rcrd.campaign;
                                        logRecord += "," + rcrd.callstart;
                                        //Log(logRecord, "record");

                                        int callid;
                                        int callcreateid;
                                        int callinfoid;
                                        int chargedateid;
                                        int donationccinfoid;
                                        int cybersourceid;
                                        bool insertcc = false;
                                        bool processpayment = false;
                                        DateTime sp_callstart = DateTime.UtcNow;
                                        // DateTime.UtcNow
                                        int sp_timezone = 0;
                                        if (DateTime.TryParse(rcrd.callstart, out sp_callstart))
                                        {
                                            Int32.TryParse(rcrd.timezone, out sp_timezone);
                                            if (sp_timezone != 0)
                                            {
                                                sp_callstart = sp_callstart.AddHours(sp_timezone);
                                            }
                                        }
                                        DateTime sp_enddate = DateTime.UtcNow;
                                        if (DateTime.TryParse(rcrd.callend, out sp_enddate))
                                        {
                                            Int32.TryParse(rcrd.timezone, out sp_timezone);
                                            if (sp_timezone != 0)
                                            {
                                                sp_enddate = sp_enddate.AddHours(sp_timezone);
                                            }
                                        }

                                        #region Insert: CALL
                                        string CallUUID = System.Guid.NewGuid().ToString("B").ToUpper(); // Create CallUUID
                                        using (SqlCommand cmd = new SqlCommand("", con))
                                        {
                                            #region Populate the SQL Command
                                            cmd.CommandTimeout = 600;
                                            cmd.CommandText = "[dbo].[callprocessing_insert_call]";
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            #endregion Populate the SQL Command
                                            #region Populate the SQL Params
                                            cmd.Parameters.Add(new SqlParameter("@sp_calluuid", CallUUID));
                                            cmd.Parameters.Add(new SqlParameter("@sp_personid", rcrd.agent_id));
                                            cmd.Parameters.Add(new SqlParameter("@sp_logindatetime", sp_callstart));
                                            cmd.Parameters.Add(new SqlParameter("@sp_dnis", rcrd.dnis));
                                            cmd.Parameters.Add(new SqlParameter("@sp_callenddatetime", sp_enddate));
                                            cmd.Parameters.Add(new SqlParameter("@sp_languageid", "0"));
                                            cmd.Parameters.Add(new SqlParameter("@sp_dispositionid", "41")); // Donation
                                            cmd.Parameters.Add(new SqlParameter("@sp_originationid", rcrd.confirmation));
                                            #endregion Populate the SQL Params
                                            #region Process SQL Command - Try
                                            try
                                            {
                                                callid = Convert.ToInt32(cmd.ExecuteScalar());
                                                LogSQL(cmd, "sqlPassed");
                                            }
                                            #endregion Process SQL Command - Try
                                            #region Process SQL Command - Catch
                                            catch (Exception ex)
                                            {
                                                callid = -1;
                                                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                                LogSQL(cmd, "sqlFailed");
                                            }
                                            #endregion Process SQL Command - Catch
                                        }
                                        if (callid == 0 || callid == -1 || callid == -2)
                                        {
                                            // 0 == error in SQL
                                            // -1 == error in .NET
                                            // -2 == record exists
                                            switch (callid)
                                            {
                                                case 0:
                                                    rcrdStatus = "call.sql";
                                                    break;
                                                case -1:
                                                    rcrdStatus = "call.net";
                                                    break;
                                                case -2:
                                                    rcrdStatus = "call.exists";
                                                    break;
                                            }
                                            throw new Exception("There was a problem inserting the [call] record {Error: " + callid.ToString() + "}.");
                                        }
                                        #endregion Insert: CALL

                                        #region Insert: AGENTLOG - DATAEXCHANGE
                                        /// agent_firstname
                                        /// If we have a [agent_firstname] - we insert the details
                                        /// If we do not, we do not insert anything
                                        /// This also goes into DE database
                                        #region Pre-Flight
                                        bool chckExists = false;
                                        int sp_interactionid = 0;
                                        int sp_duration = 0;
                                        int sp_companyid = 3; // ARC
                                        int sp_interactiontype = 1001; // Call
                                        int sp_resourcetype = 10801; // LiveOps
                                        int sp_resourceid = 10801; // LiveOps
                                        int sp_status = 1; // Default
                                        int sp_arc_disposition_id = -1;
                                        string sp_arc_disposition_name = "";
                                        Int32.TryParse(rcrd.callduration, out sp_duration);

                                        int sp_agent_id = 0;
                                        Int32.TryParse(rcrd.agent_id, out sp_agent_id);
                                        string sp_agent_full_name = String.Format("{0} {1}", rcrd.agent_firstname, rcrd.agent_lastname).Trim();
                                        int sp_agent_station_id = 0;
                                        int sp_agent_station_type = 0;
                                        string sp_agent_user_name = "";
                                        String sqlStrDE = Connection.GetConnectionString("DE_Production", "");
                                        if (oDebug)
                                        {
                                            sqlStrDE = Connection.GetConnectionString("DE_Stage", ""); // DE_Production
                                        }

                                        #endregion Pre-Flight
                                        #region Process SQL Command - Try
                                        try
                                        {
                                            #region SQL Connection
                                            using (SqlConnection conDE = new SqlConnection(sqlStrDE))
                                            {
                                                Donation_Open_Database(conDE);
                                                #region SQL Command
                                                using (SqlCommand cmd = new SqlCommand("", conDE))
                                                {
                                                    #region Build SQL Text
                                                    String cmdText = "";
                                                    cmdText += @"
                                                                SET NOCOUNT ON
	                                                                INSERT INTO [dbo].[interactions]
			                                                                ([companyid],[interactiontype],[datestart],[resourcetype],[resourceid],[originator],[destinator],[duration],[status])
		                                                            SELECT
			                                                                @sp_companyid,@sp_interactiontype,@sp_createdate,@sp_resourcetype,@sp_resourceid,@sp_originator,@sp_destinator,@sp_duration,@sp_status

                                                                SELECT SCOPE_IDENTITY()
                                                                SET NOCOUNT OFF
                                                    ";
                                                    #endregion Build SQL Text
                                                    #region SQL Config
                                                    cmd.CommandTimeout = 600;
                                                    cmd.CommandText = cmdText;
                                                    cmd.CommandType = CommandType.Text;
                                                    cmd.Parameters.Clear();
                                                    #endregion SQL Config
                                                    #region SQL Parameters
                                                    cmd.Parameters.Add(new SqlParameter("@sp_companyid", sp_companyid));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_interactiontype", sp_interactiontype));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_createdate", sp_callstart));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_resourcetype", sp_resourcetype));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_resourceid", sp_resourceid));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_originator", rcrd.ani));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_destinator", rcrd.dnis));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_duration", sp_duration));
                                                    cmd.Parameters.Add(new SqlParameter("@sp_status", sp_status));
                                                    #endregion SQL Parameters
                                                    #region SQL SqlDataReader
                                                    var cmdExists = cmd.ExecuteScalar();
                                                    if (cmdExists != null && cmdExists.ToString() != "")
                                                    {
                                                        if (Int32.TryParse(cmdExists.ToString(), out sp_interactionid))
                                                        {
                                                            chckExists = true;
                                                        }
                                                        else
                                                        {
                                                            chckExists = false;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        chckExists = false;
                                                    }
                                                    #endregion SQL SqlDataReader
                                                }
                                                #endregion SQL Command
                                                if (chckExists)
                                                {
                                                    #region SQL Command: interactions_arc
                                                    Donation_Open_Database(conDE);
                                                    using (SqlCommand cmd = new SqlCommand("", conDE))
                                                    {
                                                        #region Build SQL Text
                                                        String cmdText = "";
                                                        cmdText += @"
	                                                        INSERT INTO [dbo].[interactions_arc]
			                                                           ([companyid],[interactionid],[callid],[datestart],[dispositionid],[dispositionname],[offset_current],[offset_original])
		                                                         SELECT
			                                                           @sp_companyid,@sp_interactionid,@sp_arc_callid,@sp_createdate,@sp_arc_disposition_id,@sp_arc_disposition_name,@sp_arc_offset_current,@sp_arc_offset_original
                                                        ";
                                                        #endregion Build SQL Text
                                                        #region SQL Config
                                                        cmd.CommandTimeout = 600;
                                                        cmd.CommandText = cmdText;
                                                        cmd.CommandType = CommandType.Text;
                                                        cmd.Parameters.Clear();
                                                        #endregion SQL Config
                                                        #region SQL Parameters
                                                        cmd.Parameters.Add(new SqlParameter("@sp_companyid", sp_companyid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_interactionid", sp_interactionid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_arc_callid", callid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_createdate", sp_callstart));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_arc_disposition_id", sp_arc_disposition_id));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_arc_disposition_name", sp_arc_disposition_name));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_arc_offset_current", '0'));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_arc_offset_original", '0'));
                                                        #endregion SQL Parameters
                                                        #region SQL SqlDataReader
                                                        var chckExist = cmd.ExecuteNonQuery();
                                                        if (chckExist == 1)
                                                        {
                                                            // Need to determien if we care about this failing here...
                                                            // and how to handle it...
                                                            // interactions_arc = true;
                                                        }
                                                        else
                                                        {
                                                            // interactions_arc = false;
                                                        }
                                                        #endregion SQL SqlDataReader
                                                    }
                                                    #endregion SQL Command: interactions_arc
                                                    #region SQL Command: five9_calls_disposition
                                                    Donation_Open_Database(conDE);
                                                    using (SqlCommand cmd = new SqlCommand("", conDE))
                                                    {
                                                        #region Build SQL Text
                                                        String cmdText = "";
                                                        cmdText += @"
                                    -- We are removing this for now | We need to re-do this if we want to track this data, but it is not *Five9* data

	                                INSERT INTO [dbo].[five9_call_disposition]
			                                    ([companyid],[interactionid],[callid],[dispositionid],[agentid],[datecreated])
		                                    SELECT
			                                    @sp_companyid,@sp_interactionid,@sp_callid,@sp_dispositionid,@sp_agentid,@sp_datecreated

                                        INSERT INTO [dbo].[five9_calls_disposition]
                                                   ([companyid]
                                                   ,[interactionid]
                                                   ,[call.call_id]
                                                   ,[agent.id]
                                                   ,[createdate]
                                                   ,[agent.full_name]
                                                   ,[agent.station_id]
                                                   ,[agent.station_type]
                                                   ,[agent.user_name]
                                                   ,[call.disposition_id]
                                                   ,[call.disposition_name])
                                             SELECT
                                                   @sp_companyid
                                                   ,@sp_interactionid
                                                   ,@sp_call_call_id
                                                   ,@sp_agent_id
                                                   ,@sp_createdate
                                                   ,@sp_agent_full_name
                                                   ,@sp_agent_station_id
                                                   ,@sp_agent_station_type
                                                   ,@sp_agent_user_name
                                                   ,@sp_call_disposition_id
                                                   ,@sp_call_disposition_name

                            ";
                                                        #endregion Build SQL Text
                                                        #region SQL Config
                                                        cmd.CommandTimeout = 600;
                                                        cmd.CommandText = cmdText;
                                                        cmd.CommandType = CommandType.Text;
                                                        cmd.Parameters.Clear();
                                                        #endregion SQL Config
                                                        #region SQL Parameters
                                                        cmd.Parameters.Add(new SqlParameter("@sp_companyid", sp_companyid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_interactionid", sp_interactionid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_call_call_id", callid));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_agent_id", rcrd.agent_id));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_createdate", sp_callstart));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_agent_full_name", sp_agent_full_name));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_agent_station_id", sp_agent_station_id));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_agent_station_type", sp_agent_station_type));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_agent_user_name", sp_agent_user_name));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_call_disposition_id", sp_arc_disposition_id));
                                                        cmd.Parameters.Add(new SqlParameter("@sp_call_disposition_name", sp_arc_disposition_name));
                                                        #endregion SQL Parameters
                                                        #region SQL SqlDataReader
                                                        //We are removing this for now | We need to re-do this if we want to track this data, but it is not *Five9* data
                                                        //
                                                        //var cmdExists = cmd.ExecuteNonQuery();
                                                        //if (cmdExists == 1)
                                                        //{
                                                        //    chckExists = true;
                                                        //}
                                                        //else
                                                        //{
                                                        //    chckExists = false;
                                                        //    throw new Exception("Problem inserting Five9 Disposition");
                                                        //}
                                                        //LogSQL(cmd, "sqlPassed");
                                                        #endregion SQL SqlDataReader
                                                    }
                                                    #endregion SQL Command: five9_calls_disposition
                                                }
                                            }
                                            #endregion SQL Connection
                                            
                                        }
                                        #endregion Process SQL Command - Try
                                        #region Process SQL Command - Catch
                                        catch (Exception ex)
                                        {
                                            callcreateid = -1;
                                            Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                            //LogSQL(cmd, "sqlFailed");
                                        }
                                        #endregion Process SQL Command - Catch
                                        if (sp_interactionid == 0 || sp_interactionid == -1)
                                        {
                                            rcrdStatus = "interactionid.sql";
                                            throw new Exception("There was a problem inserting the [interaction] record {Error: " + sp_interactionid.ToString() + "}.");
                                        }
                                        #endregion Insert: AGENTLOG - DATAEXCHANGE

                                        #region Insert: CALLCREATE
                                        using (SqlCommand cmd = new SqlCommand("", con))
                                        {
                                            #region Populate the SQL Command
                                            cmd.CommandTimeout = 600;
                                            cmd.CommandText = "[dbo].[callprocessing_insert_callcreate]";
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            #endregion Populate the SQL Command
                                            #region Populate the SQL Params
                                            cmd.Parameters.Add(new SqlParameter("@sp_callid", callid));
                                            cmd.Parameters.Add(new SqlParameter("@sp_createdt", sp_callstart));
                                            cmd.Parameters.Add(new SqlParameter("@sp_originationid", rcrd.confirmation));
                                            #endregion Populate the SQL Params
                                            #region Process SQL Command - Try
                                            try
                                            {
                                                callcreateid = Convert.ToInt32(cmd.ExecuteScalar());
                                                LogSQL(cmd, "sqlPassed");
                                            }
                                            #endregion Process SQL Command - Try
                                            #region Process SQL Command - Catch
                                            catch (Exception ex)
                                            {
                                                callcreateid = -1;
                                                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                                LogSQL(cmd, "sqlFailed");
                                            }
                                            #endregion Process SQL Command - Catch
                                        }
                                        if (callcreateid == 0 || callcreateid == -1)
                                        {
                                            rcrdStatus = "callcreate.sql";
                                            throw new Exception("There was a problem inserting the [callcreate] record {Error: " + callcreateid.ToString() + "}.");
                                        }
                                        #endregion Insert: CALLCREATE

                                        #region Insert: STANDARDSELECTION
                                        // NOT NEEDED ????
                                        #endregion Insert: STANDARDSELECTION

                                        #region Insert: IMOIHO
                                        #endregion Insert: IMOIHO

                                        #region Insert: CALLINFO
                                        using (SqlCommand cmd = new SqlCommand("", con))
                                        {
                                            #region Populate the SQL Command
                                            cmd.CommandTimeout = 600;
                                            cmd.CommandText = "[dbo].[callprocessing_insert_callinfo]";
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            #endregion Populate the SQL Command
                                            #region Populate the SQL Params
                                            cmd.Parameters.Add(new SqlParameter("@sp_callid", callid));
                                            cmd.Parameters.Add(new SqlParameter("@sp_fname", rcrd.firstname));
                                            cmd.Parameters.Add(new SqlParameter("@sp_lname", rcrd.lastname));
                                            cmd.Parameters.Add(new SqlParameter("@sp_prefix", rcrd.prefix));
                                            cmd.Parameters.Add(new SqlParameter("@sp_companyyn", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_address", rcrd.address));
                                            cmd.Parameters.Add(new SqlParameter("@sp_suitetype", "")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_suitenumber", rcrd.address2));
                                            cmd.Parameters.Add(new SqlParameter("@sp_zip", rcrd.zip));
                                            cmd.Parameters.Add(new SqlParameter("@sp_city", rcrd.city));
                                            cmd.Parameters.Add(new SqlParameter("@sp_state", rcrd.state));
                                            cmd.Parameters.Add(new SqlParameter("@sp_hphone", rcrd.phone)); // phone,phone_cell,phone_other,phone_business
                                            cmd.Parameters.Add(new SqlParameter("@sp_receiveupdatesyn", rcrd.email_optin)); // Opt In
                                            cmd.Parameters.Add(new SqlParameter("@sp_email", rcrd.email));
                                            cmd.Parameters.Add(new SqlParameter("@sp_companyname", rcrd.business_name));
                                            cmd.Parameters.Add(new SqlParameter("@sp_companytypeid", "")); //??
                                            cmd.Parameters.Add(new SqlParameter("@sp_imoihoyn", "0")); // rcrd.tribute_type)); //??
                                            cmd.Parameters.Add(new SqlParameter("@sp_anonymousyn", "")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_1", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_2", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_3", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_4", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_5", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_6", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_flag_7", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_sendack", "0")); // ??
                                            cmd.Parameters.Add(new SqlParameter("@sp_ackaddress", "")); // ??
                                            #endregion Populate the SQL Params
                                            #region Process SQL Command - Try
                                            try
                                            {
                                                callinfoid = Convert.ToInt32(cmd.ExecuteScalar());
                                                LogSQL(cmd, "sqlPassed");
                                            }
                                            #endregion Process SQL Command - Try
                                            #region Process SQL Command - Catch
                                            catch (Exception ex)
                                            {
                                                callinfoid = -1;
                                                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                                LogSQL(cmd, "sqlFailed");
                                            }
                                            #endregion Process SQL Command - Catch
                                        }
                                        if (callinfoid == 0 || callinfoid == -1)
                                        {
                                            rcrdStatus = "callinfo.sql";
                                            throw new Exception("There was a problem inserting the [callinfo] record {Error: " + callinfoid.ToString() + "}.");
                                        }
                                        #endregion Insert: CALLINFO
                                        if (rcrd.amount != "0.00")
                                        {
                                            insertcc = true;
                                            if (rcrd.card_number.Length > 10)
                                            {
                                                processpayment = true;
                                            }
                                        }
                                        #region Insert: DONATIONCCINFO
                                        // Get DonationCCInfoID
                                        // Get OrderID == DonationCCInfoID PAD RIGHT 14 (0)
                                        // Get Reconcilliation ID == DonationCCInfoID
                                        using (SqlCommand cmd = new SqlCommand("", con))
                                        {
                                            #region Populate the SQL Command
                                            cmd.CommandTimeout = 600;
                                            cmd.CommandText = "[dbo].[callprocessing_insert_donationccinfo]";
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            #endregion Populate the SQL Command
                                            #region Populate the SQL Params
                                            cmd.Parameters.Add(new SqlParameter("@sp_callid", callid));
                                            int cctype = 1;
                                            if (rcrd.card_number.Length > 1)
                                            {
                                                switch (rcrd.card_number.Substring(0, 1))
                                                {
                                                    case "4":
                                                        cctype = 2; // Visa
                                                        break;
                                                    case "5":
                                                        cctype = 3; // MasterCard
                                                        break;
                                                    case "3":
                                                        cctype = 4; // American Express
                                                        break;
                                                    case "6":
                                                        cctype = 5; // Discover
                                                        break;

                                                }
                                            }
                                            cmd.Parameters.Add(new SqlParameter("@sp_cctype", cctype));
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccnum", rcrd.card_number));
                                            string ccnameappear = rcrd.firstname + " " + rcrd.lastname;
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccnameappear", ccnameappear));
                                            if (rcrd.card_exp.Length == 5)
                                            {
                                                string ccexpmonth = rcrd.card_exp.Substring(0, 2);
                                                string ccexpyear = rcrd.card_exp.Substring(3, 2);
                                                cmd.Parameters.Add(new SqlParameter("@sp_ccexpmonth", ccexpmonth));
                                                cmd.Parameters.Add(new SqlParameter("@sp_ccexpyear", ccexpyear));
                                            }
                                            // did dnis
                                            String sp_designationid = "158"; // Disaster Relief | Needs to be based on DID ... FML
                                            if (rcrd.dnis == "2132237045" || rcrd.dnis == "4159675107" || rcrd.dnis == "6179630523" || rcrd.dnis == "5038678816" || rcrd.dnis == "3012004857" || rcrd.dnis == "2132237045" || rcrd.dnis == "4159675107" || rcrd.dnis == "6179630523" || rcrd.dnis == "5038678816" || rcrd.dnis == "3012004857")
                                            { sp_designationid = "187"; } // Hurricane Harvey 
                                            else if (rcrd.dnis == "2069737996" || rcrd.dnis == "6503534788" || rcrd.dnis == "7204392540" || rcrd.dnis == "7738015228" || rcrd.dnis == "9162778261")
                                            { sp_designationid = "188"; }  // Hurricane Irma
                                            else { }

                                            cmd.Parameters.Add(new SqlParameter("@sp_designationid", sp_designationid)); // Hurricane Harvey | Needs to be based on DID ... FML
                                            cmd.Parameters.Add(new SqlParameter("@sp_donationtypeid", (processpayment) ? "1" : "2")); // Valid Donation
                                            cmd.Parameters.Add(new SqlParameter("@sp_donationamount", rcrd.amount));
                                            cmd.Parameters.Add(new SqlParameter("@sp_orderid", callid));
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccflag_1", "0"));
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccflag_2", "0"));
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccflag_3", "0"));
                                            cmd.Parameters.Add(new SqlParameter("@sp_ccchar_1", ""));
                                            #endregion Populate the SQL Params
                                            #region Process SQL Command - Try
                                            try
                                            {
                                                donationccinfoid = Convert.ToInt32(cmd.ExecuteScalar());
                                                LogSQL(cmd, "sqlPassed");
                                            }
                                            #endregion Process SQL Command - Try
                                            #region Process SQL Command - Catch
                                            catch (Exception ex)
                                            {
                                                donationccinfoid = -1;
                                                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                                LogSQL(cmd, "sqlFailed");
                                            }
                                            #endregion Process SQL Command - Catch
                                        }
                                        if (donationccinfoid == 0 || donationccinfoid == -1)
                                        {
                                            rcrdStatus = "donationccinfo.sql";
                                            throw new Exception("There was a problem inserting the [donationccinfo] record {Error: " + donationccinfoid.ToString() + "}.");
                                        }
                                        #endregion Insert: DONATIONCCINFO
                                        if (insertcc)
                                        {
                                            #region Insert: CHARGEDATE
                                            using (SqlCommand cmd = new SqlCommand("", con))
                                            {
                                                #region Populate the SQL Command
                                                cmd.CommandTimeout = 600;
                                                cmd.CommandText = "[dbo].[callprocessing_insert_chargedate]";
                                                cmd.CommandType = CommandType.StoredProcedure;
                                                #endregion Populate the SQL Command
                                                #region Populate the SQL Params
                                                cmd.Parameters.Add(new SqlParameter("@sp_callid", callid));
                                                cmd.Parameters.Add(new SqlParameter("@sp_chargedate1", sp_callstart));
                                                #endregion Populate the SQL Params
                                                #region Process SQL Command - Try
                                                try
                                                {
                                                    chargedateid = Convert.ToInt32(cmd.ExecuteScalar());
                                                    LogSQL(cmd, "sqlPassed");
                                                }
                                                #endregion Process SQL Command - Try
                                                #region Process SQL Command - Catch
                                                catch (Exception ex)
                                                {
                                                    chargedateid = -1;
                                                    Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                                    LogSQL(cmd, "sqlFailed");
                                                }
                                                #endregion Process SQL Command - Catch
                                            }
                                            if (chargedateid == 0 || chargedateid == -1)
                                            {
                                                rcrdStatus = "chargedate.sql";
                                                throw new Exception("There was a problem inserting the [chargedate] record {Error: " + chargedateid.ToString() + "}.");
                                            }
                                            #endregion Insert: CHARGEDATE

                                            if (processpayment)
                                            {
                                                #region Process: CYBERSOURCE
                                                #region Insert: CYBERSOURCE
                                                #endregion Insert: CYBERSOURCE
                                                ARC_Cybersource_Log_Auth arcRecord = new ARC_Cybersource_Log_Auth();
                                                //sendToProduction
                                                RequestMessage request = new RequestMessage();
                                                request.ccAuthService = new CCAuthService();
                                                request.ccAuthService.run = "true";
                                                request.ccCaptureService = new CCCaptureService();
                                                request.ccCaptureService.run = "true";

                                                // Reconcilliation ID from ExternalID / DonationCCInfo.ID
                                                string reconciliationID = donationccinfoid.ToString();
                                                int pad = 16; // 9 for AmEx, 16 for others
                                                if (rcrd.card_number.StartsWith("3")) { pad = 9; }
                                                string reconciliationID1 = reconciliationID.PadRight(pad, '0');
                                                // reconciliationID1 - No Longer Used

                                                request.ccAuthService.reconciliationID = reconciliationID;
                                                request.ccCaptureService.reconciliationID = reconciliationID;

                                                request.merchantReferenceCode = reconciliationID;
                                                BillTo billTo = new BillTo();
                                                billTo.firstName = rcrd.firstname;
                                                billTo.lastName = rcrd.lastname;
                                                billTo.street1 = rcrd.address; // Value but no numeric
                                                billTo.postalCode = rcrd.zip;
                                                billTo.city = rcrd.city;
                                                billTo.state = rcrd.state;
                                                billTo.country = rcrd.country;
                                                if (rcrd.email.Length > 5 && rcrd.email.Contains("@") && rcrd.email.Contains("."))
                                                {
                                                    billTo.email = rcrd.email;
                                                }
                                                else { billTo.email = "nobody@cybersource.com"; }

                                                request.billTo = billTo;
                                                Card card = new Card();
                                                card.accountNumber = rcrd.card_number;
                                                if (rcrd.card_exp.Length == 5)
                                                {
                                                    card.expirationMonth = rcrd.card_exp.Substring(0, 2);
                                                    card.expirationYear = rcrd.card_exp.Substring(3, 2);
                                                }
                                                request.card = card;
                                                PurchaseTotals purchaseTotals = new PurchaseTotals();
                                                purchaseTotals.currency = "USD";
                                                request.purchaseTotals = purchaseTotals;
                                                request.item = new Item[1];
                                                Item item = new Item();
                                                item.id = "0";
                                                item.unitPrice = rcrd.amount;
                                                item.productSKU = "DN001";
                                                item.productName = "ARC Call";
                                                request.item[0] = item;

                                                arcRecord.ExternalID = donationccinfoid.ToString();
                                                try
                                                {
                                                    ReplyMessage reply = SoapClient.RunTransaction(request);
                                                    string template = GetTemplate(reply.decision.ToUpper());
                                                    string content = "";
                                                    try { content = GetContent(reply); }
                                                    catch { content = "error"; }

                                                    Log(logRecord + ",CB: " + String.Format(template, content), "record");

                                                    #region Populate the ARC Record
                                                    if (reply.decision == "ACCEPT") arcRecord.Status = "Settled";
                                                    else if (reply.decision == "REJECT") arcRecord.Status = "Declined";
                                                    else arcRecord.Status = "Error";

                                                    arcRecord.ccContent = content;
                                                    arcRecord.decision = reply.decision;
                                                    arcRecord.merchantReferenceCode = reply.merchantReferenceCode;
                                                    try
                                                    {
                                                        arcRecord.reasonCode = Convert.ToInt32(reply.reasonCode);
                                                    }
                                                    catch { }
                                                    arcRecord.requestID = reply.requestID;
                                                    arcRecord.requestToken = reply.requestToken;
                                                    #region reply.ccAuthReply
                                                    if (reply.ccAuthReply != null)
                                                    {
                                                        arcRecord.ccAuthReply_accountBalance = reply.ccAuthReply.accountBalance;
                                                        //arcRecord.ccAuthReply_accountBalanceCurrency = String.Empty;
                                                        //arcRecord.ccAuthReply_accountBalanceSign = String.Empty;
                                                        arcRecord.ccAuthReply_amount = reply.ccAuthReply.amount;
                                                        arcRecord.ccAuthReply_authFactorCode = reply.ccAuthReply.authFactorCode;
                                                        arcRecord.ccAuthReply_authorizationCode = reply.ccAuthReply.authorizationCode;
                                                        if (reply.ccAuthReply.authorizedDateTime != null)
                                                        {
                                                            arcRecord.ccAuthReply_authorizedDateTime = reply.ccAuthReply.authorizedDateTime.Replace("T", " ").Replace("Z", "");
                                                        }
                                                        arcRecord.ccAuthReply_avsCode = reply.ccAuthReply.avsCode;
                                                        arcRecord.ccAuthReply_avsCodeRaw = reply.ccAuthReply.avsCodeRaw;
                                                        //arcRecord.ccAuthReply_cardCategory = String.Empty;
                                                        arcRecord.ccAuthReply_cavvResponseCode = reply.ccAuthReply.cavvResponseCode;
                                                        arcRecord.ccAuthReply_cavvResponseCodeRaw = reply.ccAuthReply.cavvResponseCodeRaw;
                                                        arcRecord.ccAuthReply_cvCode = reply.ccAuthReply.cvCode;
                                                        arcRecord.ccAuthReply_cvCodeRaw = reply.ccAuthReply.cvCodeRaw;
                                                        arcRecord.ccAuthReply_merchantAdviceCode = reply.ccAuthReply.merchantAdviceCode;
                                                        arcRecord.ccAuthReply_merchantAdviceCodeRaw = reply.ccAuthReply.merchantAdviceCodeRaw;
                                                        //arcRecord.ccAuthReply_ownerMerchantID = String.Empty;
                                                        //arcRecord.ccAuthReply_paymentNetworkTransactionID = String.Empty;
                                                        arcRecord.ccAuthReply_processorResponse = reply.ccAuthReply.processorResponse;
                                                        try
                                                        {
                                                            arcRecord.ccAuthReply_reasonCode = Convert.ToInt32(reply.ccAuthReply.reasonCode);
                                                        }
                                                        catch { }
                                                        arcRecord.ccAuthReply_reconciliationID = reply.ccAuthReply.reconciliationID;
                                                        arcRecord.ccAuthReply_referralResponseNumber = String.Empty;
                                                        arcRecord.ccAuthReply_requestAmount = rcrd.amount;
                                                        arcRecord.ccAuthReply_requestCurrency = String.Empty;
                                                    }
                                                    #endregion reply.ccAuthReply
                                                    #region reply.ccCaptureReply
                                                    if (reply.ccCaptureReply != null)
                                                    {
                                                        arcRecord.ccCaptureReply_amount = reply.ccCaptureReply.amount;
                                                        try
                                                        {
                                                            arcRecord.ccCaptureReply_reasonCode = Convert.ToInt32(reply.ccCaptureReply.reasonCode);
                                                        }
                                                        catch { }
                                                        arcRecord.ccCaptureReply_reconciliationID = reply.ccCaptureReply.reconciliationID;
                                                        arcRecord.ccCaptureReply_requestDateTime = reply.ccCaptureReply.requestDateTime.Replace("T", " ").Replace("Z", "");
                                                    }
                                                    #endregion reply.ccCaptureReply

                                                    #endregion Populate the ARC Record
                                                    //eRecord.STATUS = reply.decision;
                                                    //eRecord.AUTHCODE = (reply.ccAuthReply != null) ? reply.ccAuthReply.authorizationCode : "";
                                                    ////if (reply.ccAuthReply != null)eRecord.AUTHCODE = reply.ccAuthReply.authorizationCode;
                                                    //eRecord.LAS = "0";
                                                    //eRecord.RESPCODE = reply.reasonCode;
                                                    //eRecord.RESPTEXT = content;
                                                    //eRecord.CVVRESP = (reply.ccAuthReply != null) ? reply.ccAuthReply.cvCode : "";
                                                    //eRecord.CVVTEXT = (reply.ccAuthReply != null) ? reply.ccAuthReply.cvCodeRaw : "";
                                                    //eRecord.PROCTID = reply.requestID;
                                                }
                                                catch (Exception ex)
                                                {
                                                    //SaveOrderState();
                                                    //Log(logRecord + ",CB: " + "Error", "record");
                                                    Log("\r\n" + ex.Message + "\r\n" + ex.StackTrace + "\r\n" + ex.Source + "\r\n" + ex.InnerException, "error");
                                                    rcrdStatus = "cybersource.catch";
                                                }
                                                #region Save the record to SQL
                                                if (arcRecord.Status != null)
                                                {
                                                    arcRecord.Source = "CALL";
                                                    ARC_Cybersource_To_SQL(arcRecord);
                                                }
                                                #endregion Save the record to SQL
                                                #endregion Process: CYBERSOURCE
                                            }
                                        }
                                        if (rcrdStatus == string.Empty) { rcrdStatus = "success"; }
                                        #endregion Process Record
                                    }
                                    catch (Exception ex)
                                    {
                                        Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                                        if (rcrdStatus == string.Empty) { rcrdStatus = "catch.error"; }
                                    }
                                    // Log the Record
                                    if (rcrdStatus != "success")
                                    {
                                        logRecord += "," + rcrdStatus;
                                        Log(logRecord, "record");
                                    }
                                    if (pCounts.Records > 5) if (oDebug) break; // DeBug Crap
                                    //break; // DeBug Crap
                                }
                                #region File Cleanup
                                #endregion File Cleanup
                            }
                            #endregion Process Each Record in the File
                        }
                        #endregion SqlConnection
                        #region Move the file
                        bool fileMove = false;
                        try
                        {
                            File.Move(inputPath + @"\" + fName, outputPath + @"\" + fName);
                            fileMove = true;
                        }
                        catch (Exception ex)
                        {
                            Log_Exception("Error 001", ex, "standard", "Failed to move file");
                            //fileMove = false;
                            fileMove = true;
                        }
                        if (!fileMove)
                        {
                            // Sleep for 1.5 second and try again
                            System.Threading.Thread.Sleep(1500); //DeBug
                            File.Move(inputPath + fName, outputPath + fName);
                        }
                        #endregion Rename the file
                        #endregion FileInfo file in files
                        if (oDebug) break; // DeBug Crap
                        //break; // DeBug Crap
                    }
                }
            }
            #endregion Processing Dir of Files - Try
            #region Processing Dir of Files - Catch
            catch (Exception ex)
            {
                pCounts.getError++;
                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
            }
            #endregion Processing Dir of Files - Catch
        }
        private static string GetTemplate(string decision)
        {
            // Retrieves the text that corresponds to the decision.
            if ("ACCEPT".Equals(decision))
            {
                return ("The order succeeded.{0}");
            }
            if ("REJECT".Equals(decision))
            {
                return ("Your order was not approved.{0}");
            }
            // ERROR, or an unknown decision
            return ("Your order could not be completed at this time.{0}" + "Please try again later.");
        }
        private static string GetContent(ReplyMessage reply)
        {
            /*
             * This is where you retrieve the content that will be plugged
             * into the template.
             * 
             * The strings returned in this sample are mostly to demonstrate
             * how to retrieve the reply fields.  Your application should
             * display user-friendly messages.
             */

            int reasonCode = int.Parse(reply.reasonCode);
            switch (reasonCode)
            {
                // Success
                case 100:
                    return ("Approved");
                //"\nRequest ID: " + reply.requestID +
                //"\nAuthorization Code: " +
                //    reply.ccAuthReply.authorizationCode +
                //"\nCapture Request Time: " +
                //    reply.ccCaptureReply.requestDateTime +
                //"\nCaptured Amount: " +
                //    reply.ccCaptureReply.amount);

                // Missing field(s)
                case 101:
                    return (
                        "The following required field(s) are missing: " +
                        EnumerateValues(reply.missingField));

                // Invalid field(s)
                case 102:
                    return (
                        "The following field(s) are invalid: " +
                        EnumerateValues(reply.invalidField));

                // Insufficient funds
                case 204:
                    return (
                        "Insufficient funds in the account.  Please use a " +
                        "different card or select another form of payment.");

                // add additional reason codes here that you need to handle
                // specifically.

                default:
                    // For all other reason codes, return an empty string,
                    // in which case, the template will be displayed with no
                    // specific content.
                    return (String.Empty);
            }
        }
        private static string EnumerateValues(string[] array)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (string val in array)
            {
                sb.Append(val + "");
            }

            return (sb.ToString());
        }
        protected void SaveOrderState()
        {
            Log("SaveOrderState", "record");
            /*

 

             * This is where you store the order state in your system for post-transaction

             * analysis. Be sure to store the consumer information, the values of the reply

             * fields, and the details of any exceptions that occurred.

             */

        }
        protected void ARC_Cybersource_To_SQL(ARC_Cybersource_Log_Auth arcRecord)
        {
            #region Processing Start - SQL - Try
            try
            {
                #region SqlConnection
                string sqlCon = "ARC_Production";
                if (oDebug)
                {
                    sqlCon = "ARC_Stage";
                }
                using (SqlConnection con = new SqlConnection(Connection.GetConnectionString(sqlCon, "")))
                {
                    #region SqlCommand cmd
                    using (SqlCommand cmd = new SqlCommand("", con))
                    {
                        #region Populate the SQL Command
                        cmd.CommandTimeout = 600;
                        cmd.CommandText = "[dbo].[callprocessing_insert_cybersource]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        #endregion Populate the SQL Command
                        #region Populate the SQL Params
                        cmd.Parameters.Add(new SqlParameter("@Source", arcRecord.Source));
                        cmd.Parameters.Add(new SqlParameter("@ExternalID", arcRecord.ExternalID));
                        cmd.Parameters.Add(new SqlParameter("@Status", arcRecord.Status));
                        cmd.Parameters.Add(new SqlParameter("@CreateDate", arcRecord.CreateDate));

                        cmd.Parameters.Add(new SqlParameter("@decision", arcRecord.decision));
                        cmd.Parameters.Add(new SqlParameter("@merchantReferenceCode", arcRecord.merchantReferenceCode));
                        cmd.Parameters.Add(new SqlParameter("@reasonCode", arcRecord.reasonCode));
                        cmd.Parameters.Add(new SqlParameter("@requestID", arcRecord.requestID));
                        cmd.Parameters.Add(new SqlParameter("@requestToken", arcRecord.requestToken));

                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_accountBalance", arcRecord.ccAuthReply_accountBalance));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_accountBalanceCurrency", arcRecord.ccAuthReply_accountBalanceCurrency));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_accountBalanceSign", arcRecord.ccAuthReply_accountBalanceSign));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_amount", arcRecord.ccAuthReply_amount));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_authFactorCode", arcRecord.ccAuthReply_authFactorCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_authorizationCode", arcRecord.ccAuthReply_authorizationCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_authorizedDateTime", arcRecord.ccAuthReply_authorizedDateTime));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_avsCode", arcRecord.ccAuthReply_avsCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_avsCodeRaw", arcRecord.ccAuthReply_avsCodeRaw));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_cardCategory", arcRecord.ccAuthReply_cardCategory));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_cavvResponseCode", arcRecord.ccAuthReply_cavvResponseCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_cavvResponseCodeRaw", arcRecord.ccAuthReply_cavvResponseCodeRaw));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_cvCode", arcRecord.ccAuthReply_cvCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_cvCodeRaw", arcRecord.ccAuthReply_cvCodeRaw));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_merchantAdviceCode", arcRecord.ccAuthReply_merchantAdviceCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_merchantAdviceCodeRaw", arcRecord.ccAuthReply_merchantAdviceCodeRaw));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_ownerMerchantID", arcRecord.ccAuthReply_ownerMerchantID));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_paymentNetworkTransactionID", arcRecord.ccAuthReply_paymentNetworkTransactionID));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_processorResponse", arcRecord.ccAuthReply_processorResponse));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_reasonCode", arcRecord.ccAuthReply_reasonCode));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_reconciliationID", arcRecord.ccAuthReply_reconciliationID));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_referralResponseNumber", arcRecord.ccAuthReply_referralResponseNumber));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_requestAmount", arcRecord.ccAuthReply_requestAmount));
                        cmd.Parameters.Add(new SqlParameter("@ccAuthReply_requestCurrency", arcRecord.ccAuthReply_requestCurrency));
                        cmd.Parameters.Add(new SqlParameter("@ccCaptureReply_amount", arcRecord.ccCaptureReply_amount));
                        cmd.Parameters.Add(new SqlParameter("@ccCaptureReply_reasonCode", arcRecord.ccCaptureReply_reasonCode));
                        cmd.Parameters.Add(new SqlParameter("@ccCaptureReply_reconciliationID", arcRecord.ccCaptureReply_reconciliationID));
                        cmd.Parameters.Add(new SqlParameter("@ccCaptureReply_requestDateTime", arcRecord.ccCaptureReply_requestDateTime));

                        cmd.Parameters.Add(new SqlParameter("@ccContent", arcRecord.ccContent));
                        string cmdText = "\n" + cmd.CommandText;
                        bool cmdFirst = true;
                        foreach (SqlParameter param in cmd.Parameters)
                        {
                            cmdText += "\n" + ((cmdFirst) ? "" : ",") + param.ParameterName + " = " + ((param.Value != null) ? "'" + param.Value.ToString() + "'" : "default");
                            cmdFirst = false;
                        }
                        #endregion Populate the SQL Params
                        #region Process SQL Command - Try
                        try
                        {
                            Donation_Open_Database(con);
                            using (SqlDataReader sqlRdr = cmd.ExecuteReader())
                            {
                                if (sqlRdr.HasRows)
                                {
                                    while (sqlRdr.Read())
                                    {
                                        //arcNewID = sqlRdr["Response"].ToString();
                                    }
                                }
                                else
                                {
                                    //arcNewID = 0;
                                }
                            }
                            Log(cmdText.Replace("\n", " "), "sqlPassed");
                        }
                        #endregion Process SQL Command - Try
                        #region Process SQL Command - Catch
                        catch (Exception ex)
                        {
                            Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
                            Log(cmdText, "sqlFailed");
                        }
                        #endregion Process SQL Command - Catch
                    }
                    #endregion SqlCommand cmd
                }
                #endregion SqlConnection
            }
            #endregion Processing Start - SQL - Try
            #region Processing Start - SQL - Catch
            catch (Exception ex)
            {
                Log_Exception("Error 001", ex, "standard", "Step 1 Catch");
            }
            #endregion Processing Start - SQL - Catch
        }
        public sealed class ARC_Cybersource_Log_Auth
        {
            public String Source;
            public String ExternalID;
            public String Status;
            public String CreateDate;

            public String decision;
            public String merchantReferenceCode;
            public Int32 reasonCode;
            public String requestID;
            public String requestToken;

            public String ccAuthReply_accountBalance;
            public String ccAuthReply_accountBalanceCurrency;
            public String ccAuthReply_accountBalanceSign;
            public String ccAuthReply_amount;
            public String ccAuthReply_authFactorCode;
            public String ccAuthReply_authorizationCode;
            public String ccAuthReply_authorizedDateTime;
            public String ccAuthReply_avsCode;
            public String ccAuthReply_avsCodeRaw;
            public String ccAuthReply_cardCategory;
            public String ccAuthReply_cavvResponseCode;
            public String ccAuthReply_cavvResponseCodeRaw;
            public String ccAuthReply_cvCode;
            public String ccAuthReply_cvCodeRaw;
            public String ccAuthReply_merchantAdviceCode;
            public String ccAuthReply_merchantAdviceCodeRaw;
            public String ccAuthReply_ownerMerchantID;
            public String ccAuthReply_paymentNetworkTransactionID;
            public String ccAuthReply_processorResponse;
            public Int32 ccAuthReply_reasonCode;
            public String ccAuthReply_reconciliationID;
            public String ccAuthReply_referralResponseNumber;
            public String ccAuthReply_requestAmount;
            public String ccAuthReply_requestCurrency;
            public String ccCaptureReply_amount;
            public Int32 ccCaptureReply_reasonCode;
            public String ccCaptureReply_reconciliationID;
            public String ccCaptureReply_requestDateTime;

            public String ccContent;
        }
        [DelimitedRecord("|")]
        [IgnoreFirst(1)]
        [IgnoreEmptyLines]
        public sealed class ARC_CallImport
        {
            //ARC_IVRImport
            //ARC_IVRExport
            public String confirmation;
            public String callcenter;
            public String campaign;
            public String designation;
            public String disposition;
            public String callstart;
            public String callend;
            public String callduration;
            public String timezone;
            public String agent_id;
            public String agent_firstname;
            public String agent_lastname;
            public String agent_notes;
            public String filename;
            public String dnis;
            public String ani;
            public String prefix;
            public String firstname;
            public String middle;
            public String lastname;
            public String suffix;
            public String address;
            public String address2;
            public String city;
            public String state;
            public String zip;
            public String country;
            public String billing;
            public String billing_address;
            public String billing_address2;
            public String billing_city;
            public String billing_state;
            public String billing_zip;
            public String billing_country;
            public String business_match;
            public String business_type;
            public String business_name;
            public String business_address;
            public String business_address2;
            public String business_city;
            public String business_state;
            public String business_zip;
            public String business_country;
            public String phone;
            public String phone_cell;
            public String phone_other;
            public String phone_business;
            public String email;
            public String email_optin;
            public String email_business;
            public String amount;
            public String card_type;
            public String card_name;
            public String card_number;
            public String card_exp;
            public String card_verified;
            public String card_processed;
            public String receipt_method;
            public String process_status;
            public String process_code;
            public String process_message;
            public String process_time;
            public String tribute_type;
            public String tribute_to_firstname;
            public String tribute_to_lastname;
            public String tribute_sender_firstname;
            public String tribute_sender_lastname;
            public String tribute_recipient_firstname;
            public String tribute_recipient_lastname;
            public String tribute_recipient_address;
            public String tribute_recipient_address2;
            public String tribute_recipient_city;
            public String tribute_recipient_saate;
            public String tribute_recipient_zip;
            public String tribute_recipient_country;
            public String tribute_message;
            public String recurring;
            public String recurring_date;
            public String su2c_optin;
            public String celebrity;
        }
        /// <summary>
        /// We Log the loop only once every [config]
        /// And on the start...
        /// </summary>
        public void Log_Running(String msg)
        {
            bool doLog = false;
            if (cntrLoop <= 1) { doLog = true; }
            if (cntrLoop >= 10000) { cntrLoop = 0; }

            if ((DateTime.UtcNow - dtLog).TotalMinutes >= logInterval) { doLog = true; }
            doLog = true;
            if (doLog)
            {
                dtLog = DateTime.UtcNow;
                logfilename = ConfigurationSettings.AppSettings["logfileName"];
                String logStr = "Running [{4}]: {0} | Loops: {1} | Files: {2} | Status: {3}";
                String logStr2 = String.Format(logStr, dtLoop, cntrLoop, cntrFiles, running, msg);
                Log(logStr2, "standard");
            }

        }
        /// <summary>
        /// We only send an email notification once a day if there are no files
        /// Or if there are errors..
        /// </summary>
        /// <param name="msg"></param>
        public void Notification_Check(String msg, String grp)
        {
            Log_Running("Loop-Email-Check");
            if (dtStart == dtEmail)
            {
                SendCountEmail(msg, grp);
                //SendStandardEmail(msg, "Sent", "Report Completed");
                pCounts = new Process_Count();
            }
            else if ((DateTime.UtcNow - dtEmail).TotalHours >= EmailCycle)
            {
                SendCountEmail(msg, grp);
                //SendStandardEmail(msg, "Sent", "Report Completed");
                pCounts = new Process_Count();
            }
            else
            {
                Log_Running("Loop-NotEmailing");
            }
        }
        /// <summary>
        /// Send a standard email
        /// 
        /// </summary>
        /// <param name="emlNotice"></param>
        /// <param name="emlStatus"></param>
        /// <param name="emlMessage"></param>
        public void SendStandardEmail(string emlNotice, string emlStatus, string emlMessage)
        {
            DateTime dt = DateTime.UtcNow;
            string emailSubject = String.Format("{0} ({1} {2})", serviceName, dt.ToShortDateString(), dt.ToShortTimeString());
            // Get the Email Body from a file
            // Hard coded.. but eventually customized?
            System.IO.StreamReader rdr = new StreamReader(System.IO.File.OpenRead(ConfigurationSettings.AppSettings["StandardEmail"]));
            string htmlBody = rdr.ReadToEnd();
            // Process the body variables
            emlNotice = "Attached Custom Report";
            htmlBody = htmlBody.Replace("{EmailTitle}", emailSubject);
            htmlBody = htmlBody.Replace("{Name}", serviceName);
            htmlBody = htmlBody.Replace("{Notice}", emlNotice);
            htmlBody = htmlBody.Replace("{Status}", emlStatus);
            htmlBody = htmlBody.Replace("{Message}", emlMessage);
            htmlBody = htmlBody.Replace("{ScriptTime}", dt.ToString());
            // Determine the TimeZone the server/script machine is on
            // This will help the recipients determine the relevant time against their own timezone
            System.TimeZone localZone = System.TimeZone.CurrentTimeZone;
            //GMT -8:00 Pacific Time (US &amp; Canada)
            if (localZone.IsDaylightSavingTime(dt))
            {
                htmlBody = htmlBody.Replace("{ScriptTimeZone}", String.Format("GMT {0} {1}", localZone.GetUtcOffset(dt), localZone.DaylightName));
            }
            else
            {
                htmlBody = htmlBody.Replace("{ScriptTimeZone}", String.Format("GMT {0} {1}", localZone.GetUtcOffset(dt), localZone.StandardName));
            }
            //DateTime.UtcNow.ToString("zz")

            SendEmail_Config toConfig = new SendEmail_Config();
            toConfig.Subject = emailSubject;
            toConfig.Body = htmlBody;
            toConfig.Status = emlStatus;
            toConfig.DistList = "EmailGroup1";
            toConfig.Priority = 3;

            // Disabled due to auth // SendEmail(toConfig);

        }
        /// <summary>
        /// Send a Count email
        /// </summary>
        /// <param name="emlNotice"></param>
        public void SendCountEmail(string emlNotice, string emlGroup)
        {
            try
            {
                SendEmail_Config toConfig = new SendEmail_Config();

                DateTime dt = DateTime.UtcNow;
                dtEmail = dt;
                string emailSubject = String.Format("{0} ({1} {2})", serviceName, dt.ToShortDateString(), dt.ToShortTimeString());
                // Get the Email Body from a file
                // Hard coded.. but eventually customized?
                String emailPath = ConfigurationSettings.AppSettings["servicePath"];
                String emailFile = ConfigurationSettings.AppSettings["CountEmail"];
                System.IO.StreamReader rdr = new StreamReader(System.IO.File.OpenRead(emailPath + emailFile));
                string htmlBody = rdr.ReadToEnd();
                // Process the body variables
                emlNotice = "Attached Custom Report";
                htmlBody = htmlBody.Replace("{EmailTitle}", emailSubject);
                htmlBody = htmlBody.Replace("{Name}", serviceName);
                htmlBody = htmlBody.Replace("{Notice}", emlNotice);
                htmlBody = htmlBody.Replace("{FileCount}", pCounts.Records.ToString());
                htmlBody = htmlBody.Replace("{ScriptTime}", dt.ToString());
                htmlBody = htmlBody.Replace("{GMT_Offset}", "GMT -7:00");

                String responseTableRow = "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>\n\r";

                String responseHTML_Header = String.Format(responseTableRow
                    , "Start" // 0
                    , "Loops" // 1
                    , "Records" // 2
                    , "RecordsErr" // 3
                    , "Errors" // 4
                    , "getError" // 5
                    , "End" // 6
                    );

                String flResponseHTML = String.Format(responseTableRow
                    , pCounts.Start
                    , pCounts.Loops
                    , pCounts.Records
                    , pCounts.RecordsErr
                    , pCounts.Errors
                    , pCounts.getError
                    , DateTime.UtcNow
                    );

                htmlBody = htmlBody.Replace("{HTML_Response_Header}", responseHTML_Header);
                htmlBody = htmlBody.Replace("{HTML_Response}", flResponseHTML);

                // Determine the TimeZone the server/script machine is on
                // This will help the recipients determine the relevant time against their own timezone
                System.TimeZone localZone = System.TimeZone.CurrentTimeZone;
                // GMT -8:00 Pacific Time (US &amp; Canada)
                if (localZone.IsDaylightSavingTime(dt))
                {
                    htmlBody = htmlBody.Replace("{ScriptTimeZone}", String.Format("GMT {0} {1}", localZone.GetUtcOffset(dt), localZone.DaylightName));
                }
                else
                {
                    htmlBody = htmlBody.Replace("{ScriptTimeZone}", String.Format("GMT {0} {1}", localZone.GetUtcOffset(dt), localZone.StandardName));
                }
                //DateTime.UtcNow.ToString("zz")

                toConfig.Subject = emailSubject;
                toConfig.Body = htmlBody;
                toConfig.Status = "Count";
                toConfig.DistList = emlGroup; //"EmailGroup1";
                //toConfig.DistListGroup = emlGroup; //AMDA

                if (pCounts.Records == 0) { toConfig.Priority = 1; }
                else { toConfig.Priority = 3; }
                Log_Running("Loop-Email-Send");
                // Disabled due to auth // SendEmail(toConfig);
            }
            catch (Exception ex)
            {
                pCounts.getError++;
                Log_Exception("Error 006", ex, "standard", "SendCountEmail Catch");
            }
        }
        /// <summary>
        /// Send a Count email
        /// </summary>
        /// <param name="emlNotice"></param>
        public sealed class SendEmail_Config
        {
            public String Subject;
            public String Body;
            public String Status;
            public int Priority;
            public String DistList;
            public String DistListGroup;
            public String BounceAddress;
            public String FromName;
            public String FromAddress;
            public String ReplyTo;
            public String HeaderField;
            public String HeaderValue;
            public String ToName;
            public String ToAddress;
            public bool Log = true;
            public String LogFile = "standard";
            public bool FileAttached;
            public String FileName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="emlConfig"></param>
        public void SendEmail(SendEmail_Config emlConfig)
        {
            try
            {
                //We send an email to the notification people..
                // Create a mailman object for sending email.
                Chilkat.MailMan mailman = new Chilkat.MailMan();
                // Unlock the component
                mailman.UnlockComponent(ConfigurationSettings.AppSettings["chlk_unlock"].ToString());
                // Set the SMTP server hostname.
                //mailman.SmtpHost = Connection.GetSmtpHost();
                mailman.SmtpHost = ConfigurationSettings.AppSettings["chlk_mailman_smtphost"].ToString(); // "smtp.office365.com";
                mailman.SmtpPort = Convert.ToInt32(ConfigurationSettings.AppSettings["chlk_mailman_smtpport"].ToString()); // 587;
                mailman.SmtpUsername = ConfigurationSettings.AppSettings["chlk_mailman_smtpusername"].ToString(); // "do.not.reply@email.com";
                mailman.SmtpPassword = ConfigurationSettings.AppSettings["chlk_mailman_smtppassword"].ToString(); // "Tr011t0ll";
                mailman.StartTLS = Convert.ToBoolean(ConfigurationSettings.AppSettings["chlk_mailman_starttls"].ToString());
                // Create a simple HTML email.
                Chilkat.Email email = new Chilkat.Email();
                email.Subject = emlConfig.Subject;

                email.SetHtmlBody(emlConfig.Body);
                if (emlConfig.Priority == 1) { email.AddHeaderField("X-Priority", "1 (High)"); }
                if (emlConfig.Priority == 3) { email.AddHeaderField("X-Priority", "3 (Normal)"); }
                if (emlConfig.Priority == 5) { email.AddHeaderField("X-Priority", "5 (Low)"); }

                // Load the Distribution
                // Generally:
                // EmailGroup1 == MIS
                // EmailGroup2 == Internal/Client
                // EmailGroup3 == Campaign Specific
                int distList = 0;
                ARC_CallProcessing_CyberSource.DistributionListConfig configInfo = (ARC_CallProcessing_CyberSource.DistributionListConfig)ConfigurationManager.GetSection("DistributionList");
                if (emlConfig.DistList == "EmailGroup1")
                {
                    #region Process EmailGroup1
                    if (configInfo.EmailGroup1.Count != 0)
                    {
                        foreach (ARC_CallProcessing_CyberSource.EmailElement elem in configInfo.EmailGroup1)
                        {
                            distList++;
                            string toName = String.Format("{0} ({1})", elem.DisplayName, elem.Address);
                            if (elem.AddType == "CC")
                            { email.AddCC(toName, elem.Address); }
                            else if (elem.AddType == "BCc")
                            { email.AddBcc(toName, elem.Address); }
                            else
                            { email.AddTo(toName, elem.Address); }
                        }
                    }
                    #endregion Process EmailGroup1
                }
                if (emlConfig.DistList == "EmailGroup2")
                {
                    #region Process EmailGroup2
                    if (configInfo.EmailGroup2.Count != 0)
                    {
                        foreach (ARC_CallProcessing_CyberSource.EmailElement elem in configInfo.EmailGroup2)
                        {
                            distList++;
                            string toName = String.Format("{0} ({1})", elem.DisplayName, elem.Address);
                            if (elem.AddType == "CC") { email.AddCC(toName, elem.Address); }
                            else if (elem.AddType == "BCc") { email.AddBcc(toName, elem.Address); }
                            else { email.AddTo(toName, elem.Address); }
                        }
                    }
                    #endregion Process EmailGroup2
                }
                if (emlConfig.DistList == "EmailGroup3")
                {
                    #region Process EmailGroup3
                    if (configInfo.EmailGroup3.Count != 0)
                    {
                        foreach (ARC_CallProcessing_CyberSource.EmailElement elem in configInfo.EmailGroup3)
                        {
                            if (emlConfig.DistListGroup == elem.Group || elem.Group == "Default")
                            {
                                distList++;
                                string toName = String.Format("{0} ({1})", elem.DisplayName, elem.Address);
                                if (elem.AddType == "CC") { email.AddCC(toName, elem.Address); }
                                else if (elem.AddType == "BCc") { email.AddBcc(toName, elem.Address); }
                                else { email.AddTo(toName, elem.Address); }
                            }
                        }
                    }
                    #endregion Process EmailGroup3
                }

                #region Default Email Address(s)
                if (distList == 0)
                {
                    //Hard Coded Name Value
                    //We don't have an address, do a hard coded one..?
                    //Default...?
                    email.AddTo("Hard Coded (pciambotti@email.com)", "pciambotti@email.com");
                }
                #endregion

                email.BounceAddress = ConfigurationManager.AppSettings["BounceAddress"];
                email.FromName = ConfigurationManager.AppSettings["FromName"];
                email.FromAddress = ConfigurationManager.AppSettings["FromAddress"];
                email.ReplyTo = ConfigurationManager.AppSettings["ReplyTo"];

                email.AddHeaderField(ConfigurationManager.AppSettings["HeaderField_Name"], ConfigurationManager.AppSettings["HeaderField_Value"]);
                if (emlConfig.FileAttached)
                {
                    if (oDebug) { Log_Running("Email Attachment: " + emlConfig.FileName); }
                    try
                    {
                        email.AddFileAttachment(emlConfig.FileName);
                    }
                    catch (Exception ex)
                    {
                        pCounts.getError++;
                        Log_Exception("Error 008", ex, "standard", "SendEmail Add File");
                    }
                }


                // Send mail.
                bool success;
                success = mailman.SendEmail(email);
                if (success)
                {
                    Log("     Email Sent: " + emlConfig.Status, "standard");
                }
                else
                {
                    Log("     Email Failed: " + emlConfig.Status + "\r\n" + mailman.LastErrorText.ToString(), "standard");
                }
                mailman.Dispose();
                email.Dispose();
            }
            catch (Exception ex)
            {
                pCounts.getError++;
                Log_Exception("Error 007", ex, "standard", "SendEmail Catch");
            }
        }
        /// <summary>
        /// Handle the Log Writing...
        /// </summary>
        /// <param name="content"></param>
        protected void Donation_Open_Database(SqlConnection con)
        {
            bool trySql = true;
            while (trySql)
            {
                try
                {
                    if (con.State != ConnectionState.Open) { con.Close(); con.Open(); }
                    trySql = false;
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("timeout") || ex.Message.ToLower().Contains("time out"))
                    {
                        // Pause .5 seconds and try again
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
                        // throw the exception
                        
                        if (oDebug) Log("\r\nSQL Error\r\n" + con.ConnectionString.ToString() + "\r\n" + Connection.userIP(), "error");
                        trySql = false;
                        throw ex;
                    }
                }
            }
        }
        public void LogSQL(SqlCommand cmd, string type)
        {
            //LogSQL(cmd, "sqlPassed");
            //LogSQL(cmd, "sqlFailed");
            //Log(cmdText, "sqlFailed");
            //Log(cmdText, "sqlPassed");
            string cmdText = "\n" + cmd.CommandText;
            bool cmdFirst = true;
            foreach (SqlParameter param in cmd.Parameters)
            {
                cmdText += "\n" + ((cmdFirst) ? "" : ",") + param.ParameterName + " = " + ((param.Value != null) ? "'" + param.Value.ToString() + "'" : "default");
                cmdFirst = false;
            }
            if (type == "sqlFailed")
            {
                Log(cmdText, type);
            }
            else
            {
                Log(cmdText.Replace("\n", " "), type);
            }
        }
        public void Log(string content, string lf)
        {
            if (lf == "standard")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileName"]; //If File - Name
            }
            else if (lf == "files")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameFiles"]; //If File - Name
            }
            else if (lf == "record")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameRecord"]; //If File - Name
            }
            else if (lf == "error")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameError"]; //If File - Name
            }
            else if (lf == "sleep")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameSleep"]; //If File - Name
            }
            else if (lf == "sqlPassed")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameSQLPassed"]; //If File - Name
            }
            else if (lf == "sqlFailed")
            {
                logfilename = ConfigurationSettings.AppSettings["logfileNameSQLFailed"]; //If File - Name
            }
            else
            {
                logfilename = ConfigurationSettings.AppSettings["logfileName"]; //Default
            }
            string[] tmp = new string[] { content };
            Log(tmp);
        }
        public void Log(string[] content)
        {
            if (logging != "" && logging != null)
            {
                if (logging == "FILE")
                    WriteToFile(content);
                else
                    WriteToEventLog(content);
            }

        }
        public void WriteToEventLog(string[] content)
        {
            foreach (string item in content)
            {
                EventLog.WriteEntry(this.ServiceName, PrepareLine(item), EventLogEntryType.Information, 0);
            }
        }
        public void WriteToFile(string[] content)
        {
            if (System.IO.File.Exists(logfilepath + logfilename))
            {
                System.IO.FileInfo log = new System.IO.FileInfo(logfilepath + logfilename);
                if (log.Length > logfilemaxsize) log.Delete();
            }
            System.IO.StreamWriter logger = System.IO.File.AppendText(logfilepath + logfilename);
            foreach (string item in content)
            {
                logger.Write(PrepareLine(item));
            }
            logger.Close();
        }
        public string PrepareLine(string line)
        {
            //return System.DateTime.UtcNow.ToUniversalTime().ToString() + ": " + line + "\n";
            return System.DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss tt") + ": " + line + "\n";
        }
        public void Log_Exception(String Error, Exception ex, String File, String Msg)
        {
            errMsg.Count++;
            errMsg.Error = Error;
            errMsg.Message = ex.Message;
            errMsg.Source = ex.Source;
            errMsg.StackTrace = ex.StackTrace;
            errMsg.TimeStamp = DateTime.UtcNow;

            string msg = String.Format("{0}\n\r{1}\n\r{2}\n\r{3}\n\r{4}"
                , errMsg.Error
                , ex.Message
                , ex.Source
                , ex.StackTrace
                , lnBreak);
            Log(msg, "error");
            if (Msg.Length > 0)
            {
                msg = String.Format("{0}|{1}", Error, Msg);
                Log(msg, File);
            }

        }
    }
}
