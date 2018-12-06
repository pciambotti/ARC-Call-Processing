using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace ARC_CallProcessing_CyberSource
{
    /// <summary>
    /// This is the connection class
    /// Here we will have any type of connection
    /// Based on the server the app is being run from, the connection will change
    /// WAN is typically a dev enviroment, where the app is connecting to a remote server
    /// LAN is typically a live enviroment, where the app is connecting to a server on it's private net
    /// </summary>
    public class Connection
    {
        static string myHost = System.Net.Dns.GetHostName();
        static string myIP = System.Net.Dns.GetHostEntry(myHost).AddressList[0].ToString();
        static public string userIP() { return myIP; }

        static public string GetConnectionType()
        {
            /// <summary>
            /// Connection Type (ie LAN or WAN)
            /// </summary>
            //if (myHost.Contains("gh-702843f7e966") || myIP.Contains("192.168.2")) { return "Local"; }
            //else if (myIP.Contains("66.135.60.218") || myIP.Contains("172.16.14.109")) { return "LiveOld"; }
            //else { return "Live"; }

            string rtrn = myHost + "|" + myIP;
            return rtrn;
        }
        /// <summary>
        /// Standard Database Connection
        /// Will use LAN if local host is detected
        /// Will use WAN if no local host
        /// </summary>
        /// <returns></returns>
        static public String GetConnectionString(String CN_ID_Name, String CN_Source)
        {
            String CN_ID = "";
            String DB_Server = "";
            String DB_Name = "";
            String DB_UserName = "";
            String DB_Password = "";
            if (CN_Source != "LAN" && CN_Source != "WAN")
            {
                CN_Source = "WAN";
                //if (myIP.Contains("172.16.0.")) { CN_Source = "LAN"; }
            }
            if (myIP.Contains("172.16.0.") && CN_Source == "WAN") { CN_Source = "LAN"; }
            CN_Source = "LAN"; // Need to fix this...

            if (CN_ID_Name == "Default")
            {
                CN_ID = ConfigurationManager.AppSettings["DB_CN_DEFAULT"];
            }
            else
            {
                CN_ID = ConfigurationManager.AppSettings[CN_ID_Name];
            }
            #region Connection #
            DB_Server = ConfigurationManager.AppSettings["DB" + CN_ID + "_" + CN_Source];
            DB_Name = ConfigurationManager.AppSettings["DB" + CN_ID + "_NAME"];
            DB_UserName = ConfigurationManager.AppSettings["DB" + CN_ID + "_USER"];
            DB_Password = ConfigurationManager.AppSettings["DB" + CN_ID + "_PASS"];
            #endregion Connection #

            String DB_Connection = String.Format("Server={0};Database={1};Uid={2};Pwd={3};MultipleActiveResultSets={4}"
                , DB_Server // 0 - Server
                , DB_Name // 1 - Database
                , DB_UserName // 2 - Uid
                , DB_Password // 3 - Pwd
                , "True" // 4 - MultipleActiveResultSets
                );

            return DB_Connection;
        }
        static public string GetSmtpHost()
        {
            /// <summary>
            /// SMTP Connection
            /// </summary>
            String HS_LAN = ConfigurationManager.AppSettings["HS_LAN"];
            String HS_WAN = ConfigurationManager.AppSettings["HS_WAN"];
            if (myHost.Contains("gh-702843f7e966") || myIP.Contains("192.168.2")) { return HS_LAN; } else { return HS_WAN; }
        }

    }
    /// <summary>
    /// The below classes handle the Email Distribution from the Config File
    /// This was created by Milbonn Yaya and copied over
    /// This is used in the service via: LMS_Visions_File_Processing.EmailElement elem in configInfo.EmailGroup1
    /// </summary>
    public class EmailElement : ConfigurationElement
    {

        [ConfigurationProperty("address", IsRequired = true)]
        public string Address
        {
            get { return this["address"] as string; }
        }
        [ConfigurationProperty("name", IsRequired = true)]
        public string DisplayName
        {
            get { return this["name"] as string; }
        }
        [ConfigurationProperty("type", IsRequired = true)]
        public string AddType
        {
            get { return this["type"] as string; }
        }
        [ConfigurationProperty("key", IsRequired = true)]
        public string Key
        {
            get { return this["key"] as string; }
        }
        [ConfigurationProperty("group", IsRequired = false)]
        public string Group
        {
            get { return this["group"] as string; }
        }
    }
    public class EmailGroupCollection : ConfigurationElementCollection
    {
        public EmailElement this[int index]
        {
            get
            {
                return base.BaseGet(index) as EmailElement;
            }
            set
            {
                if (base.BaseGet(index) != null)
                {
                    base.BaseRemoveAt(index);
                }
                this.BaseAdd(index, value);
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new EmailElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((EmailElement)element).Key;
        }
    }
    public class DistributionListConfig : ConfigurationSection
    {

        [ConfigurationProperty("name", DefaultValue = "Default", IsRequired = false)]
        public string Name
        {
            get
            {
                return this["name"] as string;
            }
        }
        [ConfigurationProperty("EmailGroup1")]
        public EmailGroupCollection EmailGroup1
        {
            get
            {
                return this["EmailGroup1"] as EmailGroupCollection;
            }
        }
        [ConfigurationProperty("EmailGroup2")]
        public EmailGroupCollection EmailGroup2
        {
            get
            {
                return this["EmailGroup2"] as EmailGroupCollection;
            }
        }
        [ConfigurationProperty("EmailGroup3")]
        public EmailGroupCollection EmailGroup3
        {
            get
            {
                return this["EmailGroup3"] as EmailGroupCollection;
            }
        }
    }
}
