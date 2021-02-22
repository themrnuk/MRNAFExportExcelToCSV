﻿using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Security;
using System.Text;

namespace MRNAFExportExcelToCSV
{
    public class MSPApi
    {
        public string ApiName { get; set; }
        public string DeltaColumn { get; set; }
        public string ModifiedDate { get; set; }

    }
    public class ProjectData
    {
        public string name { get; set; }
        public string url { get; set; }
    }

    public class MSPResponseDataApi
    {
        public string JsonFileName  { get; set; }
    }
    public class MSPCredential
    {
        internal static readonly string BaseUrl = "https://themrn.sharepoint.com";
        private static readonly string userName = "mspsyncaccount@themrn.co.uk";//"mahavir.rawat@themrn.co.uk";
        private static readonly string password = "Knowledgeable050221&";//"Celebrated161220%"
        private static SecureString secureString   // property
        {
            get
            {
                var securePassword = new SecureString();
                foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);
                return securePassword;
            }

        }

        internal static SharePointOnlineCredentials Credentials
        {
            get
            {
                return new SharePointOnlineCredentials(userName, secureString);
            }
        }

    }

    
}
