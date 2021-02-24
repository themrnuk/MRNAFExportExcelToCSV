using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    static class JiraAPICredentials
    {
        public static string APIBaseUrl { get; } = "https://themrn.atlassian.net/rest/api/latest/";
        public static string UserName { get; } = "shivam.kumar@rsk-bsl.com";

        public static string AccessToken { get; } = "L3SIKGO84gMq9Svj9wJL153A";

        public static string ZephyrAPIBaseUrl = "https://api.adaptavist.io/tm4j/v2";

        public static string ZephyrAccessToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIyNTkyMTk5NC1iY2UzLTMyNTgtYTIxYS01YTg0NDI3NmUwYjAiLCJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczpcL1wvdGhlbXJuLmF0bGFzc2lhbi5uZXQiLCJ1c2VyIjp7ImFjY291bnRJZCI6IjVmMDMxMzE2YjU0NWUyMDAxNTRkNWNkYyJ9fSwiaXNzIjoiY29tLmthbm9haC50ZXN0LW1hbmFnZXIiLCJleHAiOjE2NDU0NzMzMDAsImlhdCI6MTYxMzkzNzMwMH0.4yN3CFhaNC059WMMft86knWygSmnkSKcRrQ6BjTrzks";
    }
}
