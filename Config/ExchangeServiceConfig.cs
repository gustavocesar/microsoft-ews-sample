using System.Collections.Generic;

namespace EWS_Sample.Config
{
    public static class ExchangeServiceConfig
    {
        //Email config
        public static bool IsGenericBox = false;
        
        //Used when IsGenericBox is set TRUE
        public static string HostGenericBox = "https://SOME-ENTERPRISE-EXCHANGE-HOST.asmx";

        //Used when IsGenericBox is set FALSE
        public static string HostOffice365 = "https://outlook.office365.com/EWS/Exchange.asmx";

        public static string Username { get; set; } = "YOUR_EMAIL@outlook.com";
        public static string Password { get; set; } = "YOUR_PASSWORD";
        public static bool UseDefaultCredentials = false;

        //Search config
        public static int PageSize = 10;
        public static int Offset = 0;

        public static List<string> FindSubjects = new List<string>
        {
            "ADD FIND SUBJECTS HERE...",
        };
    }
}