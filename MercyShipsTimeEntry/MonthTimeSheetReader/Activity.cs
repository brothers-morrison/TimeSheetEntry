using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace MercyShipsTimeEntry
{
    public enum ActivityType
    {
        [Description("")]
        BLANK,
        [Description("Maintaining")]
        MN,
        [Description("Customer")]
        CI,
        [Description("ResearchConcept")]
        RC,        
        [Description("Enhancing")]
        EN,
        [Description("Protecting")]
        PR,
        [Description("Project Work")]
        PJ,
    }

    
    public enum StatusType
    {
        [Description("")]
        BLANK,
        [Description("Approved")]
        APPROVED,
        [Description("Complete")]
        COMPLETE      
    }

    public enum LocationCode
    {
        [Description("")]
        Blank = 0,
        [Description("AfricaMercy")]
        AFM = 3,
        [Description("GlobalMercy")]
        GLM = 4,
        [Description("USOfficeOnly")]
        USO = 9,
        [Description("InternationalSupportCenter")]
        ISC = 10        
    }

    public class Activity
    {
        public String Description = "";
        public String ShortName = "";
        public ActivityType Act = ActivityType.BLANK;
        public LocationCode Loc = LocationCode.Blank;
        public StatusType Status = StatusType.BLANK;
        public const string STATUS_BLANK = "BLANK";

        public Activity(
            string shortname, 
            string activityType, 
            string location, 
            string description,
            string statusType = STATUS_BLANK)
        {
            ShortName = shortname;
            Act = (ActivityType)Enum.Parse(typeof(ActivityType), activityType.ToUpper());
            Loc = (LocationCode)Int32.Parse(location);
            Status = (StatusType)Enum.Parse(typeof(StatusType), statusType.ToUpper());
            Description = description;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("  {0}:{1}", ShortName, Description);
            return sb.ToString();
        }
    }
}
