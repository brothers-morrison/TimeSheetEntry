using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace MercyShipsTimeEntry
{
    public class TimeSheetPerson
    {
        public string Name = "Unknown Person";
        public string Id = "-1";
        public string Role = "Unknown";
        public string RoleCode = "0000";
        public string Department = "Unknown";
        public string DeptCode = "-1";

        public List<string> Routines = new List<string>();
        public List<string> Projects = new List<string>();
        public List<string> NonISHours = new List<string>();

        public Dictionary<Activity, ActivityTimeRecordRow> Activities = new Dictionary<Activity, ActivityTimeRecordRow>();

        /// <summary>
        /// Given a chunk of pipe-delimited text cells from MercyShips IS Timesheet,
        /// Read this format into an object.
        /// </summary>
        internal void FillValuesFromTimeSheetChunk(string output, char delim = '|')
        {
            // First row we're getting is Person's Name
            var lines = output.Split("\r\n");
            var stickyRowType = RowType.Name; // Init with Name first...
            foreach(var line in lines)
            {
                var columns = line.Split(delim);
                var columnList = columns.ToList();
                if(columnList.All(val => String.IsNullOrEmpty(val)))
                {
                    // Don't pay any mind to ALL BLANK rows,just skip it.
                    continue;
                }

                //Console.Write("columns:{0} \n", columns.Length);
                var Cells = new Dictionary<string, string>(
                    // I'm not sure why, but column "A" doesn't actually show up
                    // Hence the [i+1] here.  We're going to just fake it and not stick anything into "A"
                    // This causes the rest of everything to line up with Excel column names.
                    columns.Select((val, i) => new KeyValuePair<string, string>(Alphabet[i+1], val))
                );

                var currentRowType = GetRowType(columnList);
                if(currentRowType != RowType.Unknown)
                {
                    stickyRowType = currentRowType;
                }
                //if(columnList.Count < 1)
                //{
                //    stickyRowType = RowType.Unknown;
                //}
                RowTypeToSetValues(stickyRowType, Cells);
            }            
            
        }

        List<string> Alphabet = new List<string> {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"
        };        
        public static Dictionary<string, RowType> Descriptions = new Dictionary<string, RowType>(
            Enum.GetValues(typeof(RowType)).ToList<RowType>()
            .Select(num => 
                new KeyValuePair<string, RowType>(num.ToDescriptionString(), num)));

        private RowType GetRowType(List<string> columns)
        {
            string signalToken = String.Empty;
            foreach (var col in columns.Where(c => !String.IsNullOrWhiteSpace(c)))
            {
                if (Descriptions.Keys.Contains(col))
                {
                    return Descriptions[col];
                }
            }
            return RowType.Unknown;
        }

        private bool StillOnFirstRow(RowType rt, Dictionary<string, string> Cells)
        {
            return Cells[HomeRow] == rt.ToDescriptionString();
        }
        public readonly string HomeRow = "C";

        public void RowTypeToSetValues(RowType rt, Dictionary<string, string> Cells)
        {
            // if the home row is the "Decorator row" then return.
            // TBD fix this
            if (string.IsNullOrWhiteSpace(Cells[HomeRow]))
            {
                return;
            }
            switch (rt)
            {
                case RowType.Name:
                    this.Name = Cells[HomeRow];
                    this.Id = Cells["K"];
                    this.Role = Cells["Q"];
                    this.RoleCode = Cells["Z"];
                    this.Department = Cells["AD"];
                    this.DeptCode = Cells["AM"];
                    break;
                case RowType.Routine:
                    if (StillOnFirstRow(rt, Cells)) { return; }
                    var act = new Activity(
                        Cells[HomeRow], 
                        Cells["D"], 
                        Cells["E"], 
                        Cells["F"]);
                    this.Activities.Add(act, new ActivityTimeRecordRow(act, Cells));
                    break;
                case RowType.Projects:
                    if (StillOnFirstRow(rt, Cells)) { return; }
                    var act1 = new Activity(
                        Cells[HomeRow],
                        ActivityType.PJ.ToString(),
                        ((int)LocationCode.ISC).ToString(),
                        Cells["F"],
                        Cells["G"]);
                    this.Activities.Add(act1, new ActivityTimeRecordRow(act1, Cells));
                    break;
                case RowType.NonISHours:
                    if (StillOnFirstRow(rt, Cells)) { return; }
                    var act2 = new Activity(
                        Cells[HomeRow], 
                        ActivityType.BLANK.ToString(), 
                        ((int)LocationCode.Blank).ToString(), 
                        Cells["F"]);
                    this.Activities.Add(act2, new ActivityTimeRecordRow(act2, Cells));
                    break;
            }
        }

        public enum RowType
        {
            // TBD this won't work for other people...
            [Description("Will Morrison")]
            Name,
            [Description("Routine (Service/Activity/Location)")]
            Routine,
            [Description("Routine IS Hours")]            
            RoutineTotals,
            [Description("Projects")]            
            Projects,
            [Description("Project IS Hours")]
            ProjectTotals,
            [Description("Non IS Hours")]
            NonISHours,
            [Description("Remote Working")]
            RemoteWorking,
            [Description("Unknown")]
            Unknown
        }


        public TimeSheetPerson(string name)
        {
            Name = name;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0}:{1} \n", nameof(Name), Name);
            sb.AppendFormat("{0}:{1} \n", nameof(Id), Id);
            sb.AppendFormat("{0}:{1} \n", nameof(Role), Role);
            sb.AppendFormat("{0}:{1} \n", nameof(RoleCode), RoleCode);
            sb.AppendFormat("{0}:{1} \n", nameof(Department), Department);
            sb.AppendFormat("{0}:{1} \n", nameof(DeptCode), DeptCode);

            // Show me all the Activities that are non-zero, and their sums
            sb.AppendFormat("Activities: \n");
            foreach ((var act, Int32 i) in Activities
                .Where(pair => pair.Value.TimeRecord.Sum(hrs => hrs.Value) > 0)
                .Select((value, i) => (value, i)))
            {
                sb.AppendFormat("{0} \n", act.Value);
            }
            return sb.ToString();
        }
    }
}
