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

        public TimeSheetPerson(TimeSheetPerson cloneTarget)
        {
            this.Clone(cloneTarget);
        }

        /// <summary>
        /// Given a chunk of pipe-delimited text cells from MercyShips IS Timesheet,
        /// Read this format into an object.
        /// </summary>
        internal void FillValuesFromTimeSheetChunk(string output, char delim = '|')
        {
            var lines = output.Split("\r\n");
            // First row we're getting is Person's Name, so Init with Name first...
            var stickyRowType = RowType.Name;
            foreach (var line in lines)
            {
                var columns = line.Split(delim);
                stickyRowType = FillSinglePerson(columns, stickyRowType);
            }
        }

        internal static Dictionary<string, string> GetSingleColumnExcelDict(string[] columns)
        {
            List<string> columnList = columns.ToList();
            if (columnList.All(val => String.IsNullOrEmpty(val)))
            {
                // Don't pay any mind to ALL BLANK rows,just skip it.
                return new Dictionary<string, string>();
            }

            //Console.Write("columns:{0} \n", columns.Length);
            return ToExcelAbcDict(columns);
        }
        internal RowType FillSinglePerson(string[] columns, RowType stickyRowType)
        {
            var Cells = GetSingleColumnExcelDict(columns);
            RowType currentRowType = RowType.Unknown;
            if (TryMatchRowType(Cells, out currentRowType) 
                && currentRowType != RowType.Unknown)
            {
                stickyRowType = currentRowType;
            }
            RowTypeToSetValues(stickyRowType, Cells);
            return stickyRowType;
        }

        public static Dictionary<string, string> ToExcelAbcDict(string[] columns)
        {
            var Cells = new Dictionary<string, string>(
                            // I'm not sure why, but column "A" doesn't actually show up
                            // Hence the [i+1] here.  We're going to just fake it and not stick anything into "A"
                            // This causes the rest of everything to line up with Excel column names.
                            Alphabet.Skip(1).Select((abc, i) => new KeyValuePair<string, string>(abc, i<columns.Length?columns[i]:String.Empty))
                        );
            return Cells;
        }

        public static List<string> Alphabet = new List<string> {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"
        };
        public static Dictionary<string, RowType> RowTypeDescriptions = new Dictionary<string, RowType>(
            Enum.GetValues(typeof(RowType)).ToList<RowType>()
            .Select(num =>
                new KeyValuePair<string, RowType>(num.ToDescriptionString(), num)));

        private bool TryMatchRowType(Dictionary<string, string> cells, out RowType rowType)
        {
            string signalToken = String.Empty;
            rowType = RowType.Unknown;
            foreach (var col in cells.Where(c => !String.IsNullOrWhiteSpace(c.Value)))
            {
                if (RowTypeDescriptions.Keys.Contains(col.Value))
                {
                    rowType = RowTypeDescriptions[col.Value];
                    return true;
                }
            }
            return false;
        }

        private bool StillOnFirstRow(RowType rt, Dictionary<string, string> Cells)
        {
            return Cells[HomeRow] == rt.ToDescriptionString();
        }
        public readonly static string HomeRow = "C";
        public readonly static string DaysBeginRow = "K";

        public void RowTypeToSetValues(RowType rt, Dictionary<string, string> Cells)
        {
            // if the home row is the "Decorator row" then return.
            // TBD fix this
            if (!Cells.Any() || string.IsNullOrWhiteSpace(Cells[HomeRow]))
            {
                return;
            }
            switch (rt)
            {
                case RowType.Name:
                    this.Clone(NameRowToPerson(Cells));
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

        private void Clone(TimeSheetPerson cloneTarget)
        {
            Name = cloneTarget.Name;
            Id = cloneTarget.Id;
            RoleCode = cloneTarget.RoleCode;
            Department = cloneTarget.Department;
            DeptCode = cloneTarget.DeptCode;
        }

        public enum RowType
        {
            // Note: the NAME won't match for other people, so we need another identifier.
            [Description("Name")]
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


        public TimeSheetPerson(string name = "")
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
            GetNonZeroActivities().ForEach(act => sb.AppendFormat("{0} \n", act.ToString()));
            return sb.ToString();
        }

        private List<ActivityTimeRecordRow> GetNonZeroActivities()
        {
            var results = new List<ActivityTimeRecordRow>();
            foreach ((var act, Int32 i) in Activities
                            .Where(pair => pair.Value.TimeRecord.Sum(hrs => hrs.Value) > 0)
                            .Select((value, i) => (value, i)))
            {
                results.Add(act.Value);
            }
            return results;
        }

        public string ToShortString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0}:{1} \n", nameof(Name), Name);
            sb.AppendFormat("Non-Zero Activities: {0} \n", GetNonZeroActivities().Count);
            return sb.ToString();
        }



            public static List<TimeSheetPerson> SplitOnPerson(string output, string delim = "|")
        {
            var results = new List<TimeSheetPerson>();
            //Console.WriteLine(output);

            // First row we're getting is Person's Name
            var lines = output.Split("\r\n");
            var stickyRowType = RowType.Name; // Init with Name first...
            foreach (var line in lines)
            {
                var columns = line.Split(delim);
                var person = new TimeSheetPerson();
                if (IsPersonRow(columns, stickyRowType, out person))
                {
                    results.Add(person);
                }
            }
            return results;
        }

        private static bool IsPersonRow(string[] columns, RowType stickyRowType, out TimeSheetPerson person)
        {
            var Cells = GetSingleColumnExcelDict(columns);
            person = new TimeSheetPerson();
            if(!Cells.Any())
            {
                return false;
            }
            person = NameRowToPerson(Cells);
            int Id = -1;
            int roleCode = -1;

            // There is one case that passes the test below - DATES row.
            double dateNotName = 0;
            
            if(double.TryParse(person.Name, out dateNotName)
                && DateTime.FromOADate(dateNotName).Month == DateTime.Now.Month)
            {
                // This is a Date row, not a person row.
                return false;
            }

            try
            {
                if (!String.IsNullOrWhiteSpace(person.Name)
                    && Int32.Parse(person.Id) > 0
                    && !String.IsNullOrWhiteSpace(person.Role)
                    && Int32.Parse(person.RoleCode) > 0
                    && !String.IsNullOrWhiteSpace(person.Department)
                    && !String.IsNullOrWhiteSpace(person.DeptCode)
                )
                {
                    return true;
                }
            }
            catch(Exception ex)
            {
                // It's ok
            }
            return false;
        }
        private static TimeSheetPerson NameRowToPerson(Dictionary<string, string> Cells)
        {
            var person = new TimeSheetPerson();
            person.Name = Cells[HomeRow];
            person.Id = Cells["K"];
            person.Role = Cells["Q"];
            person.RoleCode = Cells["Z"];
            person.Department = Cells["AD"];
            person.DeptCode = Cells["AM"];
            return person;
        }
    }
}
