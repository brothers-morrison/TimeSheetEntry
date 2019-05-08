using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace MercyShipsTimeEntry
{
    public class ActivityTimeRecordRow
    {
        public Activity Act;
        public Dictionary<DateTime, Double> TimeRecord;

        public ActivityTimeRecordRow(Activity act, Dictionary<string, string>Cells)
        {
            Act = act;
            TimeRecord = new Dictionary<DateTime, Double>();
            // Assuming A-J are "meta info" or column names
            // Days of the month start on "K" in spreadsheet.  K = 11 in ABC
            // So we will skip the first 10 rows and start on 11
            int skipCells = 9;
            int dayOfMonth = 1;
            int maxDaysInMonth = 31;
            
            var month = DateTime.Now.Month;
            var year = DateTime.Now.Year;
            var daysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
            // Read across the row horizontally, skipping the first chunk until you get to days of the month.
            foreach ((var abcValPair, int i) in Cells.Skip(skipCells).Select((val, i) => (val, i)))
            {
                if (i+1 == maxDaysInMonth) {// TOTALS columns ?SAVE?
                }
                if (i+1 >= maxDaysInMonth)
                {
                    return;
                }
                // Starting on "K", each day of the month is read into a DateTime that represents this date
                // Zero based iterator, but calendar is 1 based, so we add (i+1) here
                var dateStr = String.Format("{0}/{1}/{2}", (i+1).ToString("00"), month.ToString("00"), year);
                var hours = 0.0d;
                DateTime date;
                Double.TryParse(abcValPair.Value, out hours);
                DateTime.TryParseExact(dateStr, "dd/MM/yyyy", 
                    CultureInfo.InvariantCulture, 
                    DateTimeStyles.AdjustToUniversal, out date);
                TimeRecord.Add(
                    date, 
                    hours);
                // i == day of the Month
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0}, total_hours:{1}", Act, TimeRecord.Sum(pair => pair.Value));
            return sb.ToString();
        }
    }
}
