using System;
using System.Linq.Expressions;
using System.Reflection;

namespace MercyShipsTimeEntry
{
    /// <summary>
    /// Needs to solve:
    /// 1. Can I input my time for today {"IGGN":"1", "IFDV":"7" }
    /// 2. Does each day's total >= 8 hours?
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var xls = @"C:\Users\morrisow\OneDrive\Documents\00 Work Docs\Copy-of-471.xlsx";
            var helper = new XLSXHelper();

            var Me = new TimeSheetPerson("Jacob Perkins");
            var output = helper.GetCroppedExcelFile(xls, Me);
            //Console.Write(output);

            Me.FillValuesFromTimeSheetChunk(output);
            Console.WriteLine(Me.ToString());
        }
    }
}
