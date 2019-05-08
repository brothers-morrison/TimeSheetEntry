using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using static MercyShipsTimeEntry.TimeSheetPerson;
using System.Linq;

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
            const string DELIMITER = "|";
            var xls = @"C:\Users\morrisow\OneDrive\Documents\00 Work Docs\Copy-of-471.xlsx";
            var helper = new XLSXHelper();
            var emptyPerson = new TimeSheetPerson();
            
            var output = helper.GetCroppedExcelFile(xls, emptyPerson, DELIMITER, crop: false);

            var people = SplitOnPerson(output, DELIMITER);
            Console.WriteLine("people: {0}", people.Count);
            foreach (var person in people)
            {
                var croppedChunk = helper.GetCroppedExcelFile(xls, person, DELIMITER, crop: true);
                person.FillValuesFromTimeSheetChunk(croppedChunk);
                Console.WriteLine(person.ToString());
            }
        }

        
    }
}
