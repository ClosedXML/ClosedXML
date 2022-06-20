using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class Database
    {
        public static void Register(CalcEngine ce)
        {
            //ce.RegisterFunction("DAVERAGE", 3, Daverage); // Returns the average of selected database entries
            //ce.RegisterFunction("DCOUNT", 1, Dcount); // Counts the cells that contain numbers in a database
            //ce.RegisterFunction("DCOUNTA", 1, Dcounta); // Counts nonblank cells in a database
            //ce.RegisterFunction("DGET", 1, Dget); // Extracts from a database a single record that matches the specified criteria
            //ce.RegisterFunction("DMAX", 1, Dmax); // Returns the maximum value from selected database entries
            //ce.RegisterFunction("DMIN", 1, Dmin); // Returns the minimum value from selected database entries
            //ce.RegisterFunction("DPRODUCT", 1, Dproduct); // Multiplies the values in a particular field of records that match the criteria in a database
            //ce.RegisterFunction("DSTDEV", 1, Dstdev); // Estimates the standard deviation based on a sample of selected database entries
            //ce.RegisterFunction("DSTDEVP", 1, Dstdevp); // Calculates the standard deviation based on the entire population of selected database entries
            //ce.RegisterFunction("DSUM", 1, Dsum); // Adds the numbers in the field column of records in the database that match the criteria
            //ce.RegisterFunction("DVAR", 1, Dvar); // Estimates variance based on a sample from selected database entries
            //ce.RegisterFunction("DVARP", 1, Dvarp); // Calculates variance based on the entire population of selected database entries
        }

        static object Daverage(List<Expression> p)
        {
            var b = true;
            foreach (var v in p)
            {
                b = b && (bool)v;
            }
            return b;
        }
    }
}
