using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeNotiPj
{
    class config
    {
        public static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return string.Empty;
            }
        }
    }
}
