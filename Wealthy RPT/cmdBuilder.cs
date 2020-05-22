using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Wealthy_RPT
{
    class cmdBuilder
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = Admin.strAddlDataInstr;// "SELECT * FROM [tblAdditional_Data_Sources]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }
}
