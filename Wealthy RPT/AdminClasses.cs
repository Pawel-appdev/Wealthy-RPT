using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Wealthy_RPT
{
    class tblAdditional
    {
            public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

            public static string SQLquery= Admin.strAddlDataInstr;// "SELECT * FROM [tblAdditional_Data_Sources]";

            public static SqlCommand myCmd = new SqlCommand(SQLquery , conn);
            public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
            public static DataTable dt = new DataTable("Tables");
    }

    class tblAgent
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblAgent_Details] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblInfo
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * From [tblInfo]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblAssoc
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblAssociation_Types]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblCombos
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblCombos]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblCRM_W
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblCRM_Weighting]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblCust
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblCustomer_Data] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblEmail
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblEmailAddress] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblEnq
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblEnquiry_Data] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblOffice
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblOfficeCRM]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblQW
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblQuestion_Weighting] Order By [Control_Name], [Score_entered]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblField
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblFieldNames]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblCalc
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblRPD_Calculations]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblScore
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblRPD_Score_Data] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblSensPID
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblSensitive_Cases_List] Order By [Authorised_User_PID]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblSensUTR
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblSensitive_Cases_List] Order By [UTR]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblRep
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblStndReports]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblUser
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [tblUsers]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblUserPopID
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [dbo].[ltUserPopID] ORDER BY UserPID, UserPopID";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }

    class tblSector
    {
        public static SqlConnection conn = new SqlConnection(Global.ConnectionString);

        public static string SQLquery = "SELECT * FROM [dbo].[ltSector]";

        public static SqlCommand myCmd = new SqlCommand(SQLquery, conn);
        public static SqlDataAdapter sda = new SqlDataAdapter(myCmd);
        public static DataTable dt = new DataTable("Tables");
    }
}
