using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Wealthy_RPT
{
    class RAG
    {
        public static double RAG2M_1;
        public static double RAG2M_2;
        public static double RAG10M_1;
        public static double RAG10M_2;
        public static double RAG20M_1;
        public static double RAG20M_2;

        public static void GetRAGBreaks()
        {
            SqlConnection conn = new SqlConnection(Global.ConnectionString);
            SqlDataAdapter da = new SqlDataAdapter("qryGetRAGBreaks", conn);
            DataSet ds = new DataSet();
            da.Fill(ds, "RAGBreaks");

            DataTable dt = new DataTable();

            dt = ds.Tables["RAGBreaks"];

            

            foreach (DataRow dr in dt.Rows)
            {

                String strPop = dr["Pop"].ToString().Replace(" ",string.Empty);

                switch (strPop)
                {

                    case "rPt2Mill":
                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "GreenAmber") { RAG2M_1 = Convert.ToDouble(dr["Percent"]); }
                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "AmberRed") { RAG2M_2 = Convert.ToDouble(dr["Percent"]); }
                        break;

                    case "rPt10Mill":

                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "GreenAmber") { RAG10M_1 = Convert.ToDouble(dr["Percent"]); }
                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "AmberRed") { RAG10M_2 = Convert.ToDouble(dr["Percent"]); }
                        break;

                    case "rPt20Mill":
                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "GreenAmber") { RAG20M_1 = Convert.ToDouble(dr["Percent"]); }
                        if (dr["RAG_Break"].ToString().Replace(" ", string.Empty) == "AmberRed") { RAG20M_2 = Convert.ToDouble(dr["Percent"]); }
                        break;

                }
            }
        }

    }
}
