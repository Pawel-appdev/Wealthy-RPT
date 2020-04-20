using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Windows.Forms;

namespace Wealthy_RPT
{
    class Lookups
    {
        SqlDataAdapter adapter = new SqlDataAdapter();
        public DataSet dsRPTDetailCombo = new DataSet();
        public DataSet dsRPTDetailOfficeCRMs = new DataSet();
        public DataSet dsRPTDetailPopulations = new DataSet();
        public DataSet dsOfficeTeams = new DataSet();

        public void GetRPTDetailLookups()
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryFrmDetailRPTCombo";
                        SqlParameter prm = cmd.Parameters.Add("@nUserPID", SqlDbType.Int);
                        prm.Value = Global.PID;
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsRPTDetailCombo);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryFrmDetailRPTCombo'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

        public void GetRPTDetailOfficeCRMs()
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        //cmd.CommandText = "qryFrmDetailsOfficeCRMs";
                        cmd.CommandText = "qryGetCRMs";
                        //SqlParameter prm = cmd.Parameters.Add("@nOffice", SqlDbType.Text);
                        //prm.Value = "";
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsRPTDetailOfficeCRMs);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryFrmDetailsOfficeCRMs'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

        public void GetRPTDetailPopulations()
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetPopulations";
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsRPTDetailPopulations);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryFrmDetailsOfficeCRMs'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

        public void GetOfficeCRMs()
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetOfficeTeams";
                        SqlParameter prm1 = cmd.Parameters.Add("@nOffice", SqlDbType.Text);
                        prm1.Value = "East Kilbride";  // ######################
                        SqlParameter prm2 = cmd.Parameters.Add("@nPop", SqlDbType.Text);
                        prm2.Value = "rPt20Mill";  // ######################
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsOfficeTeams);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryGetOfficeTeams'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

    }
}
