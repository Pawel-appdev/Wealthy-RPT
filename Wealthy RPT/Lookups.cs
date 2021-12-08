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
        public DataSet ds = new DataSet();
        public DataSet dsRPTDetailCombo = new DataSet();
        public DataSet dsRPTDetailOfficeCRMs = new DataSet();
        public DataSet dsRPTDetailPopulations = new DataSet();
        public DataSet dsOfficeTeams = new DataSet();
        public DataSet dsOffices= new DataSet();
        public DataSet dsOfficeTeamStaff = new DataSet();

        public void GetMainLookups()
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryFrmMainCombo";
                        SqlParameter prm = cmd.Parameters.Add("@nPID", SqlDbType.Int);
                        prm.Value = Global.PID;
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(ds);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryFrmMainCombo'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

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

        public void GetRPTDetailOfficeCRMs(string strOffice)
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        //cmd.CommandText = "qryGetCRMs";
                        cmd.CommandText = "qryGetOfficeCRMs";
                        SqlParameter prm = cmd.Parameters.Add("@nOffice", SqlDbType.Text);
                        prm.Value = strOffice;
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

        public void GetOfficeTeams(string strOffice, string strPopCode)
        {
            // returns matching tblOfficeCRM:  Office, [Team Identifier], Pop
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetOfficeTeams";
                        SqlParameter prm1 = cmd.Parameters.Add("@nOffice", SqlDbType.Text);
                        prm1.Value = strOffice;
                        SqlParameter prm2 = cmd.Parameters.Add("@nPop", SqlDbType.Text);
                        prm2.Value = strPopCode;
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


        public void GetOffices(string strPopCode)
        {

            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetOffices";
                        SqlParameter prm1 = cmd.Parameters.Add("@nPop", SqlDbType.Text);
                        prm1.Value = strPopCode;  
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsOffices);
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

        public void GetOfficeTeamStaff(string strOffice, string strTeam)
        {
            // returns formatted data from tblUsers
            try
            {
                using (SqlConnection con = new SqlConnection(Global.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "qryGetOfficeTeamStaff";
                        SqlParameter prm1 = cmd.Parameters.Add("@nOffice", SqlDbType.Text);
                        prm1.Value = strOffice;
                        SqlParameter prm2 = cmd.Parameters.Add("@nTeam", SqlDbType.Text);
                        prm2.Value = strTeam;
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;

                        con.Open();

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        adapter.Fill(dsOfficeTeamStaff);
                        con.Close();
                    }
                }
            }
            catch (SqlException e)
            {
                //con.Close();
                MessageBox.Show("Unable to retrieve drop-down lists using: 'qryGetTeams'." + Environment.NewLine + Environment.NewLine
                    + "Error: " + e.Number + Environment.NewLine + e.Message
                    , Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

    }
}
