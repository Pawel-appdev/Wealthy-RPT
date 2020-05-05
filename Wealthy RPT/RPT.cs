using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Wealthy_RPT
{
    class RPT
    {

        public struct RPT_Data
        {

            private Int32 _cuid;
            private string _segment;
            //private string _fullname;
            private string _surname;
            private string _firstname;
            private double _utr;
            private DateTime _dob;
            private byte _deceased;
            private DateTime _dod;
            private DateTime _deselected;
            private string _marital;
            private string _gender;
            private string _mainadd;
            private string _mainpc;
            private string _secadd;
            private string _residence;
            private string _domicile;
            private string _office;
            private string _team;
            private string _wealthlevel;
            private string _pathway;
            private string _source;
            private string _sector;
            private string _longterm;
            private string _lifeevents;
            private string _narrative;
            private Int32 _hnwupid;
            private string _pop;
            private string _crmname;
            private DateTime _crmappointed;
            private DateTime _cdlu;

            public Int32 CU_ID
            {
                get
                {
                    return _cuid;
                }
                set
                {
                    _cuid = value;
                }
            }

            public string Segment
            {
                get
                {
                    return _segment;
                }
                set
                {
                    _segment = value;
                }
            }

            //public string FullName
            //{
            //    get
            //    {
            //        return _fullname;
            //    }
            //    set
            //    {
            //        _fullname = value;
            //    }
            //}

            public string Surname
            {
                get
                {
                    return _surname;
                }
                set
                {
                    _surname = value;
                }
            }

            public string Firstname
            {
                get
                {
                    return _firstname;
                }
                set
                {
                    _firstname = value;
                }
            }

            public double UTR
            {
                get
                {
                    return _utr;
                }
                set
                {
                    _utr = value;
                }
            }

            public DateTime DOB
            {
                get
                {
                    return _dob;
                }
                set
                {
                    _dob = value;
                }
            }

            public byte Deceased
            {
                get
                {
                    return _deceased;
                }
                set
                {
                    _deceased = value;
                }
            }

            public DateTime DOD
            {
                get
                {
                    return _dod;
                }
                set
                {
                    _dod = value;
                }
            }

            public DateTime Deselected
            {
                get
                {
                    return _deselected;
                }
                set
                {
                    _deselected = value;
                }
            }

            public string Marital
            {
                get
                {
                    return _marital;
                }
                set
                {
                    _marital = value;
                }
            }

            public string Gender
            {
                get
                {
                    return _gender;
                }
                set
                {
                    _gender = value;
                }
            }

            public string MainAdd
            {
                get
                {
                    return _mainadd;
                }
                set
                {
                    _mainadd = value;
                }
            }

            public string MainPC
            {
                get
                {
                    return _mainpc;
                }
                set
                {
                    _mainpc = value;
                }
            }

            public string SecAdd
            {
                get
                {
                    return _secadd;
                }
                set
                {
                    _secadd = value;
                }
            }

            public string Residence
            {
                get
                {
                    return _residence;
                }
                set
                {
                    _residence = value;
                }
            }

            public string Domicile
            {
                get
                {
                    return _domicile;
                }
                set
                {
                    _domicile = value;
                }
            }

            public string Office
            {
                get
                {
                    return _office;
                }
                set
                {
                    _office = value;
                }
            }

            public string Team
            {
                get
                {
                    return _team;
                }
                set
                {
                    _team = value;
                }
            }

            public string WealthLevel
            {
                get
                {
                    return _wealthlevel;
                }
                set
                {
                    _wealthlevel = value;
                }
            }

            public string Pathway
            {
                get
                {
                    return _pathway;
                }
                set
                {
                    _pathway = value;
                }
            }

            public string Source
            {
                get
                {
                    return _source;
                }
                set
                {
                    _source = value;
                }
            }

            public string Sector
            {
                get
                {
                    return _sector;
                }
                set
                {
                    _sector = value;
                }
            }

            public string LongTerm
            {
                get
                {
                    return _longterm;
                }
                set
                {
                    _longterm = value;
                }
            }

            public string LifeEvents
            {
                get
                {
                    return _lifeevents;
                }
                set
                {
                    _lifeevents = value;
                }
            }

            public string Narrative
            {
                get
                {
                    return _narrative;
                }
                set
                {
                    _narrative = value;
                }
            }


            public int HNWUPID
            {
                get
                {
                    return _hnwupid;
                }
                set
                {
                    _hnwupid = value;
                }
            }

            public string Pop
            {
                get
                {
                    return _pop;
                }
                set
                {
                    _pop = value;
                }
            }

            public string CRM_Name
            {
                get
                {
                    return _crmname;
                }
                set
                {
                    _crmname = value;
                }
            }

            public DateTime CRM_Appointed
            {
                get
                {
                    return _crmappointed;
                }
                set
                {
                    _crmappointed = value;
                }
            }

            public DateTime CDLU
            {
                get
                {
                    return _cdlu;
                }
                set
                {
                    _cdlu = value;
                }
            }


            public void GetRPDData(double dblUTR)
            {

                SqlConnection con = new SqlConnection(Global.ConnectionString);
                //SqlCommand cmd = new SqlCommand("qryGetRPDData", con);
                SqlCommand cmd = new SqlCommand("qryGetCustomerData", con);
                cmd.Parameters.Clear();
                SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                prm01.Value = dblUTR;
                cmd.CommandTimeout = Global.TimeOut;
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                #region Recordset
                if (dr.HasRows)
                {
                    DateTime dtmin = new DateTime(1900, 1, 1);
                    DateTime dtActual;
                    dr.Read();
                    CU_ID = Convert.ToInt32(dr["CU_ID"]);
                    Surname = dr["Surname"].ToString();
                    Firstname = dr["FirstName"].ToString();
                    UTR = double.Parse(dr["UTR"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                    // ###### NEED NULL DATES #####
                    try { dtActual = Convert.ToDateTime(dr["DOB"]).Date; } catch { dtActual = dtmin; }
                    DOB = dtActual;
                    Deceased = Convert.ToByte(dr["Deceased"]);
                    try { dtActual = Convert.ToDateTime(dr["DOD"]).Date; } catch { dtActual = dtmin; }
                    DOD = dtActual;
                    try { dtActual = Convert.ToDateTime(dr["Deselected"]).Date; } catch { dtActual = dtmin; }
                    Deselected = dtActual;
                    Marital= dr["Marital"].ToString();
                    Gender = dr["Gender"].ToString();
                    MainAdd = dr["MainAdd"].ToString();
                    MainPC = dr["MainPC"].ToString();
                    SecAdd = dr["SecAdd"].ToString();
                    Residence = dr["Residence"].ToString();
                    Domicile = dr["Domicile"].ToString();
                    Office = dr["Office"].ToString();
                    Team = dr["Team"].ToString();
                    WealthLevel = dr["WealthLevel"].ToString();
                    Pathway = dr["Pathway"].ToString();
                    Source = dr["Source"].ToString();
                    Sector = dr["Sector"].ToString();
                    LongTerm = dr["LongTerm"].ToString();
                    LifeEvents = dr["LifeEvents"].ToString();
                    Narrative = dr["Narrative"].ToString();
                    HNWUPID = Convert.ToInt32(dr["HNWUPID"]);
                    Pop = dr["Pop"].ToString();
                    CRM_Name = dr["CRM_Name"].ToString();
                    // ###### NEED NULL DATES #####
                    try { dtActual = Convert.ToDateTime(dr["CRM_Appointed"]).Date; } catch { dtActual = dtmin; }
                    CRM_Appointed = dtActual;
                    try { dtActual = Convert.ToDateTime(dr["CDLU"]).Date; } catch { dtActual = dtmin; }
                    CDLU = dtActual;

                }
                else if (dr.HasRows == false)
                {
                    MessageBox.Show("Record not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                }

                #endregion
                con.Close();
            }

            public void GetEmailAddress(double dblUTR)
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                //SqlCommand cmd = new SqlCommand("qryGetEmailAddresses", con);
                //cmd.Parameters.Clear();
                //SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                //prm01.Value = dblUTR;
                //cmd.CommandTimeout = Global.TimeOut;
                //cmd.CommandType = CommandType.StoredProcedure;
                //con.Open();


                //////SqlConnection con = new SqlConnection();
                SqlDataAdapter ad = new SqlDataAdapter();
                SqlCommand cmd = new SqlCommand();
                String str = "SELECT Contact, Email_Address, Contact_Role, Contact_Date_Added, Email_ID FROM tblEmailAddress";
                cmd.CommandText = str;
                ad.SelectCommand = cmd;
                ////////con.ConnectionString = "Data Source=localhost; Initial Catalog=Northwind; Integrated Security=True";
                cmd.Connection = con;
                DataSet ds = new DataSet();
                ad.Fill(ds);


                RPT_Detail rptDet = new RPT_Detail();  // initialise form



                rptDet.lvwEmail.DataContext = ds.Tables[0].DefaultView;
                con.Close();











                //////SqlDataReader sqlReader = cmd.ExecuteReader();


                //////while (sqlReader.Read())
                //////{
                //////    ListViewItem lvItem = new ListViewItem();
                //////    lvItem.SubItems[0].Text = sqlReader[0].ToString();
                //////    lvItem.SubItems.Add(sqlReader[1].ToString());

                //////    RPT_Detail rptDet = new RPT_Detail();  // initialise form
                //////    rptDet.lvwEmail.Items.Add(lvItem);

                //////    //listView1.Items.Add(lvItem);
                //////}
















                //SqlDataAdapter da = new SqlDataAdapter(cmd);
                ////DataTable dt = new DataTable("Test");
                ////da.Fill(dt);

                //DataSet ds = new DataSet();
                //da.Fill(ds);


                //RPT_Detail rptDet = new RPT_Detail();  // initialise form

                ////rptDet.lvwEmail.DataContext = ds.Tables[0].DefaultView;
                //rptDet.lvwEmail.ItemsSource = ds.Tables[0].DefaultView;


                //rptDet.lvwEmail.View = View.Details;<<<<<<<<<<<<<<<<<<<<<

                ////List<User> items = new List<User>();

                //DataTable dt = new DataTable();
                //da.Fill(dt);

                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    DataRow dr = dt.Rows[i];
                //    ListViewItem listitem = new ListViewItem(dr["Email_ID"].ToString());
                //    listitem.SubItems.Add(dr["Email_ID"].ToString());
                //    listitem.SubItems.Add(dr["Email_Address"].ToString());
                //    listitem.SubItems.Add(dr["Contact_Role"].ToString());
                //    rptDet.lvwEmail.Items.Add(listitem);
                //}

            }

        }



    }
}
