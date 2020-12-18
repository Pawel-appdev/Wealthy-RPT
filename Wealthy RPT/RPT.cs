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
            #region propertydeclarations
            // Customer Data
            private Int32 _cuid;
            private string _strand;
            private string _segment;
            //private string _fullname;
            private string _surname;
            private string _firstname;
            private double _utr;
            private string _dob;
            private byte _deceased;
            private string _dod;
            private string _deselected;
            private string _marital;
            private string _gender;
            private string _mainadd;
            private string _mainpc;
            private string _secadd;
            private string _residence;
            private string _domicile;
            private string _office;
            private string _team;
            private string _allocatedto;
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
            private string _crmappointed;
            private string _cdlu;
            // Agent Data
            private Int32 _agentrecordid;
            private string _agent;
            private string _agentcode;
            private byte _agent648held;
            private string _agentaddress;
            private string _namedagent;
            private string _othercontact;
            private string _agenttelno;
            private byte _changed;
            // Behavious Data
            private Int32 _rpdid;
            private string _updateddate;
            private Int32 _updatedby;
            private int _calendaryear;
            private int _lpopen;
            private int _lpclosed;
            private double _hppenalty;
            private byte _pscurrent;
            private byte _psprevious;
            private byte _psfailures;
            private int _qsscore;
            private int _rptprscore;
            private int _rptavscore;
            private int _cgscore;
            private int _resscore;
            private int _crmscore;
            private int _priorityscore;
            private float _percentile;
            private string _supercededdate;
            private byte _riskingcomplete;
            private string _segmentrecorded;
            private string _rsdlu;
            private string _crmexplanation;
            private DataTable _historicaldata;
            private DataTable _griddata;
            private DataTable _chartdata;
            // CRMM enquiry data
            private int _risksopen;
            private int _settledrisks;
            private float _highestsettlement;
            private int _highestpercentage;
            //private string _crmmdata;
            //private float _crmmscore;
            private string _crmmdateadded;

            #endregion

            #region readwriteproperites
            // Customer Data
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

            public string Strand
            {
                get
                {
                    return _strand;
                }
                set
                {
                    _strand = value;
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

            public string DOB
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

            public string DOD
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

            public byte DeselectedTF        /* DeselectedTrueFalse for Deselected checkbox*/
            {
                get
                {
                    return Convert.ToDateTime(Deselected) == null ? Convert.ToByte(0) : Convert.ToByte(1);
                }
            }

            public string Deselected
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

            public string AllocatedTo
            {
                get
                {
                    return _allocatedto;
                }
                set
                {
                    _allocatedto = value;
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

            public string CRM_Appointed
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

            public string CDLU
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

            // Agent Data

            public Int32 AgentRecordID
            {
                get
                {
                    return _agentrecordid;
                }
                set
                {
                    _agentrecordid = value;
                }
            }

            public string Agent
            {
                get
                {
                    return _agent;
                }
                set
                {
                    _agent = value;
                }
            }

            public string AgentCode
            {
                get
                {
                    return _agentcode;
                }
                set
                {
                    _agentcode = value;
                }
            }

            public byte Agent648Held
            {
                get
                {
                    return _agent648held;
                }
            set
                {
                    _agent648held = value;
                }
            }

            public string AgentAddress
            {
                get
                {
                    return _agentaddress;
                }
                set
                {
                    _agentaddress = value;
                }
            }

            public string NamedAgent
            {
                get
                {
                    return _namedagent;
                }
                set
                {
                    _namedagent = value;
                }
            }

            public string OtherContact
            {
                get
                {
                    return _othercontact;
                }
                set
                {
                    _othercontact = value;
                }
            }

            public string AgentTelNo
            {
                get
                {
                    return _agenttelno;
                }
                set
                {
                    _agenttelno = value;
                }
            }

            public byte Changed
            {
                get
                {
                    return _changed;
                }
                set
                {
                    _changed = value;
                }
            }

            // Behaviours Data

            public Int32 RPD_ID
            {
                get
                {
                    return _rpdid;
                }
                set
                {
                    _rpdid = value;
                }
            }

            public string UpdatedDate
            {
                get
                {
                    return _updateddate;
                }
                set
                {
                    _updateddate = value;
                }
            }

            public Int32 UpdatedBy
            {
                get
                {
                    return _updatedby;
                }
                set
                {
                    _updatedby = value;
                }
            }

            public int CalendarYear
            {
                get
                {
                    return _calendaryear;
                }
                set
                {
                    _calendaryear = value;
                }
            }

            public int LPOpen
            {
                get
                {
                    return _lpopen;
                }
                set
                {
                    _lpopen = value;
                }
            }

            public int LPClosed
            {
                get
                {
                    return _lpclosed;
                }
                set
                {
                    _lpclosed = value;
                }
            }

            public double HPPenalty
            {
                get
                {
                    return _hppenalty;
                }
                set
                {
                    _hppenalty = value;
                }
            }

            public byte PSCurrent
            {
                get
                {
                    return _pscurrent;
                }
                set
                {
                    _pscurrent = value;
                }
            }

            public byte PSPrevious
            {
                get
                {
                    return _psprevious;
                }
                set
                {
                    _psprevious = value;
                }
            }

            public byte PSFailures
            {
                get
                {
                    return _psfailures;
                }
                set
                {
                    _psfailures = value;
                }
            }

            public int QSScore
            {
                get
                {
                    return _qsscore;
                }
                set
                {
                    _qsscore = value;
                }
            }

            public int RPTPRScore
            {
                get
                {
                    return _rptprscore;
                }
                set
                {
                    _rptprscore = value;
                }
            }

            public int RPTAVScore
            {
                get
                {
                    return _rptavscore;
                }
                set
                {
                    _rptavscore = value;
                }
            }

            public int CGScore
            {
                get
                {
                    return _cgscore;
                }
                set
                {
                    _cgscore = value;
                }
            }

            public int ResScore
            {
                get
                {
                    return _resscore;
                }
                set
                {
                    _resscore = value;
                }
            }

            public int CRMScore
            {
                get
                {
                    return _crmscore;
                }
                set
                {
                    _crmscore = value;
                }
            }


            public int PriorityScore
            {
                get
                {
                    return _priorityscore;
                }
                set
                {
                    _priorityscore = value;
                }
            }

            public float Percentile
            {
                get
                {
                    return _percentile;
                }
                set
                {
                    _percentile = value;
                }
            }

            public string SupecededDate
            {
                get
                {
                    return _supercededdate;
                }
                set
                {
                    _supercededdate = value;
                }
            }

            public byte RiskingComplete
            {
                get
                {
                    return _riskingcomplete;
                }
                set
                {
                    _riskingcomplete = value;
                }
            }

            public string SegmentRecorded
            {
                get
                {
                    return _segmentrecorded;
                }
                set
                {
                    _segmentrecorded = value;
                }
            }

            public string RSDLU
            {
                get
                {
                    return _rsdlu;
                }
                set
                {
                    _rsdlu = value;
                }
            }

            public string CRMExplanation
            {
                get
                {
                    return _crmexplanation;
                }
                set
                {
                    _crmexplanation = value;
                }
            }

            // CRMM
            public int Risks_Open
            {
                get
                {
                    return _risksopen;
                }
                set
                {
                    _risksopen = value;
                }
            }

            public int Settled_Risks
            {
                get
                {
                    return _settledrisks;
                }
                set
                {
                    _settledrisks = value;
                }
            }

            public float Highest_Settlement
            {
                get
                {
                    return _highestsettlement;
                }
                set
                {
                    _highestsettlement = value;
                }
            }

            public int Highest_Percentage
            {
                get
                {
                    return _highestpercentage;
                }
                set
                {
                    _highestpercentage = value;
                }
            }

            public string CRMM_Date_Added
            {
                get
                {
                    return _crmmdateadded;
                }
                set
                {
                    _crmmdateadded = value;
                }
            }

            public DataTable Historical_Data
            {
                get
                {
                    return _historicaldata;
                }
                set
                {
                    _historicaldata = value;
                }
            }

            public DataTable Grid_Data
            {
                get
                {
                    return _griddata;
                }
                set
                {
                    _griddata = value;
                }
            }
        
            public DataTable Chart_Data
            {
                get
                {
                    return _chartdata;
                }
                set
                {
                    _chartdata = value;
                }
            }

            #endregion

            public bool GetRPDData(double dblUTR, int iYear, double dPercentile, string sPop)
            {
                if (GetCustomerData(dblUTR) == false)
                {
                    return false;
                }

                if (GetAgentData(dblUTR) == false)
                {
                    return false;
                }


                if (GetRPDScoresData(dblUTR, iYear, sPop) == false)
                {
                    return false;
                }

                if (GetRPDHistoricalData(dblUTR, iYear, dPercentile, sPop) == false)
                {
                    return false;
                }

                return true;
            }

            public bool GetCustomerData(double dblUTR)
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
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
                        string sDate;
                        dr.Read();
                        CU_ID = Convert.ToInt32(dr["CU_ID"]);
                        Strand = dr["Strand"].ToString();
                        Segment = dr["Segment"].ToString();
                        Surname = dr["Surname"].ToString();
                        Firstname = dr["FirstName"].ToString();
                        UTR = double.Parse(dr["UTR"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        try { sDate = Convert.ToDateTime(dr["DOB"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        DOB = sDate;
                        bool blnDeceased = (dr["Deceased"] is DBNull) ? false : Convert.ToBoolean(dr["Deceased"]);
                        Deceased = Convert.ToByte(blnDeceased);
                        try { sDate = Convert.ToDateTime(dr["DOD"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        DOD = sDate;
                        try { sDate = Convert.ToDateTime(dr["Deselected"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        Deselected = sDate;
                        Marital = dr["Marital"].ToString();
                        Gender = dr["Gender"].ToString();
                        MainAdd = dr["MainAdd"].ToString();
                        MainPC = dr["MainPC"].ToString();
                        SecAdd = dr["SecAdd"].ToString();
                        Residence = dr["Residence"].ToString();
                        Domicile = dr["Domicile"].ToString();
                        Office = dr["Office"].ToString();
                        Team = dr["Team"].ToString();
                        AllocatedTo = dr["AllocatedTo"].ToString();
                        WealthLevel = dr["WealthLevel"].ToString();
                        Pathway = dr["Pathway"].ToString();
                        Source = dr["Source"].ToString();
                        Sector = dr["Sector"].ToString();
                        LongTerm = dr["LongTerm"].ToString();
                        LifeEvents = dr["LifeEvents"].ToString();
                        Narrative = dr["Narrative"].ToString();
                        HNWUPID = (string.IsNullOrEmpty(dr["HNWUPID"].ToString()) == true) ? 0 : Convert.ToInt32(dr["HNWUPID"]);
                        Pop = dr["Pop"].ToString();
                        CRM_Name = dr["CRM_Name"].ToString();
                        try { sDate = Convert.ToDateTime(dr["CRM_Appointed"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        CRM_Appointed = sDate;
                        try { sDate = Convert.ToDateTime(dr["CDLU"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        CDLU = sDate;
                    }
                        #endregion
                    else if (dr.HasRows == false)
                    {
                        //MessageBox.Show("Customer record not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    con.Close();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                return true;
            }

            public bool GetAgentData(double dblUTR)
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryGetAgentData", con);  // tblAgent_Details
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
                        dr.Read();
                        AgentRecordID = (dr["Agent_Record_ID"] is DBNull) ? 0 : Convert.ToInt32(dr["Agent_Record_ID"]);
                        Agent = dr["Agent"].ToString();
                        AgentCode = dr["AgentCode"].ToString();
                        bool blnAgent648Held = (dr["648_held"] is DBNull) ? false : Convert.ToBoolean(dr["648_held"]);
                        Agent648Held = Convert.ToByte(blnAgent648Held);
                        AgentAddress = dr["Agent_Address"].ToString();
                        NamedAgent = dr["Named_Agent"].ToString();
                        OtherContact = dr["Other_Contact"].ToString();
                        AgentTelNo = dr["Agent_Tel_No"].ToString();
                        bool blnChanged = (dr["Changed"] is DBNull) ? false : Convert.ToBoolean(dr["Changed"]);
                        Changed = Convert.ToByte(blnChanged);
                    }
                    #endregion
                    else if (dr.HasRows == false)
                    {
                        //MessageBox.Show("Agent record not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    con.Close();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                return true;
            }



            public bool GetRPDScoresData(double dblUTR, int iYear, string sPop)
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryGetRPDDataScore", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                    prm01.Value = dblUTR;
                    SqlParameter prm02 = cmd.Parameters.Add("@nYear", SqlDbType.Int);
                    prm02.Value = iYear;
                    SqlParameter prm03 = cmd.Parameters.Add("@nPop", SqlDbType.Text);
                    prm03.Value = sPop;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader dr = cmd.ExecuteReader();
                    #region Recordset
                    if (dr.HasRows)
                    {
                        dr.Read();
                        string sDate;
                        RPD_ID = Convert.ToInt32(dr["RPD_ID"]);
                        try { sDate = Convert.ToDateTime(dr["UpdatedDate"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        UpdatedDate = sDate;
                        UpdatedBy = dr["UpdatedBy"] == DBNull.Value ? 0 : Convert.ToInt32(dr["UpdatedBy"]);
                        CalendarYear = dr["CalendarYear"] == DBNull.Value ? 0 : Convert.ToInt16(dr["CalendarYear"]);
                        LPOpen = dr["LPOpen"] == DBNull.Value ? 0 : Convert.ToInt16(dr["LPOpen"]);
                        LPClosed = dr["LPClosed"] == DBNull.Value ? 0 : Convert.ToInt16(dr["LPClosed"]);
                        HPPenalty = dr["HPPenalty"] == DBNull.Value ? 0 : Convert.ToDouble(dr["HPPenalty"]);
                        byte MaxThreeWay = 2; /*Yes/No/Unknown*/
                        PSCurrent = (dr["PSCurrent"] is DBNull) ? MaxThreeWay : Convert.ToByte(dr["PSCurrent"]);
                        PSCurrent = (PSCurrent > MaxThreeWay) ? MaxThreeWay : PSCurrent;
                        PSPrevious = (dr["PSPrevious"] is DBNull) ? MaxThreeWay : Convert.ToByte(dr["PSPrevious"]);
                        PSPrevious = (PSPrevious > MaxThreeWay) ? MaxThreeWay : PSPrevious;
                        PSFailures = (dr["PSFailures"] is DBNull) ? MaxThreeWay : Convert.ToByte(dr["PSFailures"]);
                        PSFailures = (PSFailures > MaxThreeWay) ? MaxThreeWay : PSFailures;
                        QSScore = dr["QSScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["QSScore"]);
                        RPTPRScore = dr["RPTPRScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["RPTPRScore"]);
                        RPTAVScore = dr["RPTAVScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["RPTAVScore"]);
                        CGScore = dr["CGScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["CGScore"]);
                        ResScore = dr["ResScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["ResScore"]);
                        CRMScore = dr["CRMScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["CRMScore"]);
                        PriorityScore = dr["PriorityScore"] == DBNull.Value ? 0 : Convert.ToInt16(dr["PriorityScore"]);
                        Percentile = dr["Percentile"] == DBNull.Value ? 0 : Convert.ToInt32(dr["Percentile"]);
                        try { sDate = Convert.ToDateTime(dr["SupecededDate"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        SupecededDate = sDate;
                        RiskingComplete = Convert.ToByte(dr["RiskingComplete"]);
                        //SegmentRecorded = dr["Segment_Recorded"].ToString();
                        //RSDLU = dr["RSDLU"].ToString();
                        //CRMExplanation = dr["CRM_Explanation"].ToString();
                    }
                    #endregion
                    else if (dr.HasRows == false)
                    {
                        //MessageBox.Show("Behaviours scores data not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    con.Close();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                return true;
            }

            public bool GetRPDHistoricalData(double dblUTR, int iYear, double dPercentile, string sPop)
            {
                DataSet ds = new DataSet();
                
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryGetRPDHistoricalData", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                    prm01.Value = dblUTR;
                    SqlParameter prm02 = cmd.Parameters.Add("@nYear", SqlDbType.Int);
                    prm02.Value = iYear;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(ds);

                    //// second table
                    //DataSet ds = new DataSet();
                    //da.Fill(ds);
                    //intTotalRows = Convert.ToInt32(ds.Tables[1].Rows[0]["NumberOfRows"]);

                    Historical_Data = ds.Tables[0];
                    Globals.dtGrid = ds.Tables[1];
                    //Grid_Data = ds.Tables[1];
                    Globals.dtGraph = ds.Tables[2];
                    //Chart_Data = ds.Tables[2];

                    //previously 01/06/2020
                    //con.Open();
                    //SqlDataReader dr = cmd.ExecuteReader();
                    //#region Recordset
                    //if (dr.HasRows)
                    //{
                    //    dr.Read();
                    //    QSScore = Convert.ToInt16(dr["QSScore"]);
                    //    RPTPRScore = Convert.ToInt16(dr["RPTPRScore"]);
                    //    RPTAVScore = Convert.ToInt16(dr["RPTAVScore"]);
                    //    CGScore = Convert.ToInt16(dr["CGScore"]);
                    //    ResScore = Convert.ToInt16(dr["ResScore"]);
                    //    CRMScore = Convert.ToInt16(dr["CRMScore"]);
                    //    PriorityScore = Convert.ToInt16(dr["PriorityScore"]);
                    //    Percentile = Convert.ToInt32(dr["Percentile"]);
                    //    SegmentRecorded = dr["Segment_Recorded"].ToString();
                    //    RSDLU = dr["RSDLU"].ToString();
                    //    CRMExplanation = dr["CRM_Explanation"].ToString();
                    //}
                    //#endregion
                    //else if (dr.HasRows == false)
                    //{
                    //    //MessageBox.Show("Behaviours scores data not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //}
                    //con.Close();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                try
                {
                    SqlCommand cmd = new SqlCommand("qryGetCRMMEnquiryDataScore", con);
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
                        dr.Read();
                        string sDate;
                        Risks_Open = dr["Risks_Open"] == DBNull.Value ? 0 :  Convert.ToInt16(dr["Risks_Open"]);
                        Settled_Risks = dr["Settled_Risks"] == DBNull.Value ? 0 : Convert.ToInt16(dr["Settled_Risks"]);
                        Highest_Settlement = dr["Highest_Settlement"] == DBNull.Value ? 0 : Convert.ToInt32(dr["Highest_Settlement"]);
                        try { sDate = Convert.ToDateTime(dr["CRMM_Date_Added"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        CRMM_Date_Added = sDate;
                    }
                    #endregion
                    else if (dr.HasRows == false)
                    {
                        //MessageBox.Show("Behaviours scores data not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    con.Close();
                }
                catch
                {
                    con.Close();
                    return false;
                }

                return true;
            }

            public void GetCRM()
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    string sPop = "";
                    string sCRM_Weighting = "";
                    int iItem = 0;
                    Globals.gs_CRM.Clear();

                    SqlCommand cmd = new SqlCommand("qryGetCRMWeighting", con);
                    cmd.Parameters.Clear();
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader dr = cmd.ExecuteReader();
                    #region Recordset
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            sPop = dr["pop"].ToString().ToUpper();
                            sCRM_Weighting = Convert.ToString(dr["CRM_Weighting"]);
                            Globals.gs_CRM.Add(new System.Collections.Generic.List<String>());
                            Globals.gs_CRM[iItem].Add(sPop);
                            Globals.gs_CRM[iItem].Add(sCRM_Weighting);
                            iItem++;
                        }
                    }
                    #endregion
                    else if (dr.HasRows == false)
                    {
                        Globals.gs_CRM.Clear(); /*no records, then clear the list*/
                    }
                    con.Close();
                }
                catch
                {
                    Globals.gs_CRM.Clear(); /*on error, clear the list*/
                    con.Close();
                }
                
            }

            public bool CheckCustomer(double dUTR)
            {
                bool bRtn = false;
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryCheckCustomer", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                    prm01.Value = dUTR;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        bRtn = true;
                    }
                    con.Close();
                }
                catch
                {
                    con.Close();
                }
                return bRtn;
            }

            public bool GetCRMMEnquiryDataScore(double dUTR)
            {
                
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryGetCRMMEnquiryDataScore", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nUTR", SqlDbType.Float);
                    prm01.Value = dUTR;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader dr = cmd.ExecuteReader();
                    #region Recordset
                    if (dr.HasRows)
                    {
                        dr.Read();
                        string sDate;
                        Risks_Open = dr["Risks_Open"] == DBNull.Value ? 0 : Convert.ToInt16(dr["Risks_Open"]);
                        Settled_Risks = dr["Settled_Risks"] == DBNull.Value ? 0 : Convert.ToInt16(dr["Settled_Risks"]);
                        Highest_Settlement = dr["Highest_Settlement"] == DBNull.Value ? 0 : Convert.ToInt32(dr["Highest_Settlement"]);
                        Highest_Percentage = dr["Highest_Percentage"] == DBNull.Value ? 0 : Convert.ToInt16(dr["Highest_Percentage"]);
                        try { sDate = Convert.ToDateTime(dr["CRMM_Date_Added"]).Date.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture); } catch { sDate = ""; }
                        CRMM_Date_Added = sDate;

                    }
                    #endregion
                    else if (dr.HasRows == false)
                    {
                        //MessageBox.Show("Behaviours scores data not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    con.Close();
                    return true;
                }
                catch
                {
                    con.Close();
                    return false;
                }

        }

            public bool UpdateCustomerData()
            {
            SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryUpdateCustomer", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nStrand", SqlDbType.NVarChar);
                    prm01.Value = Strand;
                    SqlParameter prm02 = cmd.Parameters.Add("@nSegment", SqlDbType.NVarChar);
                    prm02.Value = Segment;
                    SqlParameter prm03 = cmd.Parameters.Add("@nSurname", SqlDbType.NVarChar);
                    prm03.Value = Surname;
                    SqlParameter prm04 = cmd.Parameters.Add("@nFirstname", SqlDbType.NVarChar);
                    prm04.Value = Firstname;
                    SqlParameter prm05 = cmd.Parameters.Add("@nDOB", SqlDbType.DateTime);
                    if (DOB == "")
                    {
                        prm05.Value = DBNull.Value;
                    }
                    else
                    {
                        prm05.Value = DOB;
                    }
                    SqlParameter prm06 = cmd.Parameters.Add("@nDeceased", SqlDbType.Bit);
                    prm06.Value = Deceased;
                    SqlParameter prm07 = cmd.Parameters.Add("@nDOD", SqlDbType.DateTime);
                    if (DOD == "")
                    {
                        prm07.Value = DBNull.Value;
                    }
                    else
                    {
                        prm07.Value = DOD;
                    }
                    SqlParameter prm08 = cmd.Parameters.Add("@nDeselected", SqlDbType.DateTime);
                    if (Deselected == "")
                    {
                        prm08.Value = DBNull.Value;
                    }
                    else
                    {
                        prm08.Value = Deselected;
                    }
                    SqlParameter prm09 = cmd.Parameters.Add("@nMarital", SqlDbType.NVarChar);
                    prm09.Value = Marital;
                    SqlParameter prm10 = cmd.Parameters.Add("@nGender", SqlDbType.NVarChar);
                    prm10.Value = Gender;
                    SqlParameter prm11 = cmd.Parameters.Add("@nMainAddress", SqlDbType.NVarChar);
                    prm11.Value = MainAdd;
                    SqlParameter prm12 = cmd.Parameters.Add("@nMainPostCode", SqlDbType.NVarChar);
                    prm12.Value = MainPC;
                    SqlParameter prm13 = cmd.Parameters.Add("@nSecondAddress", SqlDbType.NVarChar);
                    prm13.Value = SecAdd;
                    SqlParameter prm14 = cmd.Parameters.Add("@nResidence", SqlDbType.NVarChar);
                    prm14.Value = Residence;
                    SqlParameter prm15 = cmd.Parameters.Add("@nDomicile", SqlDbType.NVarChar);
                    prm15.Value = Domicile;
                    SqlParameter prm16 = cmd.Parameters.Add("@nOffice", SqlDbType.NVarChar);
                    prm16.Value = Office;
                    SqlParameter prm17 = cmd.Parameters.Add("@nTeam", SqlDbType.NVarChar);
                    prm17.Value = Team;
                    SqlParameter prm18 = cmd.Parameters.Add("@nWealth", SqlDbType.NVarChar);
                    prm18.Value = WealthLevel;
                    SqlParameter prm19 = cmd.Parameters.Add("@nPathway", SqlDbType.NVarChar);
                    prm19.Value = Pathway;
                    SqlParameter prm20 = cmd.Parameters.Add("@nSource", SqlDbType.NVarChar);
                    prm20.Value = Source;
                    SqlParameter prm21 = cmd.Parameters.Add("@nSector", SqlDbType.NVarChar);
                    prm21.Value = Sector;
                    SqlParameter prm22 = cmd.Parameters.Add("@nLongTerm", SqlDbType.NVarChar);
                    prm22.Value = LongTerm;
                    SqlParameter prm23 = cmd.Parameters.Add("@nLifeEvents", SqlDbType.NVarChar);
                    prm23.Value = LifeEvents;
                    SqlParameter prm24 = cmd.Parameters.Add("@nNarrative", SqlDbType.NVarChar);
                    prm24.Value = Narrative;
                    SqlParameter prm25 = cmd.Parameters.Add("@nHNWU", SqlDbType.Int);
                    prm25.Value = HNWUPID;
                    SqlParameter prm26 = cmd.Parameters.Add("@oUTR", SqlDbType.Float);
                    prm26.Value = UTR;
                    SqlParameter prm27 = cmd.Parameters.Add("@nPop", SqlDbType.NVarChar);
                    prm27.Value = Pop;
                    SqlParameter prm28 = cmd.Parameters.Add("@nCRMName", SqlDbType.NVarChar);
                    prm28.Value = CRM_Name;
                    SqlParameter prm29 = cmd.Parameters.Add("@nCRMDA", SqlDbType.DateTime);
                    if (CRM_Appointed == "")
                    {
                        prm29.Value = DBNull.Value;
                    }
                    else
                    {
                        prm29.Value = CRM_Appointed;
                    }
                    SqlParameter prm30 = cmd.Parameters.Add("@nHNWUPID", SqlDbType.DateTime);
                    if (AllocatedTo == "")
                    {
                        prm30.Value = DBNull.Value;
                    }
                    else
                    {
                        prm30.Value = AllocatedTo;
                    }
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                return true;
            }

            public bool ResetAgent()
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryResetAgent", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@fUTR", SqlDbType.Float);
                    prm01.Value = UTR;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    con.Close();
                    return false;
                }
                return true;
            }

            public bool UpdateAgentData(double fUTR)
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryUpdateAgent", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@fUTR", SqlDbType.Float);
                    prm01.Value = fUTR;
                    SqlParameter prm02 = cmd.Parameters.Add("@nAgent", SqlDbType.NVarChar);
                    prm02.Value = Agent;
                    SqlParameter prm03 = cmd.Parameters.Add("@nAgentCode", SqlDbType.NVarChar);
                    prm03.Value = AgentCode;
                    SqlParameter prm04 = cmd.Parameters.Add("@b648held", SqlDbType.Bit);
                    prm04.Value = Agent648Held;
                    SqlParameter prm05 = cmd.Parameters.Add("@nAgentAddress", SqlDbType.NVarChar);
                    prm05.Value = AgentAddress;
                    SqlParameter prm06 = cmd.Parameters.Add("@nNamedAgent", SqlDbType.NVarChar);
                    prm06.Value = NamedAgent;
                    SqlParameter prm07 = cmd.Parameters.Add("@nOtherContact", SqlDbType.NVarChar);
                    prm07.Value = OtherContact;
                    SqlParameter prm08 = cmd.Parameters.Add("@nAgentTelNo", SqlDbType.NVarChar);
                    prm08.Value = AgentTelNo;
                    SqlParameter prm09 = cmd.Parameters.Add("@bChanged", SqlDbType.Bit);
                    prm09.Value = Changed;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.Write("SQL error: " + ex.Message);
                    con.Close();
                    return false;
                }
                return true;
            }

            public bool UpdateRiskData()
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                try
                {
                    SqlCommand cmd = new SqlCommand("qryUpdateRiskScores", con);
                    cmd.Parameters.Clear();
                    SqlParameter prm01 = cmd.Parameters.Add("@nLPOpen", SqlDbType.Int);
                    prm01.Value = LPOpen;
                    SqlParameter prm02 = cmd.Parameters.Add("@nLPClosed", SqlDbType.Int);
                    prm02.Value = LPClosed;
                    SqlParameter prm03 = cmd.Parameters.Add("@nHPPenalty", SqlDbType.Float);
                    prm03.Value = HPPenalty;
                    SqlParameter prm04 = cmd.Parameters.Add("@nPSCurrent", SqlDbType.Int);
                    prm04.Value = PSCurrent;
                    SqlParameter prm05 = cmd.Parameters.Add("@nPSPrevious", SqlDbType.Int);
                    prm05.Value = PSPrevious;
                    SqlParameter prm06 = cmd.Parameters.Add("@nPSFailures", SqlDbType.Int);
                    prm06.Value = PSFailures;
                    SqlParameter prm07 = cmd.Parameters.Add("@nQSScore", SqlDbType.Int);
                    prm07.Value = QSScore;
                    SqlParameter prm08 = cmd.Parameters.Add("@nRPTPRScore", SqlDbType.Int);
                    prm08.Value = RPTPRScore;
                    SqlParameter prm09 = cmd.Parameters.Add("@nRPTAVScore", SqlDbType.Int);
                    prm09.Value = RPTAVScore;
                    SqlParameter prm10 = cmd.Parameters.Add("@nCGScore", SqlDbType.Int);
                    prm10.Value = CGScore;
                    SqlParameter prm11 = cmd.Parameters.Add("@nResScore", SqlDbType.Int);
                    prm11.Value = ResScore;
                    SqlParameter prm12 = cmd.Parameters.Add("@nCRMScore", SqlDbType.Int);
                    prm12.Value = CRMScore;
                    SqlParameter prm13 = cmd.Parameters.Add("@nPriorityScore", SqlDbType.Int);
                    prm13.Value = PriorityScore;
                    SqlParameter prm14 = cmd.Parameters.Add("@nPercentile", SqlDbType.Float);
                    prm14.Value = Percentile;
                    SqlParameter prm15 = cmd.Parameters.Add("@nRiskingComplete", SqlDbType.Bit);
                    prm15.Value = 0; // ###########################################
                    SqlParameter prm16 = cmd.Parameters.Add("@nUpdatedBy", SqlDbType.Int);
                    prm16.Value = Global.PID;
                    SqlParameter prm17 = cmd.Parameters.Add("@nUpdatedDate", SqlDbType.DateTime);
                    if (Deselected == "")
                    {
                        prm17.Value = DBNull.Value;
                    }
                    else
                    {
                        prm17.Value = DateTime.Now.ToShortDateString();
                    }
                    SqlParameter prm18 = cmd.Parameters.Add("@nCalendarYear", SqlDbType.Int);
                    prm18.Value = CalendarYear;
                    SqlParameter prm19 = cmd.Parameters.Add("@fUTR", SqlDbType.Float);
                    prm19.Value = UTR;
                    SqlParameter prm20 = cmd.Parameters.Add("@nPop", SqlDbType.NVarChar);
                    prm20.Value = Pop;
                    SqlParameter prm21 = cmd.Parameters.Add("@nSegment", SqlDbType.NVarChar);
                    prm21.Value = Segment;
                    SqlParameter prm22 = cmd.Parameters.Add("@nCRMExplanation", SqlDbType.NVarChar);
                    prm22.Value = CRMExplanation;
                    cmd.CommandTimeout = Global.TimeOut;
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.Write("SQL error: " + ex.Message);
                    con.Close();
                    return false;
                }
                return true;
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



                //rptDet.lvwEmail.DataContext = ds.Tables[0].DefaultView;
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
