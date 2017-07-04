using SAP.Middleware.Connector;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace connDB
{
    public class sapConnClass
    {
        public RfcConfigParameters setParaToConn(string sapClient)
        {
            RfcConfigParameters rfcPara = new RfcConfigParameters();
            switch (sapClient)
            {
                case "800":
                    rfcPara.Add(RfcConfigParameters.Name, "PRD");
                    rfcPara.Add(RfcConfigParameters.AppServerHost, "192.168.0.16");
                    rfcPara.Add(RfcConfigParameters.Client, "800");
                    break;
                case "620":
                    rfcPara.Add(RfcConfigParameters.Name, "DEV");
                    rfcPara.Add(RfcConfigParameters.AppServerHost, "192.168.0.15");
                    rfcPara.Add(RfcConfigParameters.Client, "620");
                    break;
                case "300":
                    rfcPara.Add(RfcConfigParameters.Name, "DEV");
                    rfcPara.Add(RfcConfigParameters.AppServerHost, "192.168.0.15");
                    rfcPara.Add(RfcConfigParameters.Client, "300");
                    break;
                default:
                    rfcPara.Add(RfcConfigParameters.Name, "DEV");
                    rfcPara.Add(RfcConfigParameters.AppServerHost, "192.168.0.15");
                    rfcPara.Add(RfcConfigParameters.Client, "620");
                    break;
            }
            rfcPara.Add(RfcConfigParameters.User, "DDIC");
            rfcPara.Add(RfcConfigParameters.Password, "Ubn3dx");
            rfcPara.Add(RfcConfigParameters.SystemNumber, "00");
            rfcPara.Add(RfcConfigParameters.Language, "ZF");

            return rfcPara;
        }

        public DataTable GetDataTableFromRFCTable(IRfcTable rfcTable)
        {
            DataTable dt = new DataTable();

            for (int i = 0; i <= rfcTable.ElementCount - 1; i++)
            {

                RfcElementMetadata metadata = rfcTable.GetElementMetadata(i);

                dt.Columns.Add(metadata.Name);

            }

            foreach (IRfcStructure rfcStructure in rfcTable)
            {

                DataRow dr = dt.NewRow();

                for (int j = 0; j <= rfcTable.ElementCount - 1; j++)
                {
                    RfcElementMetadata metadata = rfcTable.GetElementMetadata(j);

                    var rowData = rfcStructure.GetString(metadata.Name);

                    dr[metadata.Name] = rowData ;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        public DataTable GetDataTableFromRFCStructure(IRfcStructure rfcStructure)
        {            
            DataTable dt = new DataTable();

            for (int i = 0; i <= rfcStructure.ElementCount - 1; i++) dt.Columns.Add(rfcStructure.GetElementMetadata(i).Name);

            DataRow row = dt.NewRow();

            for (int j = 0; j <= rfcStructure.ElementCount - 1; j++) row[j] = rfcStructure.GetValue(j);

            dt.Rows.Add(row);

            return dt;          



        }

        public DataTable chgColName(DataTable orgDt, string[,] strArray, string lang)
        {
            DataTable tempDt = new DataTable();

            tempDt.Merge(orgDt);

            string twColName, enColName;
            int dtColCount = orgDt.Columns.Count;
            int arrayCount = strArray.GetLength(0);

            if (dtColCount==arrayCount)
            {
                for (int i=1;i<=dtColCount;i++)
                {
                    DataRow row = tempDt.NewRow();

                    if (i == 1)
                    {
                        for (int headerCount = 0; headerCount < dtColCount; headerCount++)
                        {
                            enColName = strArray[headerCount, 0];
                            twColName = strArray[headerCount, 1];
                            switch (lang)
                            {
                                case "tw":
                                    tempDt.Columns[enColName].ColumnName = twColName;
                                    break;
                                case "en":
                                    tempDt.Columns[twColName].ColumnName = enColName;
                                    break;
                            }                            
                        }
                    }

                }
            }
            return tempDt;
        }
    }


    class mssqlConnClass
    {
        public string sapDB { get; set; }
        public string sapInitDB { get; private set; }

        public string toSAPDB(string DBenv)
        {
            string sapDBConn;
            string source, catalog, auth;
            auth = "Uid=archer; Pwd=ko123vr4; ";
            switch (DBenv)
            {
                case "800":
                    source = "Data Source=192.168.0.16;";
                    sapDB = "prd";
                    catalog = "Initial Catalog=PRD;";
                    break;
                case "620":
                    source = "Data Source=192.168.0.15;";
                    sapDB = "dev";
                    catalog = "Initial Catalog=DEV;";
                    break;
                case "300":
                    source = "Data Source=192.168.0.15;";
                    sapDB = "dev";
                    catalog = "Initial Catalog=DEV;";
                    break;
                default:
                    source = "Data Source=192.168.0.15;";
                    sapDB = "dev";
                    catalog = "Initial Catalog=DEV;";
                    break;
            }
            sapDBConn = source + catalog + auth;
            return sapDBConn;
        }

        public string detectDBName(string connClient)
        {
            switch (connClient)
            {
                case "800":
                    sapInitDB = "prd";
                    break;
                case "620":
                    sapInitDB = "dev";
                    break;
                case "300":
                    sapInitDB = "dev";
                    break;
            }
            return sapInitDB;
        }


        public string toPackingDB(string DBenv)
        {
            string packingConn;
            DBenv = DBenv.ToUpper();
            string source, catalog, auth;
            source = "Data Source=SBSDB;";
            auth = "Uid=PACKING; Pwd=Admin12-1; ";
            switch (DBenv)
            {
                case "PRD":
                    catalog = "Initial Catalog=PACKING;";
                    break;
                case "DEV":
                    catalog = "Initial Catalog=PACKING_DEV;";
                    break;
                default:
                    catalog = "Initial Catalog=PACKING_DEV;";
                    break;
            }
            packingConn = source + catalog + auth;
            return packingConn;
        }

    }

    class chkFormStatusClass
    {
        private bool formStatus;

        public bool isFormActive(string formName)
        {
            string connString = "Data Source=192.168.0.9;Initial Catalog=rptsvrDB ;User ID=rptsvrDB ;Password=adminrp";
            string queryString = "select activeStatus from activeForm where formName=@formName";
            
            SqlConnection scToSbsdb = new SqlConnection(connString);
            SqlCommand cmdToSbsdb = new SqlCommand(queryString, scToSbsdb);
            cmdToSbsdb.Parameters.Add("@formName", SqlDbType.VarChar).Value = formName;

            try
            {
                scToSbsdb.Open();

                SqlDataReader sdr = cmdToSbsdb.ExecuteReader();
                formStatus = Convert.ToBoolean(sdr.Read());

                scToSbsdb.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(),"error");
            }
            return formStatus;
        }
    }
}
