using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace CSharp.CabBook
{
    class QueryControl
    {
        //CREATE DB CONNECTION
        private OleDbConnection DBCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = CabBook.DB.accdb");

        //PREPARE COMMAND CALLS
        private OleDbCommand DbCmd = new OleDbCommand();

        //PREPARE DATA COLLECTOR
        public OleDbDataAdapter DBDA;
        public DataTable DBDT;

        //QUERY PARAMETERS
        public List<OleDbParameter> Params = new List<OleDbParameter>();

        //QUERY STATISTICS
        public int RecordCount;
        public string Exception;

        public void ExecQuery(string Query)
        {
            //INITIALISE QUERY STATS
            RecordCount = 0;
            Exception = "";
            try {
                //OPEN A CONNECTION
                DBCon.Open();

                //CREATE DB COMMAND
                DbCmd = new OleDbCommand(Query, DBCon);

                //LOAD PARAMETERS INTO COMMAND
                foreach (OleDbParameter par in Params)
                {
                    DbCmd.Parameters.Add(par);
                }

                //CLEAR PARAMETERS LIST
                Params.Clear();

                //EXECUTE COMMAND & FILL DATA
                DBDT = new DataTable();
                DBDA = new OleDbDataAdapter(DbCmd);
                RecordCount = DBDA.Fill(DBDT);

            } catch (Exception ex)
            { Exception = ex.Message; }
            //CLOSE CONNECTION
            if ( DBCon.State == ConnectionState.Open)
            { DBCon.Close(); }
        }
        public void AddParam(String Name, Object Value)
        {
            OleDbParameter NewParams = new OleDbParameter(Name, Value);
            Params.Add(NewParams);
        }

    }
}