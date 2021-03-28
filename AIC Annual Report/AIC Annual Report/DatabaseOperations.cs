using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
//namespace CN_SAM_Pricing_ApprovalRobot.Database
namespace DBHelper
{
    class clsDBParameter
    {
        public string _Name;
        public DbType _DbType;
        public object _Value;

        public clsDBParameter(string Name,DbType DbType,object value)
        {
            _Name = Name;
            _DbType = DbType;
            _Value = value;
        }
    }

    class OleDBOperations
    {
        public string _strConnString;

        private OleDbConnection _OledbConn;        


        //Help
        //Access Database data type and values
        //https://ss64.com/access/syntax-datatypes.html
        //https://support.microsoft.com/en-us/office/data-types-for-access-desktop-databases-df2b83ba-cef6-436d-b679-3418f622e482
            

        public OleDBOperations(string strConnString)
        {
            this._strConnString = strConnString;
            createConnection();
        }

        public OleDBOperations(string strProvider, string strDataSource)
        {
            OleDbConnectionStringBuilder connStrBuild = new OleDbConnectionStringBuilder();
            connStrBuild.Provider =  strProvider;
            connStrBuild.DataSource = strDataSource;
            this._strConnString = connStrBuild.ConnectionString;
            createConnection();            
        }
            
        private void createConnection()
        {
            this._OledbConn = new OleDbConnection(_strConnString);
        }

        public DataTable GetDataFromDB(string strSelectQuery)
        {
            OleDbCommand cmdSelect = new OleDbCommand(strSelectQuery, _OledbConn);
            OleDbDataAdapter adapterSelect = new OleDbDataAdapter(cmdSelect);
            DataTable tblData = new DataTable();

            //Insert into Database
            try
            {
                _OledbConn.Open();
                adapterSelect.Fill(tblData);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {//Close the connection
                if (_OledbConn.State != ConnectionState.Closed)
                    _OledbConn.Close();
            }
            return tblData;

            ////Insert New Record
            //intInsertedRecords = cmdSelect.ExecuteNonQuery();

        }

        /// <summary>
        /// THis function runs Insert query and returns the no of records inserted.
        /// </summary>
        /// <param name="strSelectQuery"></param>
        /// <returns></returns>
        public int insertDataToDB(string in_strInsertQuery, clsDBParameter[] in_parms)
        {
            int intRecordsInserted = 0;
            OleDbCommand cmdInsert = new OleDbCommand(in_strInsertQuery, _OledbConn);
                        
            //Add Parameters
            if (in_parms != null && in_parms.Length > 0)
            {
                OleDbParameter[] oledbParms = new OleDbParameter[in_parms.Length];
                for (int i = 0; i < in_parms.Length; i++)
                {
                    oledbParms[i] = new OleDbParameter();
                    oledbParms[i].ParameterName = in_parms[i]._Name;
                    oledbParms[i].DbType = in_parms[i]._DbType;
                    oledbParms[i].Value = in_parms[i]._Value;
                }
                cmdInsert.Parameters.AddRange(oledbParms); 
            }

            try
            {
                _OledbConn.Open();
                intRecordsInserted = cmdInsert.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {//Close the connection
                if (_OledbConn.State != ConnectionState.Closed)
                    _OledbConn.Close();
            }

            return intRecordsInserted;
        }

     	private string GetVariableType(DbType dbType)
		{
			switch (dbType) {
				case DbType.AnsiString:
				case DbType.AnsiStringFixedLength:
				case DbType.String:
				case DbType.StringFixedLength: return "string";

				case DbType.Binary:		return "byte[]";
				case DbType.Boolean:	return "bool";
				case DbType.Byte:		return "byte";

				case DbType.Currency:
				case DbType.Decimal:
				case DbType.VarNumeric:	return "decimal";

				case DbType.Date:
				case DbType.DateTime:	return "DateTime";

				case DbType.Double:		return "double";
				case DbType.Guid:		return "Guid";
				case DbType.Int16:		return "short";
				case DbType.Int32:		return "int";
				case DbType.Int64:		return "long";
				case DbType.Object:		return "object";
				case DbType.SByte:		return "sbyte";
				case DbType.Single:		return "float";
				case DbType.Time:		return "TimeSpan";
				case DbType.UInt16:		return "ushort";
				case DbType.UInt32:		return "uint";
				case DbType.UInt64:		return "ulong";

				default:				return "string";
			}
		}

        /// <summary>
        /// THis function runs update query and returns the no of records Updated.
        /// </summary>
        /// <param name="strSelectQuery"></param>
        /// <returns></returns>
        public int UpdateDataToDB(string in_strUpdateQuery, clsDBParameter[] in_parms)
        {
            int intRecordsUpdated = 0;
            OleDbCommand cmdUpdate = new OleDbCommand(in_strUpdateQuery, _OledbConn);

            //Add Parameters
            if (in_parms != null && in_parms.Length > 0)
            {
                OleDbParameter[] oledbParms = new OleDbParameter[in_parms.Length - 1];
                for (int i = 0; i < in_parms.Length; i++)
                {
                    oledbParms[i].ParameterName = in_parms[i]._Name;
                    oledbParms[i].DbType = in_parms[i]._DbType;
                    oledbParms[i].Value = in_parms[i]._Value;
                }
                cmdUpdate.Parameters.Add(oledbParms);
            }

            try
            {
                _OledbConn.Open();
                intRecordsUpdated = cmdUpdate.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {//Close the connection
                if (_OledbConn.State != ConnectionState.Closed)
                    _OledbConn.Close();
            }

            return intRecordsUpdated;
        }

    }

}
