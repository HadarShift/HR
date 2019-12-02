using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;

namespace HR
{
    public class DBService
    {
        //string connectionString; 
        //private OdbcDataAdapter myAdapter;
        //private OdbcConnection conn;
        private OleDbDataAdapter myAdapter;
        private OleDbConnection conn;

        /// <constructor>
        /// Initialise Connection
        /// </constructor>
        public DBService()
        {
            myAdapter = new OleDbDataAdapter();
            conn = new OleDbConnection("Provider=IBMDA400;Data Source=172.16.1.158;User ID=as400;Password=as400");
            //OleDbConnection ConAS400 = new OleDbConnection();
            //ConAS400.ConnectionString = "Provider=IBMDA400;Data Source=172.16.1.158;User ID=as400;Password=as400";

        }
        /// <method>
        /// Open Database Connection if Closed or Broken
        /// </method>
        private OleDbConnection openConnection()
        {
            if (conn.State == ConnectionState.Closed || conn.State == ConnectionState.Broken)
            {
                conn.Open();
            }
            return conn;
        }
            

        /// <method>
        /// Select Query
        /// </method>
        //public DataTable executeSelectQuery(String _query, SqlParameter[] sqlParameter)
        //{
        //    SqlCommand myCommand = new SqlCommand();
        //    DataTable dataTable = new DataTable();
        //    dataTable = null;
        //    DataSet ds = new DataSet();
        //    try
        //    {
        //        myCommand.Connection = openConnection();
        //        myCommand.CommandText = _query;
        //        myCommand.Parameters.AddRange(sqlParameter);
        //        myCommand.ExecuteNonQuery();                
        //        myAdapter.SelectCommand = myCommand;
        //        myAdapter.Fill(ds);
        //        dataTable = ds.Tables[0];
        //    }
        //    catch (SqlException e)
        //    {
        //        MessageBox.Show("Error - Connection.executeSelectQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
        //        return null;
        //    }
        //    finally
        //    {

        //    }
        //    return dataTable;
        //}

        public DataTable executeSelectQueryNoParam(String _query)
        {
            DataTable dataTable = new DataTable();
            DataSet ds = new DataSet();
            using (OleDbConnection con = new OleDbConnection("Provider=IBMDA400;Data Source=172.16.1.158;User ID=as400;Password=as400"))
            {
                con.Open();
                using (OleDbCommand myCommand = new OleDbCommand(_query,con))
                {
                    try
                    {
                    //myCommand.Connection = openConnection();
                    //myCommand.CommandText = _query;
                    //myCommand.CommandTimeout = 1000;
                    //myCommand.Parameters.AddRange(sqlParameter);
                    myCommand.ExecuteNonQuery();
                    myAdapter.SelectCommand = myCommand;
                    myAdapter.Fill(ds);
                    dataTable = ds.Tables[0];
                    }
                    catch (SqlException e)
                    {
                        //MessageBox.Show("Error - Connection.executeSelectQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
                        return null;
                    }
                    finally
                    {
                        conn.Close();
                    }
                }
            }
            return dataTable;
        }

        public DataRow executeSelectQueryNoParamRow(String _query)
        {
            DataTable dataTable = new DataTable();
            DataSet ds = new DataSet();
            using (OleDbCommand myCommand = new OleDbCommand())
            {
                try
                {
                    myCommand.Connection = openConnection();
                    myCommand.CommandText = _query;
                    myCommand.CommandTimeout = 10000;
                    //myCommand.Parameters.AddRange(sqlParameter);
                    myCommand.ExecuteNonQuery();
                    myAdapter.SelectCommand = myCommand;
                    myAdapter.Fill(ds);
                    dataTable = ds.Tables[0];
                }
                catch (SqlException e)
                {
                    //MessageBox.Show("Error - Connection.executeSelectQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
                    return null;
                }
                finally
                {
                    conn.Close();
                }
            }
            return dataTable.Rows[0];
        }
        /// <method>
        /// Insert Query
        ///// </method>
        public bool executeInsertQuery(String _query)
        {
            OleDbCommand myCommand = new OleDbCommand();
            try
            {
                myCommand.Connection = openConnection();
                myCommand.CommandText = _query;
                //myCommand.Parameters.AddRange(sqlParameter);
                myAdapter.InsertCommand = myCommand;
                myCommand.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                //MessageBox.Show("Error - Connection.executeInsertQuery - Query: " + _query + " \nException: \n" + e.StackTrace.ToString());
                return false;
            }
            finally
            {
            }
            return true;
        }

        ///// <method>
        ///// Update Query
        ///// </method>
        public bool executeUpdateQuery(String _query)
        {
            OleDbCommand myCommand = new OleDbCommand();
            try
            {
                myCommand.Connection = openConnection();
                myCommand.CommandText = _query;
                //myCommand.Parameters.AddRange(sqlParameter);
                myAdapter.UpdateCommand = myCommand;
                myCommand.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                //MessageBox.Show("Error - Connection.executeUpdateQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
                return false;
            }
            finally
            {
            }
            return true;
        }
    }
}

