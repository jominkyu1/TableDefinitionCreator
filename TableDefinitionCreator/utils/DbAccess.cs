using System;
using System.Data;
using System.Data.SqlClient;

namespace TableDefinitionCreator.utils
{
    internal static class DbAccess
    {
        /// <summary>
        /// DB 의 연결 주소를 가져옵니다.
        /// </summary>
        private static string ConnectionString => ConfigManager.GetConnectionString();

        public static string InitialCatalog {
            get
            {
                try
                {
                    var builder = new SqlConnectionStringBuilder(ConnectionString);
                    return builder.InitialCatalog;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// DB와 연결 되었는지 확인합니다.
        /// </summary>
        /// <returns>연결시 True</returns>
        public static bool IsConn(out string errMsg)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    errMsg = string.Empty;
                    return conn.State == ConnectionState.Open;
                }
                catch (Exception ex)
                {
                    errMsg = ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// 쿼리를 이용하여 데이터 테이블 을 가져 옵니다.
        /// </summary>
        /// <param name="query">Sql Query</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataTable(string query)
        {
            var dt = new DataTable();
            using (var conn = new SqlConnection(ConnectionString))
            using (var comm = new SqlCommand(query, conn))
            using (var da = new SqlDataAdapter(comm))
            {
                try
                {
                    conn.Open();
                    da.Fill(dt);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }
            }
            return dt;
        }

        public static DataSet GetDataSet(string query)
        {
            var ds = new DataSet();
            using (var conn = new SqlConnection(ConnectionString))
            using (var comm = new SqlCommand(query, conn))
            using (var da = new SqlDataAdapter(comm))
            {
                try
                {
                    conn.Open();
                    da.Fill(ds);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }
            }
            return ds;
        }

        /// <summary>
        /// 쿼리를 이용하여 가져오는 데이터의 첫번째 내용을 가져옵니다.
        /// </summary>
        public static string ExecuteScalar(string query)
        {
            using (var conn = new SqlConnection(ConnectionString))
            using (var comm = new SqlCommand(query, conn))
            {
                try
                {
                    conn.Open();
                    object resultObj = comm.ExecuteScalar();
                    return resultObj?.ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// 쿼리 명령어를 DB 에 전송합니다.
        /// </summary>
        /// <param name="query"></param>
        /// <returns>영향을 받은 행의 수</returns>
        public static int ExecuteQuery(string query)
        {
            using (var conn = new SqlConnection(ConnectionString))
            using (var comm = new SqlCommand(query, conn))
            {
                try
                {
                    conn.Open();
                    return comm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }
            }
        }
    }
}
