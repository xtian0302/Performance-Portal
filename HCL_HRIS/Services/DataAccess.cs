using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace HCL_HRIS.Services
{

    public class DataAccess
    { 
        SqlConnection connection = Models.Utilities.getConn();

        public static async Task<List<String[]>> query_Top5Async(SqlConnection connection)
        {
            string[] array = {"",""};
            List<string[]> arrays = new List<string[]>();
            SqlCommand command = new SqlCommand("Select top 1 * from top5 order by date desc", connection);
            //connection.Open();
            SqlDataReader reader = await command.ExecuteReaderAsync();
            if (reader.HasRows){
                while (reader.Read()){
                    array[0] = reader["top1_sap"].ToString();
                    array[1] = reader["top1_name"].ToString();
                    arrays.Add(array);
                    array[0] = reader["top2_sap"].ToString();
                    array[1] = reader["top2_name"].ToString();
                    arrays.Add(array);
                    array[0] = reader["top3_sap"].ToString();
                    array[1] = reader["top3_name"].ToString();
                    arrays.Add(array);
                    array[0] = reader["top4_sap"].ToString();
                    array[1] = reader["top4_name"].ToString();
                    arrays.Add(array);
                    array[0] = reader["top5_sap"].ToString();
                    array[1] = reader["top5_name"].ToString();
                    arrays.Add(array);
                }
            }else{ 
                for(int i = 0; i <  5; i++) {
                    arrays.Add(array); 
                }
            } 
            closeDataObjects(connection, command, reader);
            return arrays;
        }

        private static void closeDataObjects(SqlConnection connection, SqlCommand command, SqlDataReader reader)
        {
            command.Cancel();
            reader.Close();
            //connection.Close(); 
        }
    }
}