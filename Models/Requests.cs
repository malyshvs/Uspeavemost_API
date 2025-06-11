using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;
using Uspevaemost_API.Controllers;

namespace Uspevaemost_API.Models
{
    public class Requests
    {
        private readonly string connectionString;
        private SqlConnection sql;
        public Requests(IConfiguration configuration)
        {
          
            connectionString = configuration.GetConnectionString("DefaultConnection");
            sql = new SqlConnection(connectionString);
        }




        public static string checkToken(string token,string con)
        {
            using var conn = new SqlConnection(con);
            using var cmd = new SqlCommand("publicbase.dbo.GetUserByToken", conn)
            {
                CommandType = CommandType.StoredProcedure
            };

            cmd.Parameters.AddWithValue("@Token", token);

            var outputLogin = new SqlParameter("@Login", SqlDbType.NVarChar, 256)
            {
                Direction = ParameterDirection.Output
            };
            cmd.Parameters.Add(outputLogin);

            conn.Open();
            cmd.ExecuteNonQuery();

            return outputLogin.Value?.ToString();
        }
        // 1
        public List<string[]> getData2(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs,string name)
        {
      
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=1, 
                @years='{string.Join(',',year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";
            Logger.Log(query);
            try
            {
                sql.Open();
                Console.WriteLine(sql.ToString());  
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[36];
                    for (int i = 0; i < 36; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        // 2
        public List<string[]> getDataInv(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs,string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=2, 
                         @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();

                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[36];
                    for (int i = 0; i < 36; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        // 3
        public List<string[]> getInvDolgi(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=3, 
                       @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";


            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();

                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[21];
                    for (int i = 0; i < 21; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        // 4
        public List<string[]> getGroups(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=4, 
                         @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";
            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        // 5
        public List<string[]> getbyUchp(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {

            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=5, 
                @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";
            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }


        //6
        public List<string[]> getbyUO(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=6, 
                             @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //7
        public List<string[]> getbyUOCurs(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=7, 
                    @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //8
        public List<string[]> getbyFO(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=8, 
                    @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";


            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //9
        public List<string[]> getbyCURS(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=9, 
                     @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //10
        public List<string[]> getbyQuote(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=10, 
                          @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //11
        public List<string[]> getbyCountry(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=11, 
                           @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //12
        public List<string[]> getDisc(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {

            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=12, 
                             @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[25];
                    for (int i = 0; i < 25; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //13
        public List<string[]> getbyKaf(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=13, 
                     @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //14
        public List<string[]> getbyKafCurs(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=14, 
                @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //15
        public List<string[]> getbyOPOP(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=15, 
                           @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[11];
                    for (int i = 0; i < 11; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        //16
        public List<string[]> getbyNapr(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs, string name)
        {
            string query = $@"execute publicbase.dbo.uspevaemost 
                @Querytype=16, 
                       @years='{string.Join(',', year)}',
                @sem='{string.Join(',', sem)}',
                @uo='{string.Join(',', uo)}',
                @fo='{string.Join(',', fo)}',
                @curs='{string.Join(',', curs)}',
                @name='{string.Join(',', name)}'";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
    }
}
