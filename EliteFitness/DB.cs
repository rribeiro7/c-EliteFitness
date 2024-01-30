using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace EliteFitness
{
    class DB
    {
        public string ConnectionString { get; private set; }

        public int Ginasio { get; private set; }//Rm -1  Alv -2 

        public DB()
        {
            //ConnectionString = ConfigurationManager.ConnectionStrings["DBRM"].ConnectionString;
            //Ginasio = 1;
            ConnectionString = ConfigurationManager.ConnectionStrings["DBALV"].ConnectionString;
            Ginasio = 2;

        }

        public DataTable Select(string query)
        {
            SqlConnection sqlCon = new SqlConnection(ConnectionString);
            SqlCommand sqlCom = new SqlCommand(query, sqlCon);
            try
            {
                sqlCon.Open();
                DataTable table = new DataTable("Table");
                table.Load(sqlCom.ExecuteReader());
                return table;
            }
            finally
            {
                sqlCon.Close();
            }
        }
    }
}
