using System.Data;
using System.Data.SqlClient;

namespace InfDiario.Data
{
    public class Informe
    {
        public DataTable getInforme()
        {

            Settings settings = new Settings();
            var cnStr = settings.getConnStr();

            DataTable dt = new DataTable();

            SqlConnection cn = new SqlConnection(cnStr);

            //open connection
            cn.Open();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandText = "uspInformeDiario";
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter da = new SqlDataAdapter(cmd.CommandText, cnStr);

            da.SelectCommand = cmd;
            da.Fill(dt);

            return dt;

        }

    }
}
