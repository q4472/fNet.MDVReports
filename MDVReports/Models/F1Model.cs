using System;
using System.Data;
using System.Data.SqlClient;

namespace FNet.Settings.Models
{
    public class F1Model
    {
        public DataTable Data;
        public String Status;
        public F1Model()
        {
            Status = "FNet.Settings.Models.F1Model(): ";
            String cnString = "Data Source=192.168.135.14;Initial Catalog=Grls;Integrated Security=True";
            SqlConnection cn = new SqlConnection(cnString);
            SqlCommand cmd = new SqlCommand()
            {
                Connection = cn,
                CommandType = CommandType.StoredProcedure,
                CommandText = "[Ссылки на РУ сравнение дат]"
            };
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                Data = dt;
            }
            catch (Exception e) { Status += $"Error: {e.Message}"; }
            Status += "OK.";

            Data = new DataTable();
            Data.Columns.Add("Ссылка", typeof(String));
            Data.Columns.Add("Номер", typeof(String));
            Data.Columns.Add("Дата FILE", typeof(String));
            Data.Columns.Add("Дата GRLS", typeof(String));
            Data.Columns.Add("needToRefresh", typeof(Boolean));

            foreach(DataRow dr in dt.Rows)
            {
                String dr0 = dr[0] as String;
                if (dr0 == null) dr0 = String.Empty;

                String dr1 = dr[1] as String;
                if (dr1 == null) dr1 = String.Empty;

                Object dateFile = dr[2];
                String dateFileAsString = (dateFile == DBNull.Value) ? "" : ((DateTime)dateFile).ToString("dd.MM.yyyy");

                Object dateGrls = dr[3];
                String dateGrlsAsString = (dateGrls == DBNull.Value) ? "" : ((DateTime)dateGrls).ToString("dd.MM.yyyy");

                Boolean needToRefresh = false;
                if (dateFile != DBNull.Value && dateGrls != DBNull.Value)
                {
                    needToRefresh = ((DateTime)dateFile < (DateTime)dateGrls);
                }

                Data.Rows.Add(new Object[] { 
                    dr0,
                    dr1,
                    dateFileAsString,
                    dateGrlsAsString,
                    needToRefresh
                });
            }
        }
    }
}
