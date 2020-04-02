using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.Mvc;

namespace FNet.Settings.Controllers
{
    public class F1Controller : Controller
    {
        public Object Index()
        {
            Object result = "FNet.MDVReports.Controllers.F1Controller.Index()";
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
                result = PartialView("~/Views/F1/Index.cshtml", dt);
            }
            catch (Exception e) { result += "<br>" + e.Message; }
            return result;
        }
    }
}