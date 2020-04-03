using Nskd;
using System;
using System.Data;

namespace FNet.Settings.Models
{
    public class F1Model
    {
        public DataTable Data;
        public String Status;
        public Int32 TotalRowsCount;
        public Int32 NeedToRefreshRowsCount;
        public F1Model(Guid sessionId)
        {
            Status = "FNet.Settings.Models.F1Model(): ";

            DataTable dt = new DataTable();

            RequestPackage rqp = new RequestPackage()
            {
                SessionId = sessionId,
                Command = "[Grls].[dbo].[Ссылки на РУ сравнение дат]",
                Parameters = new RequestParameter[]
                {
                    new RequestParameter() { Name = "session_id", Value = sessionId }
                }
            };
            ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
            if (rsp != null)
            {
                dt = rsp.GetFirstTable();
            }

            Data = new DataTable();
            Data.Columns.Add("Ссылка", typeof(String));
            Data.Columns.Add("Номер", typeof(String));
            Data.Columns.Add("Дата FILE", typeof(String));
            Data.Columns.Add("Дата GRLS", typeof(String));
            Data.Columns.Add("needToRefresh", typeof(Boolean));

            TotalRowsCount = dt.Rows.Count;
            NeedToRefreshRowsCount = 0;
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
                if (dateFile != DBNull.Value && dateGrls != DBNull.Value && (DateTime)dateFile < (DateTime)dateGrls)
                {
                    needToRefresh = true;
                    NeedToRefreshRowsCount++;
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
