using FNet.MDVReports.Models;
using Nskd;
using System;
using System.IO;
using System.Web.Mvc;

namespace FNet.MDVReports.Controllers
{
    public class F0Controller : Controller
    {
        public Object Index()
        {
            Object v = "FNet.MDVReports.Controllers.F0Controller.Index()";
            StreamReader reader = new StreamReader(Request.InputStream, Request.ContentEncoding);
            String body = reader.ReadToEnd();
            if (!String.IsNullOrWhiteSpace(body))
            {
                if (body[0] == '{')
                {
                    RequestPackage rqp = RequestPackage.ParseRequest(Request.InputStream, Request.ContentEncoding);
                    F0Model m = new F0Model(rqp);
                    v = PartialView("~/Views/F0/Index.cshtml", m);
                }
                else if (body[0] == 's' && body.Length == 46)
                {
                    v = PartialView("~/Views/F0/WaitingPage.cshtml");
                }
            }
            return v;
        }
        public Object DownloadExcelFile()
        {
            Object v = "FNet.MDVReports.Controllers.F0Controller.DownloadExelFile()";
            RequestPackage rqp = new RequestPackage { SessionId = new Guid() };
            F0Model m = new F0Model(rqp);
            Byte[] buff = NskdExcel.ToExcel(m.Data);
            String fileName = "MDVReport1 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + ".xlsx";
            FileContentResult fcr = File(buff, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            v = fcr;
            return v;
        }
    }
}
