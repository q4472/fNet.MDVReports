using FNet.MDVReports.Models;
using Nskd;
using System;
using System.Web.Mvc;

namespace FNet.MDVReports.Controllers
{
    public class F1Controller : Controller
    {
        public Object Index()
        {
            Object result = "FNet.MDVReports.Controllers.F1Controller.Index()";
            RequestPackage rqp = RequestPackage.ParseRequest(Request.InputStream, Request.ContentEncoding);
            try
            {
                if (rqp != null)
                {
                    F1Model m = new F1Model(rqp.SessionId);
                    result = PartialView("~/Views/F1/Index.cshtml", m);
                }
            }
            catch (Exception e) { result += "<br>" + e.Message; }
            return result;
        }
        public Object DownloadExcelFile()
        {
            Object result = "FNet.MDVReports.Controllers.F1Controller.DownloadExelFile()";
            //RequestPackage rqp = RequestPackage.ParseRequest(Request.InputStream, Request.ContentEncoding);
            try
            {
                //if (rqp != null)
                {
                    F1Model m = new F1Model(new Guid());
                    Byte[] buff = NskdExcel1.ToExcel(m.Data);
                    String fileName = "MDVReport2 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + ".xlsx";
                    FileContentResult fcr = File(buff, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    result = fcr;
                }
            }
            catch (Exception e) { result += "<br>" + e.Message; }
            return result;
        }
    }
}