using FNet.Settings.Models;
using Nskd;
using System;
using System.Web.Mvc;

namespace FNet.Settings.Controllers
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
    }
}