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
            RequestPackage rqp = null;
            StreamReader reader = new StreamReader(Request.InputStream, Request.ContentEncoding);
            String body = reader.ReadToEnd();
            if (!String.IsNullOrWhiteSpace(body))
            {
                if (body[0] == '{')
                {
                    rqp = RequestPackage.ParseRequest(Request.InputStream, Request.ContentEncoding);
                }
                else if (body[0] == 's' && body.Length == 46)
                {
                    if (Guid.TryParse(body.Substring(10, 36), out Guid sessionId))
                    {
                        rqp = new RequestPackage
                        {
                            SessionId = sessionId
                        };
                    }
                }
            }
            if (rqp == null)
            {
                rqp = new RequestPackage
                {
                    SessionId = new Guid()
                };
            }
            F0Model m = new F0Model(rqp);
            v = PartialView("~/Views/F0/Index.cshtml", m);
            return v;
        }
    }
}
