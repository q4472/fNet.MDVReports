using FNet.Settings.Models;
using System;
using System.Web.Mvc;

namespace FNet.Settings.Controllers
{
    public class F1Controller : Controller
    {
        public Object Index()
        {
            Object result = "FNet.MDVReports.Controllers.F1Controller.Index()";
            try
            {
                F1Model m = new F1Model();
                result = PartialView("~/Views/F1/Index.cshtml", m.Data);
            }
            catch (Exception e) { result += "<br>" + e.Message; }
            return result;
        }
    }
}