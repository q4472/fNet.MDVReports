using Nskd;
using System;
using System.Data;

namespace FNet.MDVReports.Models
{
    public class F0Model
    {
        public RequestPackage Rqp;
        //public FilterData Filter;
        public FilteredData Data;
        //public DataTable Поставщики;
        //public DataTable СостоянияЗаказа;

        public F0Model(RequestPackage rqp)
        {
            Rqp = rqp;
            //Filter = new FilterData(this);
            Data = new FilteredData(this);
            //Поставщики = ПолучитьСписокПоставщиков(this);
            //СостоянияЗаказа = ПолучитьСписокСостоянийЗаказа(this);
        }
        /*
        public class FilterData
        {
            public String все = "False"; // фильтр по полю "обработано". По умолчанию обработанные строки не показываем.
            public String дата_min = "";
            public String дата_max = "";
            public String менеджер = "";

            public FilterData(F0Model m)
            {
                if (m.Rqp != null)
                {
                    все = (m.Rqp["все"] == null) ? "False" : ((Boolean)m.Rqp["все"]).ToString();
                    дата_min = (m.Rqp["дата_min"] == null) ? "" : (String)m.Rqp["дата_min"];
                    дата_max = (m.Rqp["дата_max"] == null) ? "" : (String)m.Rqp["дата_max"];
                    менеджер = (m.Rqp["менеджер"] == null) ? "" : (String)m.Rqp["менеджер"];
                }
            }
        }
        */
        public class FilteredData
        {
            private DataTable dt;

            public Int32 RowsCount { get => (dt == null) ? 0 : dt.Rows.Count; }
            public class ItemArray
            {
                public String наименование;
                public String количество_в_заявке;
                public String количество_в_спецификации;
                public String количество_в_накладных_1С;
                public String количество_к_отгрузке;
                public String менеджер;
                public String контракт;
                public String дата_окончания_котракта;
                public String аукцион;
                public String заказчик;
                public String this[String fieldName]
                {
                    get
                    {
                        String s = null;
                        var field = typeof(ItemArray).GetField(fieldName);
                        if (field != null)
                        {
                            s = (String)field.GetValue(this);
                        }
                        return s;
                    }
                }
            }

            public FilteredData(F0Model m)
            {
                if (m.Rqp != null && m.Rqp.SessionId != null)
                {
                    RequestPackage rqp = new RequestPackage();
                    rqp.SessionId = m.Rqp.SessionId;
                    rqp.Command = "Supply.dbo.отчёт_мдв_1";
                    rqp.Parameters = new RequestParameter[]
                    {
                        new RequestParameter() { Name = "session_id", Value = m.Rqp.SessionId },
                        //new RequestParameter() { Name = "все", Value = m.Filter.все }
                    };
                    //if (!String.IsNullOrWhiteSpace(m.Filter.дата_min)) rqp["дата_min"] = m.Filter.дата_min;
                    //if (!String.IsNullOrWhiteSpace(m.Filter.дата_max)) rqp["дата_max"] = m.Filter.дата_max;
                    //if (!String.IsNullOrWhiteSpace(m.Filter.менеджер)) rqp["менеджер"] = m.Filter.менеджер;
                    ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
                    if (rsp != null)
                    {
                        dt = rsp.GetFirstTable();
                    }
                }
            }
            public ItemArray this[Int32 index]
            {
                get
                {
                    ItemArray items = null;
                    if (dt != null && index >= 0 && index < dt.Rows.Count)
                    {
                        DataRow dr = dt.Rows[index];
                        items = new ItemArray
                        {
                            наименование = ConvertToString(dr["наименование"]),
                            количество_в_заявке = ConvertToString(dr["количество_в_заявке"]),
                            количество_в_спецификации = ConvertToString(dr["количество_в_спецификации"]),
                            количество_в_накладных_1С = ConvertToString(dr["количество_в_накладных_1С"]),
                            количество_к_отгрузке = ConvertToString(dr["количество_к_отгрузке"]),
                            менеджер = ConvertToString(dr["менеджер"]),
                            контракт = ConvertToString(dr["контракт"]),
                            дата_окончания_котракта = ConvertToString(dr["дата_окончания_котракта"]),
                            аукцион = ConvertToString(dr["аукцион"]),
                            заказчик = ConvertToString(dr["заказчик"])
                        };
                    }
                    return items;
                }
            }
            private String ConvertToString(Object v)
            {
                String s = String.Empty;
                if (v != null && v != DBNull.Value)
                {
                    String tfn = v.GetType().FullName;
                    switch (tfn)
                    {
                        case "System.Guid":
                            s = ((Guid)v).ToString();
                            break;
                        case "System.Int32":
                            s = ((Int32)v).ToString();
                            break;
                        case "System.Boolean":
                            s = ((Boolean)v).ToString();
                            break;
                        case "System.String":
                            s = (String)v;
                            break;
                        default:
                            s = tfn;
                            break;
                    }
                }
                return s;
            }
        }
        /*
        public DataTable ПолучитьСписокПоставщиков(F0Model m)
        {
            DataTable dt = null;
            RequestPackage rqp = new RequestPackage
            {
                SessionId = m.Rqp.SessionId,
                Command = "Supply.dbo.поставщики__получить",
                Parameters = new RequestParameter[]
                {
                    new RequestParameter() { Name = "session_id", Value = m.Rqp.SessionId }
                }
            };
            ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
            if (rsp != null)
            {
                dt = rsp.GetFirstTable();
            }
            return dt;
        }
        */
        /*
        public DataTable ПолучитьСписокСостоянийЗаказа(F0Model m)
        {
            DataTable dt = null;
            RequestPackage rqp = new RequestPackage
            {
                SessionId = m.Rqp.SessionId,
                Command = "Supply.dbo.состояния_заказа__получить",
                Parameters = new RequestParameter[]
                {
                    new RequestParameter() { Name = "session_id", Value = m.Rqp.SessionId }
                }
            };
            ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
            if (rsp != null)
            {
                dt = rsp.GetFirstTable();
            }
            return dt;
        }
        */
        /*
        public static DataTable GetOrderDetail(RequestPackage rqp)
        {
            DataTable dt = null;
            if (rqp != null && rqp.SessionId != null)
            {
                Guid.TryParse(rqp["order_uid"] as String, out Guid orderUid);
                rqp.Command = "Supply.dbo.заказ_у_поставщика__получить_атрибуты_заказа";
                rqp.Parameters = new RequestParameter[]
                {
                        new RequestParameter() { Name = "session_id", Value = rqp.SessionId },
                        new RequestParameter() { Name = "order_uid", Value = orderUid }
                };
                ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
                if (rsp != null)
                {
                    dt = rsp.GetFirstTable();
                }
            }
            return dt;
        }
        */
        /*
        public static DataTable GetPriceDetail(RequestPackage rqp)
        {
            DataTable dt = null;
            if (rqp != null && rqp.SessionId != null)
            {
                Guid.TryParse(rqp["uid"] as String, out Guid uid);
                rqp.Command = "Supply.dbo.заказ_у_поставщика__получить_атрибуты_цены";
                rqp.Parameters = new RequestParameter[]
                {
                        new RequestParameter() { Name = "session_id", Value = rqp.SessionId },
                        new RequestParameter() { Name = "заказы_у_поставщиков_таблица__uid", Value = uid }
                };
                ResponsePackage rsp = rqp.GetResponse("http://127.0.0.1:11012");
                if (rsp != null)
                {
                    dt = rsp.GetFirstTable();
                }
            }
            return dt;
        }
        */
        /*
        public static void SetSupplier(RequestPackage rqp)
        {
            Hashtable setSupplierValue = (Hashtable)rqp["SetSupplier"];
            Guid supplierUid = new Guid();
            String supplierName = null;
            StringBuilder uids = new StringBuilder();
            foreach (DictionaryEntry nvp in setSupplierValue)
            {
                if (nvp.Key as String == "supplier_uid")
                {
                    Guid.TryParse(nvp.Value as String, out supplierUid);
                }
                if (nvp.Key as String == "supplier_name")
                {
                    supplierName = nvp.Value as String;
                }
                if (nvp.Key as String == "uids")
                {
                    Object[] t = nvp.Value as Object[];
                    foreach (Object o in t)
                    {
                        uids.AppendFormat($"<a b=\"{o}\"/>");
                    }
                }
            }
            RequestPackage rqp1 = new RequestPackage()
            {
                SessionId = rqp.SessionId,
                Command = "Supply.dbo.заказы_у_поставщиков__установить_поставщика"
            };
            rqp1.Parameters = new RequestParameter[]
            {
                        new RequestParameter() { Name = "session_id", Value = rqp.SessionId },
                        new RequestParameter() { Name = "supplier_uid", Value = supplierUid },
                        new RequestParameter() { Name = "supplier_name", Value = supplierName },
                        new RequestParameter() { Name = "uids", Value = uids.ToString() }
            };
            ResponsePackage rsp = rqp1.GetResponse("http://127.0.0.1:11012");
        }
        */
    }
}
