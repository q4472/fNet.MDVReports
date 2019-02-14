using Nskd;
using System;
using System.Data;
using System.IO;

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
                public String номер_в_списке;
                public String группа;
                public String номер_в_группе;
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
                            номер_в_списке = ConvertToString(dr["номер_в_списке"]),
                            группа = ConvertToString(dr["группа"]),
                            номер_в_группе = ConvertToString(dr["номер_в_группе"]),
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
                        case "System.Decimal":
                            s = ((Decimal)v).ToString("n0");
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
    public static class NskdExcel
    {
        private static class Md
        {
            public class TableColumn
            {
                public String ColumnName;
                public String Caption;
                public Type DataType;
                public String Width;
            }
            public static TableColumn[] Table1Columns = new TableColumn[]
            {
                    new TableColumn { ColumnName = "номер_в_списке", Caption = "Номер в списке", DataType = typeof(String), Width = "8" }, // 0 A
                    new TableColumn { ColumnName = "группа", Caption = "Код группы", DataType = typeof(String), Width = "14" }, // 1 B
                    new TableColumn { ColumnName = "номер_в_группе", Caption = "Номер в группе", DataType = typeof(String), Width = "8" }, // 2 C
                    new TableColumn { ColumnName = "наименование", Caption = "Наименование", DataType = typeof(String), Width = "64" }, // 3 D
                    new TableColumn { ColumnName = "количество_в_заявке", Caption = "Кол-во в заявке", DataType = typeof(String), Width = "0" }, // 4 E
                    new TableColumn { ColumnName = "количество_в_спецификации", Caption = "Кол-во в спец.", DataType = typeof(String), Width = "10" }, // 5 F
                    new TableColumn { ColumnName = "количество_в_накладных_1С", Caption = "Кол-во в накладных 1С", DataType = typeof(String), Width = "10" }, // 6 G
                    new TableColumn { ColumnName = "количество_к_отгрузке", Caption = "Кол-во к отгрузке", DataType = typeof(String), Width = "10" }, // 7 H
                    new TableColumn { ColumnName = "менеджер", Caption = "менеджер", DataType = typeof(String), Width = "16" }, // 8 I
                    new TableColumn { ColumnName = "контракт", Caption = "контракт", DataType = typeof(String), Width = "0" }, // 9 J
                    new TableColumn { ColumnName = "дата_окончания_котракта", Caption = "Дата окончания котракта", DataType = typeof(String), Width = "0" }, // 10 K
                    new TableColumn { ColumnName = "аукцион", Caption = "Номер аукциона", DataType = typeof(String), Width = "24" }, // 11 L
                    new TableColumn { ColumnName = "заказчик", Caption = "Заказчик", DataType = typeof(String), Width = "64" } // 12 M
            };
        }
        private static String GetColumnName(UInt32 index) // zero-based
        {
            const byte BASE = 'Z' - 'A' + 1;
            string name = String.Empty;
            do
            {
                name = System.Convert.ToChar('A' + index % BASE) + name;
                index = index / BASE;
            }
            while (index-- > 0);
            return name;
        }
        public static Byte[] ToExcel(F0Model.FilteredData data)
        {
            MemoryStream ms;
            UInt32 zoomScale = 100;
            String fontName = "Arial";
            Double fontSize = 10;
            using (Spreadsheet spreadsheet = new Spreadsheet(1)) // один лист
            {
                spreadsheet.SetSheetName(1, "Отчёт МДВ 1");
                spreadsheet.SetZoomScale(zoomScale);
                spreadsheet.SetFont(0, fontName, fontSize); // default font
                spreadsheet.SetFont(1, fontName, fontSize); // data font

                UInt32 dxfId = spreadsheet.AppendDifferentialFormat("font-weight: bold;");

                XlWorksheet[] wss = spreadsheet.Wss;

                GenerateColumns(wss[0], Md.Table1Columns);
                GenrateSheetData0(wss[0], data);
                //GenerateMergeCells(wss[0]);
                ///GenerateConditionalFormatting(wss[0], dxfId, t);
                ///GenerateBackgroundColor(wss[0], t);
                GeneratePageSetup(wss[0]);

                ms = spreadsheet.CreateDocument();
            }
            return ms.ToArray();
        }
        private static void GenerateColumns(XlWorksheet ws, Md.TableColumn[] cols)
        {
            uint cn = 1;
            for (int ci = 0; ci < cols.Length; ci++)
            {
                Md.TableColumn col = cols[ci];
                if (col.Width == null)
                {
                    ws.AppendColumn(cn, cn, true);
                }
                else
                {
                    Double.TryParse(col.Width, out Double width);
                    ws.AppendColumn(cn, cn, false, true, width);
                }
                cn++;
            }
        }
        public static void GenrateSheetData0(XlWorksheet ws, F0Model.FilteredData data)
        {
            UInt32 rowIndex = 0;
            UInt32 columnIndex = 0;

            // строка заголовка
            columnIndex = 0;
            foreach (Md.TableColumn column in Md.Table1Columns)
            {
                String cellValueText = column.Caption;
                ws.UpsertCell(rowIndex, columnIndex, 1, cellValueText); // CellValues.SharedString
                columnIndex++;
            }
            rowIndex++;

            // строки данных
            for (Int32 ri = 0; ri < data.RowsCount; ri++)
            {
                F0Model.FilteredData.ItemArray items = data[ri];
                columnIndex = 0;
                foreach (Md.TableColumn column in Md.Table1Columns)
                {
                    Object value = items[column.ColumnName];
                    Type type = column.DataType;
                    AppendValueToSpreadsheetCell(ws, rowIndex, columnIndex, value, type);
                    columnIndex++;
                }
                rowIndex++;
            }

            /*
                        // сумма по таблице столбца "Сумма закуп."
                        {
                            String f = "=SUM(N2:N" + rowIndex.ToString() + ")";
                            ws.UpsertCell(rowIndex, 13, 3, "0", f); // CellValues.Number
                        }
                        // сумма по таблице столбца "Вес"
                        {
                            String f = "=SUM(Q2:Q" + rowIndex.ToString() + ")";
                            ws.UpsertCell(rowIndex, 16, 3, "0", f); // CellValues.Number
                        }
                        // сумма по таблице столбца "Объём"
                        {
                            String f = "=SUM(R2:R" + rowIndex.ToString() + ")";
                            ws.UpsertCell(rowIndex, 17, 3, "0", f); // CellValues.Number
                        }
                        // сумма по таблице столбца "Предельная оптовая сумма"
                        {
                            String f = "=SUM(T2:T" + rowIndex.ToString() + ")";
                            ws.UpsertCell(rowIndex, 19, 3, "0", f); // CellValues.Number
                        }
                        // сумма по таблице столбца "Сумма продажи"
                        {
                            String f = "=SUM(V2:V" + rowIndex.ToString() + ")";
                            ws.UpsertCell(rowIndex, 21, 3, "0", f); // CellValues.Number
                        }
                        rowIndex++;
                        // сумма по таблице столбца "Сумма закуп." если "Страна" == Россия
                        {
                            ws.UpsertCell(rowIndex, 12, 2, "Россия"); // CellValues.SharedString
                            String f = String.Format(
                                "=SUMIF(I{0}:I{1},\"=Россия\",N{0}:N{1})" +
                                "+SUMIF(I{0}:I{1},\"=Республика Беларусь\",N{0}:N{1})" +
                                "+SUMIF(I{0}:I{1},\"=Беларусь\",N{0}:N{1})" +
                                "+SUMIF(I{0}:I{1},\"=Казахстан\",N{0}:N{1})" +
                                "+SUMIF(I{0}:I{1},\"=Армения\",N{0}:N{1})", 2, (rowIndex - 1));
                            ws.UpsertCell(rowIndex, 13, 3, "0", f); // CellValues.Number
                            f = "=(N" + (rowIndex + 1).ToString() + "/N" + rowIndex.ToString() + ")*100";
                            ws.UpsertCell(rowIndex, 14, 3, "0", f); // CellValues.Number
                        }
                        rowIndex++;
                        // две таблицы с итогами и данными из шапки
                        {
                            ws.UpsertCell(rowIndex, 1, 2, "НМЦК"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 2, 3, h.Rows[0]["сумма_лота"]); // CellValues.Number
                            ws.UpsertCell(rowIndex, 7, 2, "Сумма по закупке (руб)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 0D, "=N" + (rowIndex - 1).ToString()); // CellValues.Number
                            rowIndex++;
                            ws.UpsertCell(rowIndex, 1, 2, "График поставки"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 2, 2, h.Rows[0]["график_поставки"] as String); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 7, 2, "Наценка (%)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 10D); // CellValues.Number
                            rowIndex++;
                            ws.UpsertCell(rowIndex, 1, 2, "Срок годности"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 2, 2, h.Rows[0]["требования_по_сроку_годности"] as String); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 7, 2, "Прибыль (руб)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 0D, "=I" + (rowIndex - 1).ToString() + "*I" + rowIndex.ToString() + "/100"); // CellValues.Number
                            rowIndex++;
                            ws.UpsertCell(rowIndex, 7, 2, "Сумма с наценкой (руб)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 0D, "=I" + (rowIndex - 2).ToString() + "+I" + rowIndex.ToString()); // CellValues.Number
                            rowIndex++;
                            ws.UpsertCell(rowIndex, 7, 2, "Транспорт (руб)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 0D); // CellValues.Number
                            rowIndex++;
                            ws.UpsertCell(rowIndex, 7, 2, "Минимальная сумма (руб)"); // CellValues.SharedString
                            ws.UpsertCell(rowIndex, 8, 3, 0D, "=I" + (rowIndex - 1).ToString() + "+I" + rowIndex.ToString()); // CellValues.Number
                            rowIndex++;
                        }
                        */
        }
        private static void AppendValueToSpreadsheetCell(XlWorksheet ws, UInt32 rowIndex, UInt32 columnIndex, Object value, Type type)
        {
            String cellValueText = null;
            if (value != null && value != DBNull.Value)
            {
                switch (type.ToString())
                {
                    case "System.Decimal":
                        ws.UpsertCell(rowIndex, columnIndex, 3, value);
                        break;
                    case "System.Double":
                        ws.UpsertCell(rowIndex, columnIndex, 3, value);
                        break;
                    case "System.Int32":
                        ws.UpsertCell(rowIndex, columnIndex, 4, value);
                        break;
                    default:
                        if (value != null)
                        {
                            cellValueText = value.ToString();
                        }
                        ws.UpsertCell(rowIndex, columnIndex, 2, cellValueText);
                        break;
                }
            }
            else
            {
                ws.UpsertCell(rowIndex, columnIndex, 2, cellValueText);
            }
        }
        private static void GenerateMergeCells(XlWorksheet ws)
        {
            // объединение ячеек для заголовка
            ws.AppendMergeCell(
                new String[] { "A1:C1", "A2:C2", "A3:C3" }
            );
        }
        private static void GenerateConditionalFormatting(XlWorksheet ws, UInt32 dxfId, DataTable t)
        {
            ws.AppendConditionalFormatting(new String[] { "K3:K" + (t.Rows.Count + 2).ToString() }, dxfId, "Value(E3)<>K3");
            ws.AppendConditionalFormatting(new String[] { "J3:J" + (t.Rows.Count + 2).ToString() }, dxfId, "D3<>J3");
        }
        private static void GenerateBackgroundColor(XlWorksheet ws, DataTable t)
        {
            for (int ri = 0; ri < t.Rows.Count; ri++)
            {
                DataRow dr = t.Rows[ri];
                String bgColor = dr["bg_color"] as String;
                if (!String.IsNullOrWhiteSpace(bgColor))
                {
                    int ci = 0;
                    foreach (Md.TableColumn col in Md.Table1Columns)
                    {
                        ws.SetCellBackgroundColor((ri + 1), ci, bgColor); // !!! 1 - количество строк в шапке
                        ci++;
                    }
                }
            }
        }
        private static void GeneratePageSetup(XlWorksheet ws)
        {
            ws.SetPageOrientationLandscape();
            ws.SetPagePaperSizeA4();
            ws.SetPageFitTo(width: 1);
        }
    }
}
