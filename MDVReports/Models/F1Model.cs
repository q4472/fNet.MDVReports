using Nskd;
using System;
using System.Data;
using System.IO;

namespace FNet.MDVReports.Models
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
    public static class NskdExcel1
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
                    new TableColumn { ColumnName = "Ссылка", Caption = "Ссылка", DataType = typeof(String), Width = "120" }, // 0 A
                    new TableColumn { ColumnName = "Номер", Caption = "Номер", DataType = typeof(String), Width = "20" }, // 1 B
                    new TableColumn { ColumnName = "Дата FILE", Caption = "Дата FILE", DataType = typeof(String), Width = "14" }, // 2 C
                    new TableColumn { ColumnName = "Дата GRLS", Caption = "Дата GRLS", DataType = typeof(String), Width = "14" }, // 3 D
                    new TableColumn { ColumnName = "needToRefresh", Caption = "", DataType = typeof(Boolean), Width = "7" } // 4 E
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
        public static Byte[] ToExcel(DataTable data)
        {
            MemoryStream ms;
            UInt32 zoomScale = 100;
            String fontName = "Arial";
            Double fontSize = 10;
            using (Spreadsheet spreadsheet = new Spreadsheet(1)) // один лист
            {
                spreadsheet.SetSheetName(1, "Отчёт МДВ2");
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
        public static void GenrateSheetData0(XlWorksheet ws, DataTable data)
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
            for (Int32 ri = 0; ri < data.Rows.Count; ri++)
            {
                columnIndex = 0;
                foreach (Md.TableColumn column in Md.Table1Columns)
                {
                    Object value = data.Rows[ri][column.ColumnName];
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
