using System;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FNet.MDVReports.Models
{
    public class Spreadsheet : IDisposable
    {
        private MemoryStream memoryStream;              // место где будет создан документ
        private SpreadsheetDocument document;           // документ котрый может содержать несколько WorkbookPart
        private WorkbookPart workbookPart1;             // у нас будет только одна часть workbookPart1
        private Workbook workbook;                      // у workbookPart1 есть свойство Workbook 
        private Sheets sheets;                          // у workbook есть листы
        private Sheet sheet;                            // у нас будет столько листов сколько их зададут в конструкторе Spreadsheet (по умолчанию один)
        private WorkbookStylesPart workbookStylesPart;  // у workbookPart1 также может быть несколько Parts
        private Stylesheet stylesheet;                  // Fonts, Fills, Borders, CellStyleFormats, CellFormats, CellStyles, DifferentialFormats
        private SharedStringTablePart sharedStringTablePart;    //
        private SharedStringTable sharedStringTable;    //
        private WorksheetPart[] worksheetParts;         //
        private Worksheet worksheet;                    // SheetViews, Columns, SheetData, MergeCells, ConditionalFormatting

        public XlWorksheet[] Wss;                      // массив экземпляров WS для заполненых worksheet с доп процедурами

        public Spreadsheet(Int32 sheetCount = 1)
        {
            memoryStream = new System.IO.MemoryStream();

            document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);

            workbookPart1 = document.AddWorkbookPart();
            workbook = new Workbook();
            workbookPart1.Workbook = workbook;

            workbookStylesPart = workbookPart1.AddNewPart<WorkbookStylesPart>();
            stylesheet = (new StylesPart()).GenerateWorkbookStylesPartContent();
            workbookStylesPart.Stylesheet = stylesheet;

            sharedStringTablePart = workbookPart1.AddNewPart<SharedStringTablePart>();
            sharedStringTable = (new SstPart()).GenerateSharedStringTablePart();
            sharedStringTablePart.SharedStringTable = sharedStringTable;

            sheets = new Sheets();
            workbook.Append(sheets);

            worksheetParts = new WorksheetPart[sheetCount];
            Wss = new XlWorksheet[sheetCount];
            for (UInt32 i = 0; i < sheetCount; i++)
            {
                // создаём новый лист со всеми частями (SheetViews, Columns, SheetData, MergeCells) (пока пустыми)
                Wss[i] = new XlWorksheet(stylesheet, sharedStringTable);
                worksheet = Wss[i].GetWorksheet();

                // создаём новый WorksheetPart в workbookPart
                worksheetParts[i] = workbookPart1.AddNewPart<WorksheetPart>();
                // и добавляем новый лист в его корневой элемент Worksheet 
                worksheetParts[i].Worksheet = worksheet;
                // ссылка на новый WorksheetPart 
                String rId = workbookPart1.GetIdOfPart(worksheetParts[i]);

                // запись о новом листе в workbook
                sheet = new Sheet() { Name = "Sheet" + (i + 1).ToString(), SheetId = (i + 1), Id = rId };
                sheets.Append(sheet);
            }
        }
        public void Dispose()
        {
            document.Dispose();
            memoryStream.Dispose();
        }
        public MemoryStream CreateDocument()
        {
            document.Close();
            return memoryStream;
        }
        public UInt32 AppendDifferentialFormat(String cssFormat)
        {
            DifferentialFormats differentialFormats = workbookStylesPart.Stylesheet.GetFirstChild<DifferentialFormats>();
            UInt32 dxfId = (UInt32)differentialFormats.ChildElements.Count;

            String[] formats = cssFormat.Split(';');
            foreach (String format in formats)
            {
                if (!String.IsNullOrWhiteSpace(format))
                {
                    String[] kvp = format.Split(':');
                    if (kvp.Length > 1)
                    {
                        if (kvp[0].Trim() == "font-weight" && kvp[1].Trim() == "bold")
                        {
                            differentialFormats.Append(new DifferentialFormat() { Font = new Font() { Bold = new Bold() } });
                        }
                    }
                }
            }
            return dxfId;
        }
        public void SetFont(Int32 index, String fontName, Double fontSize)
        {
            Font font = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt<Font>(index);
            if (fontSize >= 4 && fontSize <= 32)
            {
                font.GetFirstChild<FontSize>().Val = fontSize;
            }
            if (!String.IsNullOrWhiteSpace(fontName))
            {
                font.GetFirstChild<FontName>().Val = fontName;
            }
        }
        public void SetZoomScale(UInt32 zoomScale)
        {
            SheetViews sheetViews = worksheet.GetFirstChild<SheetViews>();
            SheetView sheetView = sheetViews.GetFirstChild<SheetView>();
            sheetView.ZoomScale = zoomScale;
        }
        public void SetSheetName(Int32 sheetNumber, String sheetName)
        {
            sheets.Elements<Sheet>().ElementAt(sheetNumber - 1).Name = sheetName;
        }
        private class StylesPart
        {
            public Stylesheet GenerateWorkbookStylesPartContent()
            {
                Stylesheet stylesheet = new Stylesheet();
                stylesheet.Append(GenerateFonts());
                stylesheet.Append(GenerateFills());
                stylesheet.Append(GenerateBorders());
                stylesheet.Append(GenerateCellStyleFormats());
                stylesheet.Append(GenerateCellFormats());
                stylesheet.Append(GenerateCellStyles());
                stylesheet.Append(GenerateDifferentialFormats());
                return stylesheet;
            }
            private Fonts GenerateFonts()
            {
                Fonts fonts = new Fonts();
                {
                    Font font0 = new Font();
                    font0.Append(new FontSize() { Val = 9 });
                    font0.Append(new FontName() { Val = "Arial" });

                    Font font1 = new Font();
                    font1.Append(new FontSize() { Val = 9 });
                    font1.Append(new FontName() { Val = "Arial" });

                    fonts.Append(font0);
                    fonts.Append(font1);
                }
                return fonts;
            }
            private Fills GenerateFills()
            {
                Fills fills = new Fills();

                // 0
                fills.Append(new Fill());

                // 1
                fills.Append(
                    new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.Gray125
                        }
                    }
                );

                // 2
                fills.Append(
                    new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.Solid,
                            ForegroundColor = new ForegroundColor() { Rgb = "ffddeeff" }
                        }
                    }
                );

                return fills;
            }
            private Borders GenerateBorders()
            {
                Borders borders = new Borders();

                // 0
                borders.Append(new Border());

                // 1
                borders.Append(new Border()
                {
                    LeftBorder = new LeftBorder() { Style = BorderStyleValues.Thin },
                    RightBorder = new RightBorder() { Style = BorderStyleValues.Thin },
                    TopBorder = new TopBorder() { Style = BorderStyleValues.Thin },
                    BottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin }
                });

                return borders;
            }
            private CellStyleFormats GenerateCellStyleFormats()
            {
                CellStyleFormats cellStyleFormats = new CellStyleFormats();
                cellStyleFormats.Append(new CellFormat());
                return cellStyleFormats;
            }
            private CellFormats GenerateCellFormats()
            {
                CellFormats cellFormats = new CellFormats();

                // 0 - обычные ячейки по умолчанию
                cellFormats.Append(
                    (new CellFormat() { FormatId = 0 })
                );

                // 1 - заголовки в таблице
                cellFormats.Append(
                    new CellFormat()
                    {
                        FormatId = 0,
                        FontId = 1,
                        FillId = 2,
                        BorderId = 1,
                        Alignment = new Alignment() { Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    }
                );

                // 2 - данные в таблице (строки)
                cellFormats.Append(
                    new CellFormat()
                    {
                        FormatId = 0,
                        FontId = 1,
                        FillId = 0,
                        BorderId = 1,
                        Alignment = new Alignment() { Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    }
                );

                // 3 - данные в таблице (цифры # ##0.00)
                cellFormats.Append(
                    new CellFormat()
                    {
                        FormatId = 0,
                        FontId = 1,
                        FillId = 0,
                        BorderId = 1,
                        NumberFormatId = 4,
                        Alignment = new Alignment() { Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    }
                );

                // 4 - данные в таблице (цифры # ##0)
                cellFormats.Append(
                    new CellFormat()
                    {
                        FormatId = 0,
                        FontId = 1,
                        FillId = 0,
                        BorderId = 1,
                        NumberFormatId = 1,
                        Alignment = new Alignment() { Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    }
                );

                return cellFormats;
            }
            private CellStyles GenerateCellStyles()
            {
                CellStyles cellStyles = new CellStyles();
                cellStyles.Append(new CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 });
                return cellStyles;
            }
            private DifferentialFormats GenerateDifferentialFormats()
            {
                DifferentialFormats differentialFormats = new DifferentialFormats();
                return differentialFormats;
            }
        }
        private class SstPart
        {
            public SharedStringTable GenerateSharedStringTablePart()
            {
                return (GenerateSharedStringTable());
            }
            private SharedStringTable GenerateSharedStringTable()
            {
                SharedStringTable sharedStringTable = new SharedStringTable();
                return sharedStringTable;
            }
        }
    }
    public class XlWorksheet
    {
        private Stylesheet stylesheet;
        private SharedStringTable sharedStringTable;
        private Worksheet worksheet;

        private SheetViews GenerateSheetViews()
        {
            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView() { ZoomScale = (UInt32Value)100U, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            sheetViews.Append(sheetView);
            return sheetViews;
        }
        private SheetData GenerateSheetData()
        {
            SheetData sheetData = new SheetData();
            return sheetData;
        }
        private Row GetRow(UInt32 index) // zero-based
        {
            Row row = null;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            foreach (Row r in sheetData.Elements<Row>())
            {
                if (r.RowIndex == (index + 1))
                {
                    row = r;
                    break;
                }
            }
            return row;
        }
        private Cell GetCell(Row row, UInt32 index) // zero-based
        {
            Cell cell = null;
            foreach (Cell c in row.Elements<Cell>())
            {
                String r = GetColumnName(index) + row.RowIndex;
                if (c.CellReference == r)
                {
                    cell = c;
                    break;
                }
            }
            return cell;
        }
        private void SetCell(Cell cell, UInt32 styleIndex, Object cellValue = null, String formula = null)
        {
            cell.StyleIndex = styleIndex;
            if (cellValue != null && cellValue != DBNull.Value)
            {
                System.Globalization.CultureInfo ic = System.Globalization.CultureInfo.InvariantCulture;
                switch (cellValue.GetType().ToString())
                {
                    case "System.Decimal":
                        cell.DataType = CellValues.Number;
                        cell.Append(new CellValue() { Text = ((Decimal)cellValue).ToString(ic) });
                        break;
                    case "System.Double":
                        cell.DataType = CellValues.Number;
                        cell.Append(new CellValue() { Text = ((Double)cellValue).ToString(ic) });
                        break;
                    case "System.Int32":
                        cell.DataType = CellValues.Number;
                        cell.Append(new CellValue() { Text = ((Int32)cellValue).ToString(ic) });
                        break;
                    case "System.String":
                    default:
                        cell.DataType = CellValues.SharedString;
                        Int32 index = AddStringToSst(cellValue as String);
                        cell.Append(new CellValue() { Text = index.ToString() });
                        break;
                }
            }
            if (formula != null)
            {
                //cell.DataType = CellValues.Number;
                cell.CellFormula = new CellFormula(formula);
            }
        }
        private Int32 AddStringToSst(String s)
        {
            SharedStringTable sst = sharedStringTable;
            Int32 index = sst.ChildElements.Count;
            SharedStringItem sharedStringItem = new SharedStringItem();
            Text text = new Text() { Text = (s ?? "") };
            sharedStringItem.Append(text);
            sst.Append(sharedStringItem);
            return index;
        }
        private String GetColumnName(UInt32 index) // zero-based
        {
            const byte BASE = 'Z' - 'A' + 1;
            string name = String.Empty;
            do
            {
                name = Convert.ToChar('A' + index % BASE) + name;
                index = index / BASE;
            }
            while (index-- > 0);
            return name;
        }
        private Row AppendRow(UInt32 rowIndex)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row = new Row() { RowIndex = (rowIndex + 1) };
            sheetData.Append(row);
            return row;
        }
        private PageSetup GetPageSetup()
        {
            PageSetup pageSetup = worksheet.GetFirstChild<PageSetup>();
            if (pageSetup == null)
            {
                pageSetup = new PageSetup();
                // ищем куда вставить
                ConditionalFormatting conditionalFormatting = null;
                var conditionalFormattings = worksheet.Elements<ConditionalFormatting>();
                if (conditionalFormattings != null && conditionalFormattings.Count<ConditionalFormatting>() > 0)
                {
                    conditionalFormatting = conditionalFormattings.Last<ConditionalFormatting>();
                }
                if (conditionalFormatting != null)
                {
                    conditionalFormatting.InsertAfterSelf<PageSetup>(pageSetup);
                }
                else
                {
                    MergeCells mergeCells = worksheet.GetFirstChild<MergeCells>();
                    if (mergeCells != null)
                    {
                        mergeCells.InsertAfterSelf<PageSetup>(pageSetup);
                    }
                    else
                    {
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                        if (sheetData != null)
                        {
                            sheetData.InsertAfterSelf<PageSetup>(pageSetup);
                        }
                    }
                }
            }
            return pageSetup;
        }

        public XlWorksheet(Stylesheet stylesheet, SharedStringTable sharedStringTable)
        {
            this.stylesheet = stylesheet;
            this.sharedStringTable = sharedStringTable;
            worksheet = new Worksheet();
            worksheet.Append(GenerateSheetViews());
            worksheet.Append(GenerateSheetData()); //  это необходимая часть
        }
        public void SetRowHeight(UInt32 rowIndex, Double height)
        {
            Row row = GetRow(rowIndex);
            if (row == null)
            {
                row = AppendRow(rowIndex);
            }
            row.CustomHeight = true;
            row.Height = height;
        }
        public void UpsertCell(UInt32 rowIndex, UInt32 columnIndex, UInt32 styleIndex, Object cellValue = null, String formula = null)
        {
            Row row = GetRow(rowIndex);
            if (row == null)
            {
                row = AppendRow(rowIndex);
            }
            Cell cell = GetCell(row, columnIndex);
            if (cell == null)
            {
                String reference = GetColumnName(columnIndex) + (rowIndex + 1).ToString();
                cell = new Cell() { CellReference = reference };
                SetCell(cell, styleIndex, cellValue, formula);
                row.Append(cell);
            }
            else
            {
                cell.RemoveAllChildren();
                SetCell(cell, styleIndex, cellValue, formula);
            }
        }
        public Worksheet GetWorksheet()
        {
            return worksheet;
        }
        public void AppendColumn(UInt32 min = 1, UInt32 max = 1, Boolean bestFit = true, Boolean customWidth = true, Double width = 1)
        {
            Columns columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                worksheet.GetFirstChild<SheetViews>().InsertAfterSelf<Columns>(columns);
            }
            Column column = new Column() { Min = min, Max = max };
            if (bestFit)
            {
                column.BestFit = true;
            }
            else if (customWidth)
            {
                column.CustomWidth = true;
                column.Width = width;
            }
            columns.Append(column);
        }
        public void SetCellBackgroundColor(Int32 rowIndex, Int32 columnIndex, String bgColor)
        {
            // ищем ячейку и её старый формат
            Row row = GetRow((UInt32)rowIndex);
            Cell cell = GetCell(row, (UInt32)columnIndex);
            CellFormats cellFormats = stylesheet.GetFirstChild<CellFormats>();
            CellFormat cellFormat = cellFormats.Elements<CellFormat>().ElementAt<CellFormat>((Int32)cell.StyleIndex.Value);
            Int32 newCellFormatIndex = cellFormats.Elements<CellFormat>().Count<CellFormat>();
            CellFormat newCellFormat = (CellFormat)cellFormat.Clone();
            {
                // новый background ----------------------------------------- проверить существующий to do ...
                Fills fills = stylesheet.GetFirstChild<Fills>();
                Int32 fillIndex = fills.Elements<Fill>().Count<Fill>();
                fills.Append(
                    new Fill
                    {
                        PatternFill = new PatternFill
                        {
                            PatternType = PatternValues.Solid,
                            ForegroundColor = new ForegroundColor() { Rgb = "ff" + bgColor }
                        }
                    }
                );
                newCellFormat.FillId = (UInt32)fillIndex;
            }
            cellFormats.Append(newCellFormat);
            cell.StyleIndex = (UInt32)newCellFormatIndex;
        }
        public void AppendConditionalFormatting(String[] sequenceOfReferences, UInt32 formatId, String formula)
        {
            ConditionalFormatting conditionalFormatting = new ConditionalFormatting()
            {
                SequenceOfReferences = new ListValue<StringValue>()
                {
                    InnerText = sequenceOfReferences[0]  // пока только первая стока
                }
            };
            ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule()
            {
                Type = ConditionalFormatValues.Expression,
                FormatId = formatId,
                Priority = 1
            };
            conditionalFormattingRule.Append(new Formula() { Text = formula });
            conditionalFormatting.Append(conditionalFormattingRule);
            worksheet.Append(conditionalFormatting);
        }
        public void AppendMergeCell(String reference)
        {
            MergeCells mergeCells = worksheet.GetFirstChild<MergeCells>();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheet.GetFirstChild<SheetData>().InsertAfterSelf<MergeCells>(mergeCells);
            }
            mergeCells.Append(new MergeCell() { Reference = reference });
        }
        public void AppendMergeCell(String[] references)
        {
            MergeCells mergeCells = worksheet.GetFirstChild<MergeCells>();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheet.GetFirstChild<SheetData>().InsertAfterSelf<MergeCells>(mergeCells);
            }
            foreach (String reference in references)
            {
                mergeCells.Append(new MergeCell() { Reference = reference });
            }
        }
        public void SetPageOrientationLandscape()
        {
            PageSetup pageSetup = GetPageSetup();
            if (pageSetup != null)
            {
                pageSetup.Orientation = OrientationValues.Landscape;
            }
        }
        public void SetPagePaperSizeA4()
        {
            PageSetup pageSetup = GetPageSetup();
            if (pageSetup != null)
            {
                pageSetup.PaperSize = 9;
            }
        }
        public void SetPageFitTo(UInt32 width, UInt32 height = 32767)
        {
            SheetProperties sheetProperties = worksheet.SheetProperties;
            if (sheetProperties == null)
            {
                sheetProperties = new SheetProperties();
                worksheet.SheetProperties = sheetProperties;
            }

            PageSetupProperties pageSetupProperties = sheetProperties.PageSetupProperties;
            if (pageSetupProperties == null)
            {
                pageSetupProperties = new PageSetupProperties();
                sheetProperties.PageSetupProperties = pageSetupProperties;
            }
            pageSetupProperties.FitToPage = true;

            PageSetup pageSetup = GetPageSetup();
            if (pageSetup != null)
            {
                pageSetup.FitToWidth = width;
                pageSetup.FitToHeight = height;
            }
        }
    }
}

