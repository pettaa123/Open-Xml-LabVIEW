using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;


namespace OpenXmlGenericsWrapper
{
    // Non-generic wrapper: ensures a SharedStringTablePart exists
    public static class NonGenerics
    {
        public static TableDefinitionPart GetTableDefinitionPart(WorksheetPart worksheetPart, string tableName)
        {
            // If there are no table definitions, return null
            if (worksheetPart?.TableDefinitionParts == null || !worksheetPart.TableDefinitionParts.Any())
            {
                return null;
            }

            // If tableName is null or empty, return the first table definition
            if (string.IsNullOrEmpty(tableName))
            {
                return worksheetPart.TableDefinitionParts.FirstOrDefault();
            }

            // Otherwise, try to find the table by name
            return worksheetPart.TableDefinitionParts
                                .FirstOrDefault(t => string.Equals(t.Table?.Name, tableName, StringComparison.OrdinalIgnoreCase));
        }

        public static void InsertAfterCell(Row row, Cell referenceChild, Cell newChild)
        {
            row.InsertAfter<Cell>(newChild, referenceChild);
            return;
        }

        public static void InsertAfterRow(SheetData sheetData, Row referenceRow, Row newRow)
        {
            sheetData.InsertAfter<Row>(newRow,referenceRow);
            return;
        }

        public enum DataType
        {
            Boolean,
            Number,
            Error,
            SharedString,
            String,
            InlineString,
            Date
        }

        public static void SetCellDatatype(Cell cell, DataType datatype)
        {
            switch (datatype)
            {
                case DataType.Number:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case DataType.SharedString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                case DataType.Date:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                    break;
                case DataType.Boolean:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                    break;
                case DataType.String:
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case DataType.InlineString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                    break;
                case DataType.Error:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Error);
                    break;
                default:
                    cell.DataType = null;
                    break;
            }

            return;
        }

        public static void SetCellReference(Cell cell, string value)
        {
            cell.CellReference = new StringValue(value);
            return;
        }

        public enum WorkbookPartType
        {
            Connections,
            PivotTableCacheDefinition,
            PivotTableCacheRecords,
            SharedStringTable,
            Worksheet,
            WorkbookStyles
        }
        // Add a WorksheetPart to the WorkbookPart
        public static OpenXmlPart AddNewPartToWorkbook(WorkbookPart workbookPart, WorkbookPartType partType)
        {
            switch (partType)
            {
                case WorkbookPartType.Worksheet:
                    return workbookPart.AddNewPart<WorksheetPart>();
                case WorkbookPartType.SharedStringTable:
                    return workbookPart.AddNewPart<SharedStringTablePart>();
                case WorkbookPartType.WorkbookStyles:
                    return workbookPart.AddNewPart<WorkbookStylesPart>();
                default:
                    throw new ArgumentException("Unsupported WorkbookPartType");
            }
        }
        public enum WorksheetPartType
        {
            Chart,
            Drawings,
            ImagePart,
            PivotTable,
            TableDefinition,
            WorksheetComments,
            WorksheetDrawing
        }

        public static OpenXmlPart AddNewPartToWorksheet(WorksheetPart worksheetPart, WorksheetPartType partType)
        {
            switch (partType)
            {
                case WorksheetPartType.TableDefinition:
                    return worksheetPart.AddNewPart<TableDefinitionPart>();
                default:
                    throw new ArgumentException("Unsupported WorksheetPartType");
            }
        }

        public static uint CreateUniqueTableId(WorkbookPart workbookPart)
        {
            // Default starting Id if no tables exist
            uint nextId = 1;

            // Collect all existing table Ids across all worksheets
            var existingIds = workbookPart.WorksheetParts
                .SelectMany(ws => ws.TableDefinitionParts)
                .Where(t => t.Table?.Id != null)
                .Select(t => t.Table.Id.Value);

            if (existingIds.Any())
            {
                nextId = existingIds.Max() + 1;
            }

            return nextId;
        }

        public enum WorksheetChildType
        {
            Columns,
            SheetViews,
            SheetFormatPr,
            PageMargins,
            PageSetup,
            HeaderFooter,
            ConditionalFormatting,
            DataValidations,
            MergeCells,
            Hyperlinks,
            Drawing,
            ExtLst
        }
        public static void InsertWorksheetChild(Worksheet worksheet, WorksheetChildType type, Object newChild)
        {
            switch (type)
            {
                case WorksheetChildType.Columns:
                    // Columns must come before SheetData
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    worksheet.InsertBefore((Columns)newChild, sheetData);
                    break;

                case WorksheetChildType.SheetViews:
                    // SheetViews must come before SheetFormatPr and SheetData
                    var sheetFormatPr = worksheet.GetFirstChild<SheetFormatProperties>();
                    if (sheetFormatPr != null)
                        worksheet.InsertBefore((SheetViews)newChild, sheetFormatPr);
                    else
                    {
                        var sheetDataForViews = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertBefore((SheetViews)newChild, sheetDataForViews);
                    }
                    break;

                case WorksheetChildType.SheetFormatPr:
                    // SheetFormatPr must come before Columns and SheetData
                    var columns = worksheet.GetFirstChild<Columns>();
                    if (columns != null)
                        worksheet.InsertBefore((SheetFormatProperties)newChild, columns);
                    else
                    {
                        var sheetDataForFormat = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertBefore((SheetFormatProperties)newChild, sheetDataForFormat);
                    }
                    break;

                case WorksheetChildType.PageMargins:
                    // PageMargins must come after SheetData
                    var sheetDataForMargins = worksheet.GetFirstChild<SheetData>();
                    worksheet.InsertAfter((PageMargins)newChild, sheetDataForMargins);
                    break;

                case WorksheetChildType.PageSetup:
                    // PageSetup must come after PageMargins
                    var pageMargins = worksheet.GetFirstChild<PageMargins>();
                    if (pageMargins != null)
                        worksheet.InsertAfter((PageSetup)newChild, pageMargins);
                    else
                    {
                        var sheetDataForSetup = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertAfter((PageSetup)newChild, sheetDataForSetup);
                    }
                    break;

                case WorksheetChildType.HeaderFooter:
                    // HeaderFooter must come after PageSetup
                    var pageSetup = worksheet.GetFirstChild<PageSetup>();
                    if (pageSetup != null)
                        worksheet.InsertAfter((HeaderFooter)newChild, pageSetup);
                    else
                    {
                        var sheetDataForHeader = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertAfter((HeaderFooter)newChild, sheetDataForHeader);
                    }
                    break;

                case WorksheetChildType.ConditionalFormatting:
                    // ConditionalFormatting must come after SheetData
                    var sheetDataForCond = worksheet.GetFirstChild<SheetData>();
                    worksheet.InsertAfter((ConditionalFormatting)newChild, sheetDataForCond);
                    break;

                case WorksheetChildType.DataValidations:
                    // DataValidations must come after ConditionalFormatting or SheetData
                    var condFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
                    if (condFormatting != null)
                        worksheet.InsertAfter((DataValidations)newChild, condFormatting);
                    else
                    {
                        var sheetDataForValidations = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertAfter((DataValidations)newChild, sheetDataForValidations);
                    }
                    break;

                case WorksheetChildType.MergeCells:
                    // MergeCells must come after SheetData
                    var sheetDataForMerge = worksheet.GetFirstChild<SheetData>();
                    worksheet.InsertAfter((MergeCells)newChild, sheetDataForMerge);
                    break;

                case WorksheetChildType.Hyperlinks:
                    // Hyperlinks must come after DataValidations or SheetData
                    var dataValidations = worksheet.GetFirstChild<DataValidations>();
                    if (dataValidations != null)
                        worksheet.InsertAfter((Hyperlinks)newChild, dataValidations);
                    else
                    {
                        var sheetDataForLinks = worksheet.GetFirstChild<SheetData>();
                        worksheet.InsertAfter((Hyperlinks)newChild, sheetDataForLinks);
                    }
                    break;

                case WorksheetChildType.Drawing:
                    // Drawing must come after HeaderFooter or PageSetup
                    var headerFooter = worksheet.GetFirstChild<HeaderFooter>();
                    if (headerFooter != null)
                        worksheet.InsertAfter((Drawing)newChild, headerFooter);
                    else
                    {
                        var pageSetupForDrawing = worksheet.GetFirstChild<PageSetup>();
                        if (pageSetupForDrawing != null)
                            worksheet.InsertAfter((Drawing)newChild, pageSetupForDrawing);
                        else
                        {
                            var sheetDataForDrawing = worksheet.GetFirstChild<SheetData>();
                            worksheet.InsertAfter((Drawing)newChild, sheetDataForDrawing);
                        }
                    }
                    break;

                case WorksheetChildType.ExtLst:
                    // ExtLst must come last
                    worksheet.Append((ExtensionList)newChild);
                    break;
            }
            return;
        }
    }
}
