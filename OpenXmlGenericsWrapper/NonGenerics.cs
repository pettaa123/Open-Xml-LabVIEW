using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using static OpenXmlGenericsWrapper.NonGenerics;


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

        public static void InsertBeforeCell(Row row, Cell newChild, Cell referenceChild)
        {
            row.InsertBefore<Cell>(newChild, referenceChild);
            return;
        }

        public static void InsertBeforeRow(SheetData sheetData, Row newRow, Row referenceRow)
        {
            sheetData.InsertBefore<Row>(newRow, referenceRow);
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
                case DataType.Boolean:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                    break;
                case DataType.Number:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case DataType.Error:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Error);
                    break;
                case DataType.SharedString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                case DataType.String:
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case DataType.InlineString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                    break;
                case DataType.Date:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
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
    }
}
