using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using static OpenXmlGenericsWrapper.NonGenerics;

namespace OpenXmlGenericsWrapper.Tests
{
    [TestClass]
    public class NonGenericsTests
    {
        [TestMethod]
        public void GetTableDefinitionPart_ReturnsNull_WhenNoTables()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var result = GetTableDefinitionPart(worksheetPart, "MyTable");
                Assert.IsNull(result);
            }
        }

        [TestMethod]
        public void GetTableDefinitionPart_ReturnsFirst_WhenTableNameIsNull()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
                tablePart.Table = new Table() { Name = "Table1", Id = 1 };

                var result = GetTableDefinitionPart(worksheetPart, null);
                Assert.AreEqual("Table1", result.Table.Name);
            }
        }

        [TestMethod]
        public void GetTableDefinitionPart_ReturnsMatchingTable_WhenNameProvided()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var table1 = worksheetPart.AddNewPart<TableDefinitionPart>();
                table1.Table = new Table() { Name = "TableA", Id = 1 };

                var table2 = worksheetPart.AddNewPart<TableDefinitionPart>();
                table2.Table = new Table() { Name = "TableB", Id = 2 };

                var result = GetTableDefinitionPart(worksheetPart, "TableB");
                Assert.AreEqual("TableB", result.Table.Name);
            }
        }

        [TestMethod]
        public void InsertBeforeCell_InsertsCorrectly()
        {
            var row = new Row();
            var cell1 = new Cell() { CellReference = "A1" };
            var cell2 = new Cell() { CellReference = "B1" };
            row.Append(cell2);

            InsertBeforeCell(row, cell1, cell2);

            Assert.AreEqual(cell1, row.FirstChild);
        }

        [TestMethod]
        public void InsertBeforeRow_InsertsCorrectly()
        {
            var sheetData = new SheetData();
            var row1 = new Row() { RowIndex = 2 };
            var row2 = new Row() { RowIndex = 3 };
            sheetData.Append(row2);

            InsertBeforeRow(sheetData, row1, row2);

            Assert.AreEqual(row1, sheetData.FirstChild);
        }


        [TestMethod]
        public void SetCellDatatype_Boolean()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.Boolean);
            Assert.AreEqual(CellValues.Boolean, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_Number()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.Number);
            Assert.AreEqual(CellValues.Number, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_Error()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.Error);
            Assert.AreEqual(CellValues.Error, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_SharedString()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.SharedString);
            Assert.AreEqual(CellValues.SharedString, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_String()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.String);
            Assert.AreEqual(CellValues.String, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_InlineString()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.InlineString);
            Assert.AreEqual(CellValues.InlineString, cell.DataType.Value);
        }

        [TestMethod]
        public void SetCellDatatype_Date()
        {
            var cell = new Cell();
            SetCellDatatype(cell, DataType.Date);
            Assert.AreEqual(CellValues.Date, cell.DataType.Value);
        }



        [TestMethod]
        public void SetCellReference_SetsCorrectValue()
        {
            var cell = new Cell();
            SetCellReference(cell, "C5");
            Assert.AreEqual("C5", cell.CellReference.Value);
        }

        [TestMethod]
        public void AddNewPartToWorkbook_AddsWorksheetPart()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();

                var part = AddNewPartToWorkbook(workbookPart, WorkbookPartType.Worksheet);
                Assert.IsInstanceOfType(part, typeof(WorksheetPart));
            }
        }

        [TestMethod]
        public void AddNewPartToWorksheet_AddsTableDefinitionPart()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var part = AddNewPartToWorksheet(worksheetPart, WorksheetPartType.TableDefinition);
                Assert.IsInstanceOfType(part, typeof(TableDefinitionPart));
            }
        }

        [TestMethod]
        public void CreateUniqueTableId_Returns1_WhenNoTablesExist()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.AddNewPart<WorksheetPart>();

                var id = CreateUniqueTableId(workbookPart);
                Assert.AreEqual((uint)1, id);
            }
        }

        [TestMethod]
        public void CreateUniqueTableId_ReturnsNextId_WhenTablesExist()
        {
            using (var doc = SpreadsheetDocument.Create(new System.IO.MemoryStream(), SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var table1 = worksheetPart.AddNewPart<TableDefinitionPart>();
                table1.Table = new Table() { Id = 5 };

                var table2 = worksheetPart.AddNewPart<TableDefinitionPart>();
                table2.Table = new Table() { Id = 10 };

                var id = CreateUniqueTableId(workbookPart);
                Assert.AreEqual((uint)11, id);
            }
        }
    }
}
