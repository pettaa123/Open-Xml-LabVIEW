# OpenXML LabVIEW

A lightweight wrapper for creating and reading `.xlsx` Excel files using Open XML.  
This wrapper has been tested against [OpenXML SDK](https://github.com/dotnet/Open-XML-SDK) version **3.3.0** and **.NET Framework 4.6**.

---

## VIPM

The `vipm` folder contains a VIPM package with everything needed to use this wrapper.  
It creates an **Open Xml** palette under *Addons*.

---

## Prerequisites

- **LabVIEW 2019**
- **(Optional) Open XML SDK** – Download the following DLLs via NuGet when cloning this repository:
  - `DocumentFormat.OpenXml.dll`
  - `DocumentFormat.OpenXml.Framework.dll`

  > **Note:** Place these DLLs in the repository’s `Source` folder.

- **(Optional) LUnit Test Framework (LabVIEW 2020)** – [GitHub Repository](https://github.com/Astemes/astemes-lunit)

---

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/pettaa123/Open-Xml-LabVIEW

## Public API

This library provides functions for setting and retrieving cell values and named ranges in an Excel worksheet.  
For example usage, refer to the `Test Open Xml` folder.

![palette](pallette.png)

---

## Features

### Set/Get Cell Value
Sets and returns the value of a given worksheet cell, handling numeric, boolean, string, datetime types, and arrays.

![setcell](sample_set_cell_value.png)

### Set/Get Cell Range Values
Sets and returns the values of a given worksheet cell range.

![setcells](sample_set_cell_value_2d_int.png)

### Set/Get Row Values
Sets and returns the values of a given worksheet row, handling numeric, boolean, string, and datetime types.

![setrowvals](sample_set_cell_row_values.png)

### Set/Get Cell Font
Sets and returns the font applied to a cell.

![setfont](sample_set_cell_font.png)

### Add/List Workbook Sheets
Adds new sheets and lists existing sheets.

![addsheet](sample_add_sheet.png)

### Get Table Rows by Column Value
Returns table rows filtered by a specified column value.

![gettablerowsbycolumnvalue](sample_get_table_rows_by_column_value.png)  
![gettablerowsbycolumnvalue_excel](sample_get_table_rows_by_column_value_excel.png)

### Get Named Range Values (String)
Retrieves a specified named range from a worksheet as strings.

![namedrangestr](sample_get_named_range_values_str.png)

### Get Named Range Values (Variant)
Retrieves a specified named range from a worksheet as variants.

![namedrangevar](sample_get_named_range_values_var.png)