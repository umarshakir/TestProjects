﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ColumnMerge
{
    public class GenerateExcel
    {
        public static void GenerateOrderDetailsExcel(string orderNumber,
                                                   string orderDetailsPath, string headers, int indexToRead,
                                                   string destinationFile)
        {
            DataTable lstOrderDetails = ReadAsDataTable(orderDetailsPath, indexToRead);
            DataTable result = new DataTable();
            int index = 0;
            string[] resultHeaders = headers.Split(',');
            foreach (string header in resultHeaders)
            {
                result.Columns.Add(header, typeof(string));
            }
            foreach(DataRow row in lstOrderDetails.Rows)
            {
                DataRow dataRow = result.NewRow();
                dataRow[0] = orderNumber;
                dataRow[1] = row[resultHeaders[1]];
                result.Rows.Add(dataRow);
            }
            CreateExcelFile(result, destinationFile, resultHeaders.Length);
        }
        private static DataTable ReadAsDataTable(string fileName, int indexToRead)
        {
            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                int counter = 0;
                foreach (Cell cell in rows.ElementAt(0))
                {
                    if(counter++ == indexToRead)
                    {
                        dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    }                  
                }
                foreach (Row row in rows)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        if (i == indexToRead)
                        {
                            dataRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            dataTable.Rows.RemoveAt(0);
            return dataTable;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
        private static void CreatePartsForExcel(SpreadsheetDocument document, DataTable data, int colLength)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data, colLength);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }
        public static void CreateExcelFile(DataTable data, string OutPutFileDirectory, int colLength)
        {
            //var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            //string fileFullname = Path.Combine(OutPutFileDirectory, "Output.xlsx");

            //if (File.Exists(fileFullname))
            //{
            //    fileFullname = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");
            //}

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(OutPutFileDirectory, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data, colLength);
            }
        }
        private static void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }

        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet1;
        }

        private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        private static SheetData GenerateSheetdataForDetails(DataTable data, int colLength)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel(data));

            foreach (DataRow testmodel in data.Rows)
            {
                Row partsRows = GenerateRowForChildPartDetail(testmodel, colLength);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
        private static Row CreateHeaderRowForExcel(DataTable data)
        {
            Row workRow = new Row();
            foreach (DataColumn col in data.Columns)
            {
                workRow.Append(CreateCell(col.ColumnName, 2U));
            }
            return workRow;
        }

        private static Row GenerateRowForChildPartDetail(DataRow testmodel, int colLength)
        {
            Row tRow = new Row();
            for (int i = 0; i < colLength; i++)
            {
                tRow.Append(CreateCell(Convert.ToString(testmodel[i])));
            }
            //tRow.Append(CreateCell(testmodel.TestId.ToString()));
            //tRow.Append(CreateCell(testmodel.TestName));
            //tRow.Append(CreateCell(testmodel.TestDesc));
            //tRow.Append(CreateCell(testmodel.TestDate.ToShortDateString()));

            return tRow;
        }

        private static Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private static Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}
//https://www.c-sharpcorner.com/article/creating-excel-file-using-openxml/