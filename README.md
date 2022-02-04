using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace dwsw.Integrations
{
    public class ImportService
    {

        private readonly ILogger<ImportService> logger;
        private readonly ArrayList arrDateStyleIndexes;
        private Logger<ImportService> logger1;

        public ImportService(ILogger<ImportService> logger)
        {
            this.logger = logger;
            this.arrDateStyleIndexes = new ArrayList {1, 8, 11, 14, 13, 15 };
        }

        public ImportService(Logger<ImportService> logger1)
        {
            this.logger1 = logger1;
        }

        #region Excel2007...

        public DataTable ImportMultipleExcel2007AsDataTable(FileInfo[] filesInfo, int startSheetIndex = 0, int startRowIndex = 0)
        {
            DataTable[] dataTableArray = new DataTable[filesInfo.Length];
            for (int i = 0; i <= filesInfo.Length - 1; i++)
            {
                dataTableArray[i] = ImportFromExcel2007AsDataTables(filesInfo[i].Open(FileMode.Open), startSheetIndex, startRowIndex)[0];
            }

            DataTable merged = new DataTable();

            foreach (var dataTable in dataTableArray)
            {
                merged.Merge(dataTable);
            }

            return merged;
        }

        public DataTable ImportMultipleExcel2007AsAMergedDataTable(Stream[] streams, int startSheetIndex = 0, int startRowIndex = 0)
        {
            DataTable[] dataTableArray = new DataTable[streams.Length];
            for (int i = 0; i <= streams.Length - 1; i++)
            {
                dataTableArray[i] = ImportFromExcel2007AsDataTables(streams[i], startSheetIndex, startRowIndex)[0];
            }

            DataTable merged = new DataTable();

            foreach (var dataTable in dataTableArray)
            {
                merged.Merge(dataTable);
            }

            return merged;
        }


        /// <summary>
        /// Import sheets in specified excel 2007 (& up) file into array of DataTable object.
        /// </summary>
        /// <param name="stream">Source excel 2007 (& up) file.</param>
        /// <param name="startSheetIndex">(Optional) Specify which is the starting sheet in zero based sheet index.</param>
        /// <param name="startRowIndex">(Optional) Specify row index from which header starts.</param>
        /// <returns>Array of DataTable object.</returns>
        public DataTable[] ImportFromExcel2007AsDataTables(Stream stream, int startSheetIndex = 0, int startRowIndex = 0)
        {
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(stream, false);

            IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            var enumerable = sheets as Sheet[] ?? sheets.Skip(startSheetIndex).ToArray();
            var list = new DataTable[enumerable.Length];

            int iCurrentSheetCounter = 0;
            foreach (Sheet sheet in enumerable)
            {
                try
                {
                    list[iCurrentSheetCounter] = GetSheetAsDataTable(spreadSheetDocument, sheet, startRowIndex);
                    iCurrentSheetCounter++;
                }
                catch (System.ObjectDisposedException)
                {
                    spreadSheetDocument = SpreadsheetDocument.Open(stream, false);
                    list[iCurrentSheetCounter] = GetSheetAsDataTable(spreadSheetDocument, sheet, startRowIndex);
                    iCurrentSheetCounter++;
                }
            }

            spreadSheetDocument.Dispose();
            return list;

        }

        public IEnumerable<dynamic> ImportFromExcel2007AsDynamicObjects(Stream stream)
        {
            return ImportFromExcel2007AsDataTables(stream)?.Select(dt => dt.AsDynamicEnumerable()).ToList();
        }

        public string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            try
            {
                //To handle System.NullReferenceException
                if (cell.CellValue == null)
                {
                    return "NULL";
                }

                SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                string value = cell.CellValue?.InnerXml;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                }
                else
                {
                    if (cell.StyleIndex != null && arrDateStyleIndexes.Contains(Convert.ToInt32((uint)cell.StyleIndex)) && value.Length > 4)//date
                    {
                        var oleValue = value;
                        if (oleValue.Contains("."))
                        {
                            oleValue = oleValue.Substring(0, oleValue.IndexOf("."));
                        }


                        if (oleValue.Length != 5)//Not an OLE Automation Date value.
                            return value;

                        try
                        {
                            value = DateTime.FromOADate(Convert.ToDouble(oleValue)).ToString();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.GetBaseException());
                        }

                    }
                    else//numeric or double
                    {
                        value = Convert.ToString(Convert.ToDouble(value));
                    }
                    return value;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.GetBaseException());
            }
            return "NULL";
        }

        /// <summary>
        /// Load specified sheet into Datatable object.
        /// </summary>
        /// <param name="spreadSheetDocument">Sheet to be parsed into datatable.</param>
        /// <param name="sheet">Sheet to be exported as DataTable.</param>
        /// <param name="startRowIndex">(Optional) Specify row index from which header starts.</param>
        /// <returns>DataTable representing excel sheet.</returns>
        public DataTable GetSheetAsDataTable(SpreadsheetDocument spreadSheetDocument, Sheet sheet, int startRowIndex = 0)
        {
            DataTable dt = new DataTable();
            using (spreadSheetDocument)
            {
                string relationshipId = sheet.Id.Value;

                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();


                var enumerable = rows as Row[] ?? rows.ToArray();
                foreach (var openXmlElement in enumerable.ElementAt(startRowIndex))
                {
                    var cell = (Cell)openXmlElement;
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in enumerable.Skip(startRowIndex))
                {
                    DataRow tempRow = dt.NewRow();
                    int columnIndex = 0;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                        cellColumnIndex--;
                        while (columnIndex < cellColumnIndex)
                        {
                            //Insert NULL here to prevent columns from shifting to left when one of the cells have a blank value;
                            tempRow[columnIndex] = "NULL";
                            columnIndex++;
                        }
                        tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);

                        columnIndex++;
                    }
                    dt.Rows.Add(tempRow);
                }
            }

            return GetTableRemovingColumnHeaderSpaces(dt);

        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        /// A length of three can be implemented when needed.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        public static int? GetColumnIndexFromName(string columnName)
        {

            //return columnIndex;
            string name = columnName;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }

        [Obsolete("Use GetSheetAsDataTable SpreadsheetDocument version.")]
        public DataTable GetSheetAsDataTable(Sheet sheet, SpreadsheetDocument spreadSheetDocument)
        {
            string relationshipId = sheet.Id.Value;
            string tabName = sheet.Name.Value;
            DataTable dataTable = new DataTable(tabName);

            WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();

            //header
            var enumerable = rows as Row[] ?? rows.ToArray();
            foreach (var openXmlElement in enumerable.ElementAt(0))
            {
                var cell = (Cell)openXmlElement;
                dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
            }

            //datarow
            try
            {
                foreach (Row row in enumerable)
                {
                    DataRow tempRow = dataTable.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dataTable.Rows.Add(tempRow);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw exception;
            }

            return GetTableRemovingColumnHeaderSpaces(dataTable);
        }

        #endregion Excel2007...

        #region CSV Import...

        public DataTable ReadCsvAsDataTable(string fileName)
        {
            DataTable dtFromCsv = new DataTable();
            using (StreamReader streamReader = new StreamReader(File.OpenRead(fileName)))
            {
                int colIndex = 0;
                while (!streamReader.EndOfStream)
                {
                    #region Column Headers...

                    if (colIndex == 0)
                    {
                        string[] columns = null;

                        var possibleDirtyLineHeader = streamReader.ReadLine();

                        if (possibleDirtyLineHeader != null && possibleDirtyLineHeader.Contains("\""))
                            possibleDirtyLineHeader = possibleDirtyLineHeader.Replace("\"", string.Empty);

                        if (possibleDirtyLineHeader != null)
                        {
                            string[] columnString = possibleDirtyLineHeader.Split(",'");

                            if (columnString != null && columnString.Length == 1)
                            {
                                //Normal CSV
                                columns = columnString[0].Split(",");
                            }
                            else
                            {
                                //Complex CSV
                                columns = columnString;
                            }
                        }

                        columns = FilterBlanks(columns);

                        if (columns != null)
                        {
                            foreach (string name in columns)
                            {
                                dtFromCsv.Columns.Add(name, typeof(String));
                            }
                        }

                    }
                    colIndex++;

                    #endregion Column Headers...

                    #region DataRows...
                    string[] dataRow = null;

                    var possibleDirtyLineRow = streamReader.ReadLine();

                    if (possibleDirtyLineRow != null && possibleDirtyLineRow.Contains("\""))
                        possibleDirtyLineRow = possibleDirtyLineRow.Replace("\"", string.Empty);

                    if (possibleDirtyLineRow != null)
                    {
                        string[] dataString = possibleDirtyLineRow.Split(",'");

                        if (dataString != null && dataString.Length == 1)
                        {
                            //Normal CSV
                            dataRow = dataString[0].Split(",");
                        }
                        else
                        {
                            //Complex CSV
                            dataRow = dataString;
                        }
                    }

                    dataRow = FilterBlanks(dataRow);

                    if (dataRow == null)
                        return null;

                    DataRow row = dtFromCsv.NewRow();
                    row.ItemArray = dataRow;
                    dtFromCsv.Rows.Add(row);

                    #endregion DataRows...
                }
            }

            DataTable dataTable = GetTableRemovingColumnHeaderSpaces(dtFromCsv);

            return dataTable;
        }

        #endregion CSV Import...


        public static DataTable GetTableRemovingColumnHeaderSpaces(DataTable dataTable)
        {
            for (var i = 0; i <= dataTable.Columns.Count - 1; i++)
            {
                dataTable.Columns[i].ColumnName = RemoveWhiteSpacesAndSpecialChars(dataTable.Columns[i].ColumnName);
            }
            return RemoveEmptyRowsFromDataTable(dataTable);
        }

        private static DataTable RemoveEmptyRowsFromDataTable(DataTable dt)
        {
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                if (dt.Rows[i][1] == DBNull.Value)
                    dt.Rows[i].Delete();

            }
            dt.AcceptChanges();

            dt = dt.DefaultView.ToTable( /*distinct*/ true);

            return dt;
        }

        public static string RemoveWhiteSpacesAndSpecialChars(string input)
        {
            input = RemoveWhitespace(input);
            input = RemoveSpecialCharacters(input);
            input = RemoveNewLineCharFromColumnName(input);
            return input;
        }

        public static string RemoveWhitespace(string input)
        {
            return new string(input.ToCharArray()
                .Where(c => !Char.IsWhiteSpace(c))
                .ToArray());
        }

        public static string RemoveSpecialCharacters(string input)
        {
            return Regex.Replace(input, @"[^0-9a-zA-Z]+", "");
        }

        public static string RemoveNewLineCharFromColumnName(string input)
        {
            return Regex.Replace(input, @"(\r\n|\r|\n)", "");
        }

        public static string[] FilterBlanks(string[] columns)
        {
            var result = columns.Select(x => x.Replace("'", "").Replace(",", "").Replace("\"", "")).Where(x => !string.IsNullOrEmpty(x)).ToArray();
            return result;
        }
    }
}

