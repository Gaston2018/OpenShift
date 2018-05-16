using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace app.Models
{
    public class DocumentosOpenXml
    {
        private string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.CellValue != null)
            {
                string value = cell.CellValue.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                }
                return value;
            }
            return null;
        }

        public DataTable ConvertExceltoDataTable(string rutaExcel)
        {
            DataTable dt = new DataTable();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(rutaExcel, false))
            {

                Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                foreach (Row row in rows)
                {
                    //Read the first row as header
                    if (row.RowIndex.Value == 1)
                    {
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = GetCellValue(doc, cell);

                            try
                            {
                                dt.Columns.Add(colunmName);
                            }
                            catch (Exception ex)
                            {
                                throw;
                            }

                        }
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                            i++;
                        }
                    }
                }
            }
            return dt;
        }

        public DataTable ConvertCSVtoDataTable(string rutaCSV)
        {
            DataTable dt = new DataTable();
            bool primeraLinea = true;

            using (var reader = new StreamReader(rutaCSV, Encoding.UTF7))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    if (primeraLinea)
                    {
                        foreach (string item in values)
                        {
                            dt.Columns.Add(item);
                        }
                        primeraLinea = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (string item in values)
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = item;
                            i++;
                        }
                    }
                }
            }
            return dt;
        }
    }
}
