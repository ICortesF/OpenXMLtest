using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testread
{
    class Program
    {
        static void Main(string[] args)
        {
            //string fileName = @"D:\Desarrollo\test\OpenXML\testread\libro1.xlsx";
            string fileName = @"..\..\libro1.xlsx";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>()
                                  .FirstOrDefault();


                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                int numerofila = 0;
                
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        numerofila += 1;
                    }
                    if (reader.ElementType == typeof(Cell))
                    {
                        Cell theCell = (Cell)reader.LoadCurrentElement();
                        string value = null;

                        if (theCell != null)
                        {
                            value = theCell.InnerText;
                            if (theCell.DataType != null)
                            {
                                switch (theCell.DataType.Value)
                                {
                                    case CellValues.SharedString:

                                        // For shared strings, look up the value in the
                                        // shared strings table.


                                        // If the shared string table is missing, something 
                                        // is wrong. Return the index that is in
                                        // the cell. Otherwise, look up the correct text in 
                                        // the table.
                                        if (stringTable != null)
                                        {
                                            value =
                                                stringTable.SharedStringTable
                                                .ElementAt(int.Parse(value)).InnerText;
                                        }
                                        break;

                                    case CellValues.Boolean:
                                        switch (value)
                                        {
                                            case "0":
                                                value = "FALSE";
                                                break;
                                            default:
                                                value = "TRUE";
                                                break;
                                        }
                                        break;
                                }
                            }
                            else
                            {
                                //Datetime it is read in OADate format
                                if (theCell.StyleIndex != null && value.Length > 0)
                                {
                                    value = DateTime.FromOADate(double.Parse(value)).ToShortDateString();
                                }
                            }
                            string columna = GetColumn(theCell.CellReference);
                        }
                    }
                }
            }
        }

        public static int GetRow(string cellreference)
        {
            string aux = "";
            int nposfinal = 0;
            for (int npos = 0; !Char.IsNumber(cellreference, npos); npos++)
                nposfinal = npos;

            aux = cellreference.Substring(nposfinal + 1);
            return Int32.Parse(aux);

        }
        public static string GetColumn(string cellreference)
        {
            string columna = "";
            int nposfinal = 0;
            for (int npos = 0; !Char.IsNumber(cellreference, npos); npos++)
                nposfinal = npos;

            columna = cellreference.Substring(0, nposfinal + 1);

            return columna;
        }
    }
}
