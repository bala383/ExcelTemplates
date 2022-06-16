using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace API.Services.Excel
{
    public class ExcelTemplateReader : IExcelTemplateReader
    {
        public byte[] ExportListToExcel<T>(MemoryStream mem, List<T> lstBookingTemplates, string SheetName,string Columns)
        {
            var workbook = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook);

            var workbookPart = workbook.AddWorkbookPart();
            workbook.WorkbookPart.Workbook = new()
            {
                Sheets = new Sheets()
            };
            workbook = GetStylesheet(workbook);

            uint sheetId = 1;
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            sheetPart.Worksheet = new Worksheet(sheetData);

            Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            if (sheets.Elements<Sheet>().Any())
            {
                sheetId =
                    sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            Sheet sheet = new() { Id = relationshipId, SheetId = sheetId, Name = SheetName };
            sheets.Append(sheet);

            Row headerRow = new();
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo[] PropsOrder = new PropertyInfo[Props.Length];
            if (Columns != null)  //based on configuration excel order sorting properties 
            {
                var orderArray = Columns.Split(',').Distinct().ToArray();
                for (int j = 0; j < orderArray.Length; j++)
                {
                    for (int i = 0; i < Props.Length; i++)
                    {
                        if (Props[i].Name.ToLower() == orderArray[j].ToLower())
                        {
                            PropsOrder[j] = Props[i];
                        }
                    }
                }
            }
            else
                Array.Copy(Props, PropsOrder, Props.Length); 
            int index=0;
            int count=0;
            foreach (PropertyInfo prop in PropsOrder)
            {
                if (prop != null)
                {
                    //Setting column names as Property names
                    Cell cell = new()
                    {
                        DataType =  CellValues.String,
                        CellValue = new CellValue(prop.Name),
                        StyleIndex = 1  // make column header as bold
                    };
                    headerRow.AppendChild(cell);
                    if (prop.Name.ToLower() == "remarks")
                        index = count;
                    count++;
                }
            }
            sheetData.AppendChild(headerRow);
            uint styleIndex = 0;
            string Remarks = "";
            foreach (var bookingTemplate in lstBookingTemplates)
            {
                if(index>0)               //checking if template type is validation or not
                Remarks = PropsOrder[index].GetValue(bookingTemplate).ToString() ?? "";
                Row newRow = new();
              
                foreach (PropertyInfo prop in PropsOrder)
                {
                    string cellValue = "";
                    if (prop != null)
                    {
                        if (prop.GetValue(bookingTemplate) != null)
                        {
                             cellValue = prop.GetValue(bookingTemplate).ToString() ?? "";
                        }
                        if (index > 0)
                        {
                            if (Remarks.ToLower().Contains(prop.Name.ToLower()))
                                styleIndex = 2; // fill red to invalid
                            if (cellValue == "Valid")
                            {
                                styleIndex = 3; // Fill green to valid
                            }
                            if (cellValue == "Invalid")
                            {
                                styleIndex = 2; // fill red to invalid
                            }
                        }
                        //Values binding to excel
                        Cell cell = new()
                        {
                            DataType = Regex.IsMatch(cellValue, @"^-?\d+\.?\d*$") ? CellValues.Number:CellValues.String,
                            CellValue = new CellValue(cellValue),
                            StyleIndex = styleIndex
                        };
                        styleIndex = 0;
                        newRow.AppendChild(cell);
                    }
                }
                sheetData.AppendChild(newRow);
            }
            workbook.WorkbookPart.Workbook.Save();
            workbook.Close();
            return mem.ToArray();
        }

        public IEnumerable<T> ReadExcelFile(IFormFile file, int bookingId, CancellationToken token)
        {
            List<T> lstObj = new List<T>();
            try
            {

                // Open a SpreadsheetDocument based on a stream.
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(file.OpenReadStream(), false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets sheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    // SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    StringBuilder excelErrorList = new StringBuilder();


                    bool isSheetFound = false;
                    foreach (Sheet thesheet in sheetcollection)
                    {
                        if (thesheet.Name == "{SheetName}")
                        {
                            isSheetFound = true;
                            break;
                        }
                    }

                    if (!isSheetFound)
                    {
                        throw new Exception("Sheet with name 'SheetName' not exist, data should be available in 'SheetName' sheet");
                    }

                    int count = 0;


                    foreach (Sheet thesheet in sheetcollection)
                    {
                        if (thesheet.Name != "{SheetName}")
                        {
                            continue;
                        }
                        bool firstRow = false;
                        Dictionary<string, string> dic = new Dictionary<string, string>();
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData sheetData = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        foreach (Row currentrow in sheetData)
                        {                           
                            if (!firstRow)
                            {
                                foreach (Cell c in currentrow.Elements<Cell>())
                                {
                                    string value = c.CellValue?.Text;
                                    value = GetValue(c, workbookPart);
                                    dic.Add(c.CellReference, value);
                                    count++;
                                }
                                firstRow = true;
                            }
                            else
                            {
                                T item = new T();

                                foreach (Cell c in currentrow.Elements<Cell>())
                                {
                                    string value = GetValue(c, workbookPart);
                                    string excelCol =  new String(c.CellReference.ToString().Where(c => Char.IsLetter(c) && Char.IsUpper(c)).ToArray());

                                    item = SetData(item, dic, value, excelCol);
                                }
                                lstObj.Add(item);
                            }
                        }
                    }
                }

            }
            catch
            {
                throw;
            }

            return lstObj;
        }

        private SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        private string GetValue(Cell cell, WorkbookPart workbookPart)
        {
            string cellValue = string.Empty;
            if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.SharedString)
                {
                    int id = -1;
                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        SharedStringItem item = GetSharedStringItemById(workbookPart, id);
                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
            }
            else
            {
                cellValue = cell.InnerText;
            }

            return cellValue;
        }

        private T SetData<T>(T item, Dictionary<string, string> dic, string value, string column)
        {
            var propName = dic.GetValueOrDefault(column + "1");
            if (!string.IsNullOrEmpty(propName))
            {
                var propInfo = item.GetType().GetProperty(propName);
                if (propInfo != null)
                {
                    if (propInfo.PropertyType == typeof(string))
                        propInfo.SetValue(item, value?.Trim());
                    else if (propInfo.PropertyType == typeof(double))
                    {
                        // double data = 0;
                        _ = double.TryParse(value, out double data);
                        propInfo.SetValue(item, data);
                    }
                    else if (propInfo.PropertyType == typeof(int))
                    {
                        //double data = 0;
                        _ = double.TryParse(value, out double data);
                        propInfo.SetValue(item, data);
                    }
                    else if (propInfo.PropertyType == typeof(decimal))
                    {
                        //double data = 0;
                        _ = decimal.TryParse(value, out decimal data);
                        propInfo.SetValue(item, data);
                    }
                }
            }
            return item;
        }

        private SpreadsheetDocument GetStylesheet(SpreadsheetDocument spreadsheet)
        {
            var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            // blank font list
            stylesPart.Stylesheet.Fonts = new Fonts();
            Font font = new(); //default
            Font font1 = new (new Bold()); //blod 
            stylesPart.Stylesheet.Fonts.Append(font);
            stylesPart.Stylesheet.Fonts.Append(font1);
            stylesPart.Stylesheet.Fonts.Count = 2;

            stylesPart.Stylesheet.Fonts.AppendChild(new Font());

            // create fills
            stylesPart.Stylesheet.Fills = new Fills();

            // create a solid red fill
            var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
            solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("ff8086") }; // red fill
            solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };
            // create green fill
            var solidGreen = new PatternFill() { PatternType = PatternValues.Solid };
            solidGreen.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("80ff80") }; // green fill
            solidGreen.BackgroundColor = new BackgroundColor { Indexed = 64 };
            //add fill to style sheet
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidGreen });
            stylesPart.Stylesheet.Fills.Count = 4;

            // blank border list
            stylesPart.Stylesheet.Borders = new()
            {
                Count = 1
            };
            stylesPart.Stylesheet.Borders.AppendChild(new Border());

            // blank cell format list
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats
            {
                Count = 1
            };
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            // cell format list
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0, ApplyFill = false }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 3, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            stylesPart.Stylesheet.CellFormats.Count = 4;
            stylesPart.Stylesheet.Save();
            return spreadsheet;
        }
    }
}
