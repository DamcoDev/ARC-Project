using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace CasesToCustomer
{
    public class Cases:CodeActivity
    {
        [Input("Customer")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> Cust { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1) { return; }
            Guid CustomerGUID = Cust.Get<EntityReference>(context).Id;
            executeBusinessLogic(CustomerGUID, service);
        }

        private void executeBusinessLogic(Guid customerGUID, IOrganizationService service)
        {
            try
            {
                string customerId = customerGUID.ToString();//context.InputParameters["CustomerId"].ToString();
                List<Entity> customerContacts = null;
                customerContacts = GetCustomerContacts(customerId, service);

                string emailTemplateId = "37318889-1b50-ec11-8c62-000d3aba9b1b";
                EmailTemplate emailTemplate = GetEmailTemplate(emailTemplateId, service);
                if (emailTemplate != null)
                {
                    bool areEmailsSent = SendEmailToCustomers(customerId, customerContacts, emailTemplate, service);
                    //Trace.AppendLine("areEmailsSent = " + areEmailsSent.ToString());
                    //Trace.AppendLine("emailTemplate != null");
                    //context.OutputParameters["Message"] = "Success";
                }
                else
                {
                    //Trace.AppendLine("Email template not found");
                    //context.OutputParameters["Message"] = "Failure";
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
        public class EmailTemplate
        {
            public Guid EmailTemplateId { get; set; }
            public string Subject { get; set; }
            public string Body { get; set; }
            public EmailTemplate()
            {
                EmailTemplateId = Guid.Empty;
                Subject = string.Empty;
                Body = string.Empty;
            }
        }
        public class Customer
        {
            public Guid CustomerId { get; set; }
            public string CustomerName { get; set; }
            public string CompanyName { get; set; }

            public Customer()
            {
                CustomerId = Guid.Empty;
                CustomerName = string.Empty;
                CompanyName = string.Empty;
            }
        }
        public class CustomerExcel
        {
            public IOrganizationService CrmService { get; set; }
            public Customer CustomerObj { get; set; }
            public Guid LoggedInUserId { get; set; }
            //public string FileNameWithPath { get; set; }
            public string FileName { get; set; }
            public StringBuilder Trace { get; set; }

            public CustomerExcel(IOrganizationService crmService, Customer customerObj)
            {
                //Trace = trace;
                CrmService = crmService;
                //LoggedInUserId = loggedInUserId;
                CustomerObj = customerObj;
                this.FileName = CustomerObj.CompanyName + "_" + DateTime.Now.ToFileTime() + "_Accounts.xlsx";
                //FileNameWithPath = CreateFilePath();
                //Trace.AppendLine("CE=>...0...FileNameWithPath = " + FileNameWithPath);
            }

            public byte[] GetExcelFileInBytesArray()
            {
                Trace.AppendLine("CE=>1...Inside GetExcelFileInBytesArray()");
                CreateExcelDoc();
                return null;
            }

            public byte[] CreateExcelDoc()
            {
                MemoryStream memoryStream = new MemoryStream();
                //Trace.AppendLine("CE=>---2...Inside CreateExcelDoc()");
                //FileNameWithPath = @"C:\_Temp\UNI STYLE TAILORING BR_131596023494075050_Accounts.xlsx";
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                    //// add styles to sheet                
                    Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                    wbsp.Stylesheet = stylesheet1;//CreateStylesheet();
                    wbsp.Stylesheet.Save();

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Accounts" };

                    sheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    sheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                    //sheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();

                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    Columns _cols = new Columns(); // Created to allow bespoke width columns
                    CreateColumns(document, worksheetPart.Worksheet, _cols);
                    AddHeader(worksheetPart.Worksheet, new List<string>() { CustomerObj.CompanyName, string.Empty, string.Empty, string.Empty });
                    worksheetPart = MergeRowCells(worksheetPart);
                    AddHeader(worksheetPart.Worksheet, new List<string>() { "Account Number", "Product Name", "Fax", "Main Phone" });

                    // Inserting each employee
                    UInt32 currentRowNumber = Convert.ToUInt32(sheetData.ChildElements.Count()) + 1;
                    int accountNumber = 123456;
                    int fax = 92321555;
                    int phone = 92321666;

                    for (int i = 0; i < 5; i++)
                    {
                        //var localDate = DateTime.UtcNow.ToLocalTime();
                        //var localDateTime = DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified);
                        //string fromDate = ConvertCrmDateToUserDate(localDateTime).ToString();
                        //string toDate = ConvertCrmDateToUserDate(localDateTime.AddDays(2)).ToString();

                        string productName = "Product " + (i + 1).ToString();
                        accountNumber++;
                        Row dataRow = CreateExcelRow(accountNumber.ToString(), productName, fax.ToString(), phone.ToString(), currentRowNumber, false);
                        sheetData.AppendChild(dataRow);
                        currentRowNumber++;
                    }

                    wbsp.Stylesheet = CreateStylesheet_New(); //CreateStylesheet();
                    wbsp.Stylesheet.Save();

                    int countRows = sheetData.Elements<Row>().Count();

                    if (countRows >= 3)
                    {
                        Row row_1 = sheetData.Elements<Row>().ElementAt<Row>(0);
                        int index0 = 0;
                        foreach (Cell c in row_1.Elements<Cell>())
                        {
                            c.StyleIndex = Convert.ToUInt32(index0);
                        }

                        Row row_2 = sheetData.Elements<Row>().ElementAt<Row>(1);
                        int index1 = 1;
                        foreach (Cell c in row_2.Elements<Cell>())
                        {
                            c.StyleIndex = Convert.ToUInt32(index1);
                        }

                        for (int i = 0; i < sheetData.Elements<Row>().Count(); i++)
                        {
                            Row dataRow = sheetData.Elements<Row>().ElementAt(i);
                            if (i > 1)
                            {
                                int index2 = 2;
                                foreach (Cell c in dataRow.Elements<Cell>())
                                {
                                    c.StyleIndex = Convert.ToUInt32(index2);
                                }
                            }
                        }
                    }

                    worksheetPart.Worksheet.Save();
                }
                // Trace.AppendLine("CE=>---3...Executed method CreateExcelDoc()");
                return memoryStream.ToArray();
            }

            private Columns GenerateColumnsData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)
            {
                Columns columns = new Columns();

                for (uint index = StartColumnIndex; index <= EndColumnIndex; index++)
                {
                    Column column = new Column();

                    column.Min = index;
                    column.Max = index;
                    column.Width = ColumnWidth;
                    column.CustomWidth = true;

                    columns.Append(column);
                }

                return columns;
            }

            public WorksheetPart MergeRowCells(WorksheetPart worksheetPart)
            {
                MergeCells mergeCells = new MergeCells();
                //append a MergeCell to the mergeCells for each set of merged cells
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:D1") });
                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                return worksheetPart;
            }

            public Row CreateExcelRow(string column1Value, string column2Value, string column3Value, string column4Value, UInt32 rowNumber, bool isFirstRow)
            {
                Row row1 = new Row();
                if (isFirstRow == true)
                {
                    row1.CustomHeight = true;
                    row1.Height = 30;
                }

                ///---Row 1
                Cell row1_CellA1 = new Cell();
                row1_CellA1.CellReference = "A" + rowNumber.ToString();
                row1_CellA1.CellValue = new CellValue(column1Value);
                row1_CellA1.DataType = CellValues.String;
                row1.InsertAt(row1_CellA1, 0);

                ///---Row 2
                Cell row1_CellB1 = new Cell();
                row1_CellB1.CellReference = "B" + rowNumber.ToString();
                row1_CellB1.CellValue = new CellValue(column2Value);
                row1_CellB1.DataType = CellValues.String;
                row1.InsertAt(row1_CellB1, 1);

                ///---Row 3
                Cell row1_CellC1 = new Cell();
                row1_CellC1.CellReference = "C" + rowNumber.ToString();
                row1_CellC1.CellValue = new CellValue(column3Value);
                row1_CellC1.DataType = CellValues.String;
                row1.InsertAt(row1_CellC1, 2);

                Cell row1_CellD1 = new Cell();
                row1_CellD1.CellReference = "D" + rowNumber.ToString();
                row1_CellD1.CellValue = new CellValue(column4Value);
                row1_CellD1.DataType = CellValues.String;
                row1.InsertAt(row1_CellD1, 3);

                return row1;
            }

            public DateTime ConvertCrmDateToUserDate(DateTime inputDate)
            {
                //replace userid with id of user
                Entity userSettings = CrmService.Retrieve("usersettings", LoggedInUserId, new ColumnSet("timezonecode"));
                //timezonecode for UTC is 85
                int timeZoneCode = 85;

                //retrieving timezonecode from usersetting
                if ((userSettings != null) && (userSettings["timezonecode"] != null))
                {
                    timeZoneCode = (int)userSettings["timezonecode"];
                }
                //retrieving standard name
                var qe = new QueryExpression("timezonedefinition");
                qe.ColumnSet = new ColumnSet("standardname");
                qe.Criteria.AddCondition("timezonecode", ConditionOperator.Equal, timeZoneCode);
                EntityCollection TimeZoneDef = CrmService.RetrieveMultiple(qe);

                TimeZoneInfo userTimeZone = null;
                if (TimeZoneDef.Entities.Count == 1)
                {
                    String timezonename = TimeZoneDef.Entities[0]["standardname"].ToString();
                    userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timezonename);
                }
                //converting date from UTC to user time zone
                DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(inputDate, userTimeZone);

                return cstTime;
            }

            public byte[] ReadFile()
            {
                //FileStream fs = new FileStream(FileNameWithPath, FileMode.Open, FileAccess.Read);
                // int length = Convert.ToInt32(fs.Length);
                //byte[] data = new byte[length];
                //fs.Read(data, 0, length);
                //fs.Close();
                //return data;
                return null;
            }

            public string CreateFilePath()
            {
                string result = string.Empty;
                try
                {
                    //result = Path.GetTempPath();
                    //result = Path.GetTempFileName();
                    result = @"C:\\_Temp\\";
                }
                catch (Exception ex)
                {
                    Trace.AppendLine("CE=>---3...Exception = " + ex.InnerException);
                }

                //if(!string.IsNullOrEmpty(result))
                //{
                //    Trace.AppendLine("CE=>---3...CreateFilePath()...result = " + result);
                //}

                //if (!string.IsNullOrEmpty(CustomerObj.CompanyName))
                //{
                //    Trace.AppendLine("CE=>---4...CreateFilePath()...CustomerObj.CompanyName = " + CustomerObj.CompanyName);
                //}

                FileName = CustomerObj.CompanyName + "_" + DateTime.Now.ToFileTime() + "_Accounts.xlsx";
                //string fileNameWithPath = result;
                string fileNameWithPath = result + FileName;
                return fileNameWithPath;
            }

            public void CreateColumns(SpreadsheetDocument spreadSheet, Worksheet currentWorkSheet, Columns _cols)
            {
                CreateColumnWidth(spreadSheet, currentWorkSheet, _cols, 1, 1, 30);
                CreateColumnWidth(spreadSheet, currentWorkSheet, _cols, 2, 2, 30);
                CreateColumnWidth(spreadSheet, currentWorkSheet, _cols, 3, 3, 30);
                CreateColumnWidth(spreadSheet, currentWorkSheet, _cols, 4, 4, 30);
            }

            /// <summary>
            /// add the bespoke columns for the list spreadsheet
            /// </summary>
            public void CreateColumnWidth(SpreadsheetDocument spreadSheet, Worksheet currentWorkSheet, Columns _cols, uint startIndex, uint endIndex, double width)
            {
                // Find the columns in the worksheet and remove them all
                if (currentWorkSheet.Where(x => x.LocalName == "cols").Count() > 0)
                    currentWorkSheet.RemoveChild<Columns>(_cols);

                // Create the column
                Column column = new Column();
                column.Min = startIndex;
                column.Max = endIndex;
                column.Width = width;
                column.CustomWidth = true;
                _cols.Append(column); // Add it to the list of columns

                // Make sure that the column info is inserted *before* the sheetdata
                currentWorkSheet.InsertBefore<Columns>(_cols, currentWorkSheet.Where(x => x.LocalName == "sheetData").First());
                currentWorkSheet.Save();
                spreadSheet.WorkbookPart.Workbook.Save();

            }

            public void AddHeader(Worksheet currentWorkSheet, List<string> headers)
            {
                // Find the sheetdata of the worksheet
                SheetData sd = (SheetData)currentWorkSheet.Where(x => x.LocalName == "sheetData").First();
                Row header = new Row();
                // increment the row index to the next row
                header.RowIndex = Convert.ToUInt32(sd.ChildElements.Count()) + 1;
                sd.Append(header); // Add the header row

                foreach (string heading in headers)
                {
                    AppendCell(header, header.RowIndex, heading, 1);
                }

                // save worksheet
                currentWorkSheet.Save();

            }

            /// <summary>
            /// Add cell into the passed row.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="rowIndex"></param>
            /// <param name="value"></param>
            /// <param name="styleIndex"></param>
            private void AppendCell(Row row, uint rowIndex, string value, uint styleIndex)
            {
                Cell cell = new Cell();
                cell.DataType = CellValues.InlineString;
                //cell.StyleIndex = styleIndex;  // Style index comes from stylesheet generated in GenerateStyleSheet()
                Text t = new Text();
                t.Text = value;

                // Append Text to InlineString object
                InlineString inlineString = new InlineString();
                inlineString.AppendChild(t);

                // Append InlineString to Cell
                cell.AppendChild(inlineString);

                // Get the last cell's column
                string nextCol = "A";
                Cell c = (Cell)row.LastChild;
                if (c != null) // if there are some cells already there...
                {
                    int numIndex = c.CellReference.ToString().IndexOfAny(new char[] { '1', '2', '3', '4' });

                    // Get the last column reference
                    string lastCol = c.CellReference.ToString().Substring(0, numIndex);
                    // Increment
                    nextCol = IncrementColRef(lastCol);
                }

                cell.CellReference = nextCol + rowIndex;

                row.AppendChild(cell);
            }

            // Increment the column reference in an Excel fashion, i.e. A, B, C...Z, AA, AB etc.
            // Partly stolen from somewhere on the Net and modified for my use.
            private string IncrementColRef(string lastRef)
            {
                char[] characters = lastRef.ToUpperInvariant().ToCharArray();
                int sum = 0;
                for (int i = 0; i < characters.Length; i++)
                {
                    sum *= 26;
                    sum += (characters[i] - 'A' + 1);
                }

                sum++;

                string columnName = String.Empty;
                int modulo;

                while (sum > 0)
                {
                    modulo = (sum - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    sum = (int)((sum - modulo) / 26);
                }

                return columnName;
            }

            private static Stylesheet CreateStylesheet_New()
            {
                Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                Fonts fonts = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

                Font font1 = new Font();
                Bold bold = new Bold();
                bold.Val = true;
                DocumentFormat.OpenXml.Spreadsheet.FontSize fontSize1 = new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 15D };
                //Color color1 = new Color() { Theme = (UInt32Value)1U };
                Color color1 = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }; //-- 000000 = Black
                FontName fontName1 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 0 };
                FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

                font1.Append(bold);
                font1.Append(fontSize1);
                font1.Append(color1);
                font1.Append(fontName1);
                font1.Append(fontFamilyNumbering1);
                font1.Append(fontScheme1);

                fonts.Append(font1);

                Font font2 = new Font();
                DocumentFormat.OpenXml.Spreadsheet.FontSize fontSize2 = new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11D };
                //Color color2 = new Color() { Theme = (UInt32Value)2U };
                Color color2 = new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } }; //---2F4F4F = Dark Slate Gray
                                                                                                //BackgroundColor backColor2 = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } };
                FontName fontName2 = new FontName() { Val = "Calibri" };
                //FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 1 };
                //FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

                font2.Append(fontSize2);
                font2.Append(color2);
                font2.Append(fontName2);

                fonts.Append(font2);

                Font font3 = new Font();
                DocumentFormat.OpenXml.Spreadsheet.FontSize fontSize3 = new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11D };
                //Color color2 = new Color() { Theme = (UInt32Value)2U };
                Color color3 = new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }; //---696969 = Gray
                FontName fontName3 = new FontName() { Val = "Calibri" };
                FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
                FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

                font3.Append(fontSize3);
                font3.Append(color3);
                font3.Append(fontName3);
                font3.Append(fontFamilyNumbering3);
                font3.Append(fontScheme3);

                fonts.Append(font3);

                Fills fills = new Fills(
                        new Fill(                                                           // Index 0 - The default fill.
                            new PatternFill() { PatternType = PatternValues.None }),
                        new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                            new PatternFill() { PatternType = PatternValues.Gray125 }),
                        new Fill(                                                           // Index 2 - The yellow fill.
                            new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "808080" } }
                            )
                            { PatternType = PatternValues.Solid })
                            );

                //Fills fills = new Fills() { Count = (UInt32Value)1U };

                // FillId = 0
                //Fill fill1 = new Fill();
                //PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
                //fill1.Append(patternFill1);

                // FillId = 1
                Fill fill2 = new Fill();
                //PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
                PatternFill patternFill2 = new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "1E90FF" } }) { PatternType = PatternValues.Solid };
                //ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } };//A9A9A9
                //ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFAF0" };//---696969 = Gray
                //BackgroundColor backgroundColor2 = new BackgroundColor() { Rgb = "2F4F4F" };//---FFFAF0 = White
                //BackgroundColor backgroundColor2 = new BackgroundColor() { Rgb = "000000" };//---FFFAF0 = White
                //BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)0U };
                //BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                //patternFill2.Append(foregroundColor2);
                //patternFill2.Append(backgroundColor2);
                fill2.Append(patternFill2);



                //// FillId = 2,RED
                //Fill fill3 = new Fill();
                //PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
                ////ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFF0000" };//A9A9A9
                //ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "A9A9A900" };//A9A9A9
                //BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                //patternFill3.Append(foregroundColor1);
                //patternFill3.Append(backgroundColor1);
                //fill3.Append(patternFill3);

                //// FillId = 3,BLUE
                //Fill fill4 = new Fill();
                //PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
                //ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FF0070C0" };
                //BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                //patternFill4.Append(foregroundColor4);
                //patternFill4.Append(backgroundColor4);
                //fill4.Append(patternFill4);

                //// FillId = 4,YELLO
                //Fill fill5 = new Fill();
                //PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
                //ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFFF00" };
                //BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                //patternFill5.Append(foregroundColor3);
                //patternFill5.Append(backgroundColor3);
                //fill5.Append(patternFill5);

                //fills.Append(fill1);
                fills.Append(fill2);
                //fills1.Append(fill3);
                //fills1.Append(fill4);
                //fills1.Append(fill5);

                Borders borders1 = new Borders() { Count = (UInt32Value)1U };

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

                borders1.Append(border1);

                //CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
                //CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
                //cellStyleFormats1.Append(cellFormat1);

                CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)3U };
                //CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
                //CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                //CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                //CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };

                CellFormat cellFormat1 = new CellFormat() { FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
                CellFormat cellFormat2 = new CellFormat() { FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, ApplyFill = true, ApplyFont = true };
                CellFormat cellFormat3 = new CellFormat() { FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true };
                //CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };

                cellFormats.Append(cellFormat1);
                cellFormats.Append(cellFormat2);
                cellFormats.Append(cellFormat3);
                //cellFormats1.Append(cellFormat5);

                //CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
                //CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
                //cellStyles1.Append(cellStyle1);

                //DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
                //TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

                StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

                StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
                stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

                stylesheetExtensionList1.Append(stylesheetExtension1);

                stylesheet1.Append(fonts);
                stylesheet1.Append(fills);
                stylesheet1.Append(borders1);
                stylesheet1.Append(cellFormats);
                stylesheet1.Append(stylesheetExtensionList1);
                return stylesheet1;
            }
        }

        public EmailTemplate GetEmailTemplate(string emailTemplateId, IOrganizationService crmService)
        {
            //Outage Alert Email Notification Template
            EmailTemplate emailTemplate = null;
            string fetchXmlString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='template'>
                                            <attribute name='subject' />
                                            <attribute name='body' />
                                            <filter type='and'>
                                              <condition attribute='templateid' operator='eq' uitype='template' value='{" + emailTemplateId + @"}' />
                                            </filter>
                                          </entity>
                                        </fetch>";
            EntityCollection entColl = crmService.RetrieveMultiple(new FetchExpression(fetchXmlString));
            List<Entity> entitiesList = GetEntitiesList(entColl);
            if (entitiesList != null && entitiesList.Count > 0)
            {
                emailTemplate = new EmailTemplate();
                emailTemplate.EmailTemplateId = entitiesList[0].Id;
                emailTemplate.Subject = GetDataFromXml(entitiesList[0]["subject"].ToString());
                emailTemplate.Body = GetDataFromXml(entitiesList[0]["body"].ToString());
            }

            return emailTemplate;
        }

        private string GetDataFromXml(string value)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            XDocument document = XDocument.Parse(value);
            XNode node = document.Descendants().Elements().Last().LastNode;
            string nodeText = node.ToString().Replace("<![CDATA[", string.Empty).Replace("]]>", string.Empty);
            return nodeText;
        }

        public List<Entity> GetCustomerContacts(string customerId, IOrganizationService crmService)
        {
            List<Entity> customerContacts = new List<Entity>();
            string fetchXmlString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='contact'>
                                                <attribute name='fullname' />
                                                <filter type='and'>
                                                  <condition attribute='parentcustomerid' operator='eq' uitype='account' value='{" + customerId + @"}' />
                                                </filter>
                                              </entity>
                                            </fetch>";

            EntityCollection entColl = crmService.RetrieveMultiple(new FetchExpression(fetchXmlString));
            List<Entity> entitiesList = GetEntitiesList(entColl);
            if (entitiesList.Count > 0)
            {
                customerContacts.AddRange(entitiesList);
                //Trace.AppendLine("2.GetCustomerContacts..." + entitiesList.Count.ToString() + "...");
            }
            return customerContacts;
        }

        public List<Entity> GetEntitiesList(EntityCollection entColl)
        {
            List<Entity> entities = new List<Entity>();
            if (entColl != null && entColl.Entities != null && entColl.Entities.Count > 0)
            {
                entities.AddRange(entColl.Entities);
            }
            return entities;
        }

        public Entity[] GetToListOfCustomerContacts(List<Entity> customerContacts)
        {
            EntityCollection to = new EntityCollection();
            for (int j = 0; j < customerContacts.Count; j++)
            {
                Entity contact = customerContacts[j];
                if (contact.Id != Guid.Empty)
                {
                    //EntityReference toEntRefParty = new EntityReference("contact", authorizedContact.Id);
                    Entity toParty = new Entity("activityparty");
                    toParty["partyid"] = new EntityReference("contact", contact.Id);
                    to.Entities.Add(toParty);
                }
            }

            Entity[] arrayTo = to.Entities.ToArray();

            return arrayTo;
        }

        public bool SendEmailToCustomers(string customerId, List<Entity> customerContacts, EmailTemplate emailTemplate, IOrganizationService crmService)
        {
            bool areEmailsSent = false;

            if (emailTemplate == null || emailTemplate.EmailTemplateId == Guid.Empty || string.IsNullOrEmpty(emailTemplate.Body))
            {
                //Trace.AppendLine("4.EmailTemplate is missing in CRM or body text is not there in email template.");
            }
            else
            {
                EntityReference entRefCustomer = new EntityReference("account", new Guid(customerId));
                //Trace.AppendLine("4.EmailTemplate found.");
                if (customerContacts.Count > 0)
                {

                    //  Trace.AppendLine("4.customers[i].OutageAlert is found");
                    //--Get customer's accounts in byte array for email attachment
                    //byte[] fileStream = GetCustomersAccountsStream(crmService, loggedInUserId, customers[i]);
                    //Trace.AppendLine("5.fileStream..." + fileStream.Length.ToString() + "...");

                    Entity fromPartyGuid = new Entity("activityparty");
                    fromPartyGuid["partyid"] = new EntityReference("queue", new Guid("43786172-c2ac-ea11-a812-000d3ab19dd4"));

                    Entity[] fromParty = new[] { fromPartyGuid };
                    Entity[] to = GetToListOfCustomerContacts(customerContacts);

                    if (to.Length > 0)
                    {
                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = string.Format(emailTemplate.Subject);
                        var localDate = DateTime.UtcNow.ToLocalTime();
                        var localDateTime = DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified);
                        //string fromDate = ConvertCrmDateToUserDate(crmService, loggedInUserId, localDateTime).ToString();
                        //string toDate = ConvertCrmDateToUserDate(crmService, loggedInUserId, localDateTime.AddDays(2)).ToString();
                        entEmail["description"] = emailTemplate.Body;
                        entEmail["to"] = to;
                        if (fromParty != null)
                        {
                            entEmail["from"] = fromParty;

                        }
                        entEmail["regardingobjectid"] = entRefCustomer;
                        Guid emailId = crmService.Create(entEmail);
                        // Trace.AppendLine("emailId = " + emailId.ToString() + "...");


                        //---Create attachment of email
                        Customer customer = new Customer();
                        customer.CompanyName = "Company 123";
                        CustomerExcel customerExcel = new CustomerExcel(crmService, customer);
                        //customerExcel.CreateExcelDoc_Test();
                        byte[] fileStream = customerExcel.CreateExcelDoc();
                        //Trace.AppendLine("1122----fileStream = " + fileStream.Length.ToString());
                        CreateExcelFileAttachmentForEmail(customerExcel.FileName, fileStream, emailId, crmService);

                        //--- Commenting temporarily this line of code as email sending is not enabled on current crm org. Due to this it is throwing error
                        SendEmailResponse response = SendEmail(emailId, crmService);
                        areEmailsSent = true;

                    }
                    else
                    {
                        //Trace.AppendLine("4.to.Length is ZEROOOOOOOOO");
                    }

                }
                else
                {
                    //Trace.AppendLine("4.While sending email, either object OutageAlert is null or OutageAlertId is empty GUID.");
                }


            }
            return areEmailsSent;
        }

        public void CreateExcelFileAttachmentForEmail(string fileName, byte[] fileStream, Guid emailId, IOrganizationService crmService)
        {
            //---Create attachment of email
            Entity attachment = new Entity("activitymimeattachment");
            attachment["subject"] = "Customer Accounts";
            attachment["filename"] = fileName;
            //byte[] fileStream = new byte[] { }; //Set file stream bytes
            attachment["body"] = Convert.ToBase64String(fileStream);
            attachment["mimetype"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";//"text/csv";
            attachment["attachmentnumber"] = 1;
            attachment["objectid"] = new EntityReference("email", emailId);
            attachment["objecttypecode"] = "email";
            crmService.Create(attachment);
        }

        public SendEmailResponse SendEmail(Guid emailId, IOrganizationService crmService)
        {
            SendEmailResponse response = null;
            SendEmailRequest sendRequest = new SendEmailRequest();
            sendRequest.EmailId = emailId;
            sendRequest.TrackingToken = "";
            sendRequest.IssueSend = true;
            try
            {
                response = (SendEmailResponse)crmService.Execute(sendRequest);
            }
            catch (Exception ex)
            {
                ;
            }
            return response;
        }

        public DateTime ConvertCrmDateToUserDate(IOrganizationService service, Guid userGuid, DateTime inputDate)
        {
            //replace userid with id of user
            Entity userSettings = service.Retrieve("usersettings", userGuid, new ColumnSet("timezonecode"));
            //timezonecode for UTC is 85
            int timeZoneCode = 85;

            //retrieving timezonecode from usersetting
            if ((userSettings != null) && (userSettings["timezonecode"] != null))
            {
                timeZoneCode = (int)userSettings["timezonecode"];
            }
            //retrieving standard name
            var qe = new QueryExpression("timezonedefinition");
            qe.ColumnSet = new ColumnSet("standardname");
            qe.Criteria.AddCondition("timezonecode", ConditionOperator.Equal, timeZoneCode);
            EntityCollection TimeZoneDef = service.RetrieveMultiple(qe);

            TimeZoneInfo userTimeZone = null;
            if (TimeZoneDef.Entities.Count == 1)
            {
                String timezonename = TimeZoneDef.Entities[0]["standardname"].ToString();
                userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timezonename);
            }
            //converting date from UTC to user time zone
            DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(inputDate, userTimeZone);

            return cstTime;
        }
    }
}
