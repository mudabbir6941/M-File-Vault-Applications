using MFiles.VAF;
using MFiles.VAF.AppTasks;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;
using System;
using System.Diagnostics;
using OfficeOpenXml;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace ExcelFileReader
{
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        double ResultAmount;
        double excelrows;
   
        [StateAction("WFS.PrepaymentWorkfow.UploadFile")]
        public void WorkflowStateAction(StateEnvironment env)
        {
            
            string PrepaymentNo = "PD.PrepaymentNo";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelFilePath = @"D:\M&P_Client\FilePath\NewSheet.xlsx";
            FileInfo fileInfo = new FileInfo(excelFilePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                // Assuming the data is in the first worksheet, adjust as needed
                var worksheet = package.Workbook.Worksheets[0];
                // Assuming Company Name is in column A, Prepayment Amount in column B, and Invoice Amount in column C
                int rowCount = worksheet.Dimension.Rows;
                string ExcelPrepaymentNo = "";
                double ExcelPrepaymentAmount = 0;
                double ExcelInvoiceTotalAmount = 0;
                ValueListItemClass InvoiceList = new ValueListItemClass();
                ValueListItemClass MFInvoiceList = new ValueListItemClass();
                for (int row = 2; row <= rowCount; row++)
                {
                    ExcelPrepaymentNo = worksheet.Cells[row, 2].GetValue<string>();
                    ExcelPrepaymentAmount = worksheet.Cells[row, 4].GetValue<double>();
                    double invoiceAmount = worksheet.Cells[row, 9].GetValue<double>();                  
                    ExcelInvoiceTotalAmount += invoiceAmount;
                    InvoiceList.Name = worksheet.Cells[row, 8].GetValue<string>();
                    string InvoiceDate = worksheet.Cells[row, 10].GetValue<string>();

                    // Check if the value already exists in the Value List
                    bool valueExists = CheckIfValueExistsInValueList(InvoiceList.Name);
                    if (!valueExists)
                    {
                        // Value doesn't exist, so add it to the Value List
                        env.Vault.ValueListItemOperations.AddValueListItem(121, InvoiceList, true);
                    }                   

                }

                // Function to check if a value exists in the Value List
                bool CheckIfValueExistsInValueList(string valueToCheck)
                {
                    var temp = false;
                    var valuelistitems = env.Vault.ValueListItemOperations.GetValueListItems(121);
                    foreach (ValueListItem item in valuelistitems)
                    {

                        var name = item.Name;
                        if (name == valueToCheck)
                        {
                            temp = true;
                            break;
                        }
                    }
                    return temp;
                }             
                var PrepaymentNo2 = env.ObjVerEx.GetProperty(PrepaymentNo).Value.DisplayValue;
                foreach (var item in ExcelPrepaymentNo)
                {
                    if (PrepaymentNo2 == ExcelPrepaymentNo)
                    {
                        // Calculate the remaining prepayment according to voucher number
                        ResultAmount = ExcelPrepaymentAmount - ExcelInvoiceTotalAmount;
                    }

                }

                //set InvoiceTotal Amount Into Metadata
                var objID = new MFilesAPI.ObjID();
                objID.SetIDs(env.ObjVer.Type, env.ObjVer.ID);
                // Create a property value to update.
                var InvoiceTotalAmount = new MFilesAPI.PropertyValue
                {
                    PropertyDef = 1118
                };
                InvoiceTotalAmount.Value.SetValue(MFDataType.MFDatatypeInteger, ExcelInvoiceTotalAmount);// This must be correct for the property definition.

                // Update the property on the server.
                env.Vault.ObjectPropertyOperations.SetProperty(ObjVer: env.ObjVer, PropertyValue: InvoiceTotalAmount);

                //set Balance Amount Into Metadata
                objID.SetIDs(env.ObjVer.Type, env.ObjVer.ID);
                // Create a property value to update.
                var BalanceAmount = new MFilesAPI.PropertyValue
                {
                    PropertyDef = 1120
                };
                BalanceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, ResultAmount);// This must be correct for the property definition.

                // Update the property on the server.
                env.Vault.ObjectPropertyOperations.SetProperty(ObjVer: env.ObjVer, PropertyValue: BalanceAmount);



                //Save New Excel Sheet
                using (var newPackage = new ExcelPackage())
                {
                    var newWorksheet = newPackage.Workbook.Worksheets.Add("FilteredData");
                    // Create headers for the new worksheet
                    newWorksheet.Cells[1, 1].Value = "PrepaymentNo";
                    newWorksheet.Cells[1, 2].Value = "PrepaymentAmount";
                    newWorksheet.Cells[1, 3].Value = "Invoices No";
                    newWorksheet.Cells[1, 4].Value = "InvoiceAmount";
                    newWorksheet.Cells[1, 5].Value = "InvoicDate";

                    int newRow = 2;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        //ExcelPrepaymentNo = worksheet.Cells[row, 2].GetValue<string>();
                        //ExcelPrepaymentAmount = worksheet.Cells[row, 4].GetValue<double>();
                        //double invoiceAmount2 = worksheet.Cells[row, 9].GetValue<double>();
                        InvoiceList.Name = worksheet.Cells[row, 8].GetValue<string>();                    
                        double invoiceAmount = worksheet.Cells[row, 9].GetValue<double>();
                        string InvoiceDate = worksheet.Cells[row, 10].GetValue<string>();
                        // Add the filtered data to the new worksheet
                        newWorksheet.Cells[newRow, 1].Value = ExcelPrepaymentNo;
                        newWorksheet.Cells[newRow, 2].Value = ExcelPrepaymentAmount;
                        newWorksheet.Cells[newRow, 3].Value = InvoiceList.Name;
                        newWorksheet.Cells[newRow, 4].Value = invoiceAmount;
                        newWorksheet.Cells[newRow, 5].Value = InvoiceDate;

                        newRow++;

                    }
                    excelrows = worksheet.Dimension.Rows;
                    // Save the new Excel file
                    string newExcelFilePath = @"D:\M&P_Client\PrePayment_Report\Excle\FilteredData.xlsx";
                    FileInfo newFileInfo = new FileInfo(newExcelFilePath);
                    newPackage.SaveAs(newFileInfo);
                    string pdfFilePath = @"D:\M&P_Client\PrePayment_Report\report\PDFFilteredData.pdf";

                    //    //Convert Excel to PDF using iTextSharp

                    ConvertExcelToPdf(newExcelFilePath, pdfFilePath);

                    //    // Function to convert Excel to PDF using iTextSharp
                    void ConvertExcelToPdf(string a, string b)
                    {
                        using (var excelPackage = new ExcelPackage(new FileInfo(newExcelFilePath)))
                        {
                            
                                var workbook = excelPackage.Workbook;
                            var worksheet1 = workbook.Worksheets[0]; // Assuming you have one worksheet

                                using (var stream = new FileStream(pdfFilePath, FileMode.Create))
                                {
                                    var document = new Document();
                                    PdfWriter writer = PdfWriter.GetInstance(document, stream);
                                    document.Open();

                                    var pdfTable = new PdfPTable(worksheet1.Dimension.Columns);
                                    pdfTable.DefaultCell.Padding = 2;
                                    pdfTable.WidthPercentage = 100;
                                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                                    pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;


                                Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10); // You can adjust the font size as needed
                               
                                Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 8); // Adjust the font size as needed


                                foreach (var cell in worksheet1.Cells)
                                    {
                                        if (cell.Start.Row == 1)
                                        {
                                            var phrase = new Phrase(cell.Text, boldFont);
                                            pdfTable.AddCell(phrase);
                                        }
                                        else
                                        {
                                            var phrase2 = new Phrase(cell.Text, normalFont);
                                            pdfTable.AddCell(phrase2);
                                        }
                                    }

                                    document.Add(pdfTable);
                                    document.Close();
                                }                                
                        }
                    }
                    var objID2 = new MFilesAPI.ObjID();
                    objID2.SetIDs(env.ObjVer.Type, env.ObjVer.ID);
                    var objFileNew = env.Vault.ObjectFileOperations.GetFilesForModificationInEventHandler(env.ObjVer);
                   // string uniqueNumber = Guid.NewGuid().ToString();
                    env.Vault.ObjectFileOperations.AddFile(
                    ObjVer: env.ObjVer,
                    Title: excelrows+"-Report-" +ExcelPrepaymentNo,
                    Extension: "pdf",
                    SourcePath: @"D:\M&P_Client\PrePayment_Report\report\PDFFilteredData.pdf");               
                    if (File.Exists(pdfFilePath) || File.Exists(newExcelFilePath))
                    {
                        File.Delete(pdfFilePath);
                        File.Delete(newExcelFilePath);
                    }
                }
            }
        }
    }
}
                
                





 