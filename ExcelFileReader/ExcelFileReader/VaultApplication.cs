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
        string Prepaymentreport = "OT.Prepaymentreport";     
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
                //var PrepaymentNo = Convert.ToString( env.Vault.ValueListItemOperations.GetValueListItems(1111));
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
                    // Save the new Excel file
                    string newExcelFilePath = @"D:\M&P_Client\PrePayment_Report\Excle\FilteredData.xlsx";
                    FileInfo newFileInfo = new FileInfo(newExcelFilePath);
                    newPackage.SaveAs(newFileInfo);

                   // string NEWPath = @"D:\M&P_Client\PrePayment_Report\Excle\FilteredData.xlsx";
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

            }

        }
      }
    }
}
                
                




































                // [EventHandler(MFEventHandlerType.MFEventHandlerAfterCreateNewObjectFinalize, ObjectType = "OT.Excelfilereader")]
                // public void Calculation(EventHandlerEnvironment env)
                // {
                //string ExcelFileReaderClassObj = "CL.ExcelFileReader";
                //string CompanyNamePD = "PD.CompanyName";
                //string PrepaymentVoucherNo = "PD.PrepaymentVoucherNumber";

                //string PrePaymentVoucherNo = "PD.PrepaymentNo";
                //string PrepaymentAmount = "PD.PrepaymentAmount";
                //string InvoiceTotalAmount = "PD.InvoiceAmount";
                //string BalanceAmount1 = "PD.BalanceAmount";





                ////string PrepaymentAmountPD = "PD.PrepaymentAmount";
                //string InvoiceTotalAmountPD = "PD.Invoicetotalamount";
                ////string BalanceAmountPD = "PD.BalanceAmount";
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //string excelFilePath = @"D:\FilePath\NewSheet.xlsx";
                //FileInfo fileInfo = new FileInfo(excelFilePath);




                //using (var package = new ExcelPackage(fileInfo))
                //{
                //    // Assuming the data is in the first worksheet, adjust as needed
                //    var worksheet = package.Workbook.Worksheets[0];

                //    // Assuming Company Name is in column A, Prepayment Amount in column B, and Invoice Amount in column C
                //    int rowCount = worksheet.Dimension.Rows;

                //    double ExcelPrepaymentAmount = 0;
                //    double ExcelInvoiceTotalAmount = 0;
                //    string ExcelPrePaymentVoucherNumber = "";
                //    string ExcelCompanyName = "";
                //    for (int row = 2; row <= rowCount; row++) // Start from row 2 to skip the header
                //    {

                //        double invoiceAmount = worksheet.Cells[row, 3].GetValue<double>();
                //        ExcelPrepaymentAmount = worksheet.Cells[row, 2].GetValue<double>();
                //        ExcelPrePaymentVoucherNumber = worksheet.Cells[row, 4].GetValue<string>();
                //        ExcelCompanyName = worksheet.Cells[row, 1].GetValue<string>();
                //        ExcelInvoiceTotalAmount += invoiceAmount;
                //    }

                //    // Calculate the remaining prepayment
                //    double ResultAmount = ExcelPrepaymentAmount - ExcelInvoiceTotalAmount;

                //    ////Now FindOut The Object Of Class Where We Have To Set Above Values
                //    var objID = new MFilesAPI.ObjID();
                //    objID.SetIDs(ObjType: env.ObjVer.Type, ID: env.ObjVer.ID);
                //    var SearchClass = new MFSearchBuilder(env.Vault);
                //    SearchClass.ObjType(ExcelFileReaderClassObj); //Find the object of Prepayment Class
                //    SearchClass.Property(PrepaymentVoucherNo, MFDataType.MFDatatypeText, ExcelPrePaymentVoucherNumber);
                //    SearchClass.Deleted(false);
                //    var ExcelFileReaderClasObj = SearchClass.FindEx(); //Save The Object Of Class In The Variable

                //    //ExcelFileReaderClasObj.CheckOut();
                //    //set Value Of CompanyName into M-File Class
                //    var CompanyName = new MFilesAPI.PropertyValue
                //    {
                //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(CompanyNamePD)
                //    };
                //    CompanyName.Value.SetValue(MFDataType.MFDatatypeInteger, ExcelCompanyName);
                //    ExcelFileReaderClasObj[0].Vault.ObjectPropertyOperations.SetProperty(ObjVer: ExcelFileReaderClasObj[0].ObjVer, PropertyValue: CompanyName);

                //    //set Value Of PrepaymentVoucherNumber into M-File Class
                //    var PrepaymentVoucher = new MFilesAPI.PropertyValue
                //    {
                //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(PrepaymentVoucherNo)
                //    };
                //    PrepaymentVoucher.Value.SetValue(MFDataType.MFDatatypeInteger, ExcelPrePaymentVoucherNumber);
                //    ExcelFileReaderClasObj[1].Vault.ObjectPropertyOperations.SetProperty(ObjVer: ExcelFileReaderClasObj[1].ObjVer, PropertyValue: PrepaymentVoucher);

                //    //set Value Of InvoicAmount into M-File Class
                //    var InvoiceAmount = new MFilesAPI.PropertyValue
                //    {
                //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(InvoiceTotalAmountPD)

                //    };
                //    InvoiceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, ExcelInvoiceTotalAmount);
                //    ExcelFileReaderClasObj[2].Vault.ObjectPropertyOperations.SetProperty(ObjVer: ExcelFileReaderClasObj[2].ObjVer, PropertyValue: InvoiceAmount);

            
    


                //  Declaring Function Toout the value of Required field from ExcelReport
                //    var RequiredPrePaymentVoucher = env.ObjVerEx.GetProperty(PrePaymentVoucherNo).Value.DisplayValue.ToString();
                //var RequiredPrePaymentAmount = Convert.ToDecimal(env.ObjVerEx.GetProperty(PrepaymentAmount).Value.DisplayValue.ToString());
                //var RequiredInvoiceTotalAmount = Convert.ToDecimal(env.ObjVerEx.GetProperty(InvoiceTotalAmount).Value.DisplayValue.ToString());
                //// var RequiredBalanceAmount = env.ObjVerEx.GetProperty(BalanceAmount1).Value.DisplayValue.ToString();
                //var RequiredBalanceAmount = RequiredPrePaymentAmount - RequiredInvoiceTotalAmount;

    ////Now FindOut The Object Of Class Where We Have To Set Above Values
    //var objID = new MFilesAPI.ObjID();
    //objID.SetIDs(ObjType: env.ObjVer.Type, ID: env.ObjVer.ID);
    //var SearchClass = new MFSearchBuilder(env.Vault);
    //SearchClass.ObjType(PrePaymentVoucher); //Find the object of Prepayment Class
    //SearchClass.Property(Prepayment_prepaymentNO, MFDataType.MFDatatypeText, RequiredPrePaymentVoucher);
    //SearchClass.Deleted(false);
    //var PrepaymantClassObj = SearchClass.FindEx(); //Save The Object Of Class In The Variable
    //for (int i = 0; i < PrepaymantClassObj.Count; i++)
    //{
    //    //Set Field Values On PrePaymentVoucher Class
    //    PrepaymantClassObj[i].CheckOut();
    //    var PrePaymentAmount = new MFilesAPI.PropertyValue
    //    {
    //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(PrepaymentAmount)
    //    };
    //    PrePaymentAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredPrePaymentAmount);
    //    PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: PrePaymentAmount);

    //    var InoiceAmount = new MFilesAPI.PropertyValue
    //    {
    //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(InvoiceTotalAmount)
    //    };
    //    InoiceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredInvoiceTotalAmount);
    //    PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: InoiceAmount);

    //    var BalanceAmount = new MFilesAPI.PropertyValue
    //    {
    //        PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(BalanceAmount1)
    //    };
    //    BalanceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredBalanceAmount);
    //    PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: BalanceAmount);
    //    PrepaymantClassObj[i].CheckIn();
    //    break;
    //}

 