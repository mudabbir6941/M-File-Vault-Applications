using MFiles.VAF;
using MFiles.VAF.AppTasks;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;






namespace PrePaymentWorkFlow
{
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        string PrePaymentVoucherNo = "PD.PrepaymentNo";
        string PrepaymentAmount = "PD.PrepaymentAmount";
        string InvoiceTotalAmount = "PD.InvoiceAmount";
        string BalanceAmount1 = "PD.BalanceAmount";

        //PrePayment Class Alias.
        string PrePaymentVoucher = "OT.PrepaymentVoucher";
        string Prepayment_prepaymentNO = "PD.PrepaymentNo";


        [EventHandler(MFEventHandlerType.MFEventHandlerAfterCreateNewObjectFinalize, ObjectType = "OT.Prepaymentreport")]
        public void Calculation(EventHandlerEnvironment env)
        {
            //Declaring Function Toout the value of Required field from ExcelReport
            var RequiredPrePaymentVoucher = env.ObjVerEx.GetProperty(PrePaymentVoucherNo).Value.DisplayValue;
            var RequiredPrePaymentAmount =Convert.ToDecimal(env.ObjVerEx.GetProperty(PrepaymentAmount).Value.DisplayValue.ToString());
            var RequiredInvoiceTotalAmount = Convert.ToDecimal(env.ObjVerEx.GetProperty(InvoiceTotalAmount).Value.DisplayValue.ToString());
            // var RequiredBalanceAmount = env.ObjVerEx.GetProperty(BalanceAmount1).Value.DisplayValue.ToString();
            var RequiredBalanceAmount = RequiredPrePaymentAmount - RequiredInvoiceTotalAmount;

            //Now FindOut The Object Of Class Where We Have To Set Above Values
            var objID = new MFilesAPI.ObjID();
            objID.SetIDs( ObjType: env.ObjVer.Type,ID: env.ObjVer.ID);
            var SearchClass = new MFSearchBuilder(env.Vault);
            SearchClass.ObjType(PrePaymentVoucher); //Find the object of Prepayment Class
            SearchClass.Property(Prepayment_prepaymentNO, MFDataType.MFDatatypeText, RequiredPrePaymentVoucher);
            SearchClass.Deleted(false);
            var PrepaymantClassObj = SearchClass.FindEx(); //Save The Object Of Class In The Variable
            for (int i = 0; i < PrepaymantClassObj.Count; i++)
            {
                //Set Field Values On PrePaymentVoucher Class
                PrepaymantClassObj[i].CheckOut();
                 var PrePaymentAmount = new MFilesAPI.PropertyValue
                {
                    PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(PrepaymentAmount)
                };
                PrePaymentAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredPrePaymentAmount);
                PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: PrePaymentAmount);

                var InoiceAmount = new MFilesAPI.PropertyValue
                {
                    PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(InvoiceTotalAmount)
                };
                InoiceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredInvoiceTotalAmount);
                PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: InoiceAmount);

                var BalanceAmount = new MFilesAPI.PropertyValue
                {
                    PropertyDef = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias(BalanceAmount1)
                };
                BalanceAmount.Value.SetValue(MFDataType.MFDatatypeInteger, RequiredBalanceAmount);
                PrepaymantClassObj[i].Vault.ObjectPropertyOperations.SetProperty(ObjVer: PrepaymantClassObj[i].ObjVer, PropertyValue: BalanceAmount);
                PrepaymantClassObj[i].CheckIn();
                break;
            }
        }

    }
}
