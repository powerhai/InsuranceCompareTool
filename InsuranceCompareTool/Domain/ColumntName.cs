using System.ComponentModel;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Domain {
    public enum LogicColumnName
    {
        [DescriptionAttribute(BillSheetColumns.CURRENT_SERVICE_ID)]
        ServiceID,

        [DescriptionAttribute("客服在职状态")]
        ServiceStatus,
        [DescriptionAttribute(BillSheetColumns.SELLER_STATE)]

        SellerStatus,
        [DescriptionAttribute(BillSheetColumns.BILL_PRICE)]
        CurrentPrice,

        [Description(BillSheetColumns.CUSTOMER_PASSPORT_ID)]
        CustomerPassportId,

        [Description(BillSheetColumns.BILL_ID)] 
        BillID,

        [Description(BillSheetColumns.PAY_ADDRESS)]
        PAY_ADDRESS,

        [Description(BillSheetColumns.PAY_ADDRESS2)]
        PAY_ADDRESS2

    }
}