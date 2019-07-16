using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Domain
{
    public static class BillTableColumns
    {
        public static readonly ColumnDefine COL_SERVICE_STATUS = new ColumnDefine()
        {
            Name = BillSheetColumns.SYS_SERVICE_STATUS,
            Type = typeof(String),
            IsSystemColumn = false
        };
        public static readonly ColumnDefine COL_PREVIOUS_SERVICE_ID = new ColumnDefine()
        {
            Name = BillSheetColumns.PREVIOUS_SERVICE_ID,
            Type = typeof(String)
        };

        public static readonly ColumnDefine COL_CLIENT_AREA = new ColumnDefine()
        {
            Name = BillSheetColumns.CLIENT_AREA,
            Type = typeof(String),
            IsSystemColumn = false
        };

        public static readonly ColumnDefine COL_ASSIGNED_SERVICE_ID = new ColumnDefine()
        {
            Name = BillSheetColumns.ASSIGNED_SERVICE_ID,
            Type = typeof(String),
            IsSystemColumn = false
        };


        public static readonly ColumnDefine COL_SYS_SERVICE = new ColumnDefine()
        {
            Name = BillSheetColumns.SYS_SERVICE,
            Type = typeof(String),
            IsSystemColumn = true
        };
        public static readonly ColumnDefine COL_SYS_ERROR = new ColumnDefine()
        {
            Name = BillSheetColumns.SYS_ERROR,
            Type = typeof(String),
            IsSystemColumn = true
        };

        public static List<ColumnDefine> Columns = new List<ColumnDefine>()
        {
            new ColumnDefine(){ Name = BillSheetColumns.CURRENT_SERVICE_NAME , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CURRENT_SERVICE_ID , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.BILL_ID , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.BILL_ID2 , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CUSTOMER_NAME , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.PAY_DATE , Type = typeof(DateTime)}, 
            new ColumnDefine(){ Name = BillSheetColumns.PAY_DATE2 , Type = typeof(DateTime)}, 
            new ColumnDefine(){ Name = BillSheetColumns.SELLER_NAME , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.SELLER_ID , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.SELLER_STATE , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.SELL_MODEL , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CUSTOMER_ACCOUNT , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CUSTOMER_ADDRESS , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CUSTOMER_PASSPORT_ID , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CUSTOMER_PASSPORT_ID2 , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.BILL_PRICE , Type = typeof(double)}, 
            new ColumnDefine(){ Name = BillSheetColumns.PRODUCT_NAME , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.PER_PAY , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.MOBILE_PHONE , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.ASSIGNED_SERVICE_NAME , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.ASSIGNED_SERVICE_ID , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.SYS_FILTER , Type = typeof(String), IsSystemColumn = true}, 
            new ColumnDefine(){ Name = BillSheetColumns.SYS_FINISHED , Type = typeof(String), IsSystemColumn = true}, 
            new ColumnDefine(){ Name = BillSheetColumns.SYS_HISTORY , Type = typeof(String), IsSystemColumn = true},
            COL_SERVICE_STATUS,
            COL_CLIENT_AREA,
            COL_ASSIGNED_SERVICE_ID,
            COL_SYS_SERVICE,
            COL_SYS_ERROR,
            new ColumnDefine(){ Name = BillSheetColumns.PREVIOUS_PAY_DATE , Type = typeof(DateTime)},  
            new ColumnDefine(){ Name = BillSheetColumns.PREVIOUS_SERVICE , Type = typeof(String)},
            COL_PREVIOUS_SERVICE_ID,
            new ColumnDefine(){ Name = BillSheetColumns.IS_OURS , Type = typeof(String)},
            new ColumnDefine(){ Name = BillSheetColumns.IS_OURS2 , Type = typeof(String)},
            new ColumnDefine(){ Name = BillSheetColumns.IS_OURS3 , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.CREDITOR , Type = typeof(String)}, 
            new ColumnDefine(){ Name = BillSheetColumns.PAY_NO , Type = typeof(int)},  

        };
    }
}
