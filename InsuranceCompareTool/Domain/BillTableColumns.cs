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
        private const int SORT_SERVICE_NAME = 8000;
        private const int SORT_SERVICE_ID = SORT_SERVICE_NAME - 1;
        private const int SORT_SERVICE_ZT = SORT_SERVICE_ID - 1;
        private const int SORT_SERVICE_AREA = SORT_SERVICE_ZT - 1;

        private const int SORT_PREVIOUS_SERVICE = 7950;
        private const int SORT_PREVIOUS_SERVICE_ID = SORT_PREVIOUS_SERVICE - 1;


        private const int SORT_ASSIGNED_SERVICE_NAME = 7930; 
        private const int SORT_ASSIGNED_SERVICE_ID = SORT_ASSIGNED_SERVICE_NAME - 1;
         
        private const int SORT_SELLER_NAME = 7900;
        private const int SORT_SELLER_ID = SORT_SELLER_NAME - 1;
        private const int SORT_SELLER_ZT = SORT_SELLER_ID - 1;
        private const int SORT_SELLER_AREA = SORT_SELLER_ZT - 1;

        private const int SORT_CUSTOMER_NAME = 7800;
        private const int SORT_CUSTOMER_PASSPORT = SORT_CUSTOMER_NAME - 1;
        private const int SORT_CUSTOMER_AREA = SORT_CUSTOMER_PASSPORT - 1;

        private const int SORT_BILL_ID = 7700;
        private const int SORT_BILL_NAME = SORT_BILL_ID - 1;
        private const int SORT_BILL_PRICE = SORT_BILL_NAME - 1;
        private const int SORT_PAY_DATE = SORT_BILL_PRICE - 1;
        private const int SORT_PAY_NO = SORT_PAY_DATE - 1;
        private const int SORT_ADDRESS_A = SORT_PAY_NO - 1;
        private const int SORT_ADDRESS_B = SORT_ADDRESS_A - 1;

          
        public static readonly ColumnDefine COL_SERVICE_STATUS = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_SERVICE_STATUS },
            Type = typeof(String),
            IsSystemColumn = false,
            Sort = SORT_SERVICE_ZT
        };
        public static  readonly  ColumnDefine COL_SERVICE_AREA = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_SERVICE_AREA },
            Type = typeof(String),
            IsSystemColumn = false,
            IsRequired = true,
            Sort = SORT_SERVICE_AREA
        };
        
        public static readonly ColumnDefine COL_PREVIOUS_SERVICE_ID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.PREVIOUS_SERVICE_ID },
            Type = typeof(String),
            IsRequired = true, Sort = SORT_PREVIOUS_SERVICE_ID,
            IsCopyAble = true
        };

        public static readonly ColumnDefine COL_CUSTOMER_AREA = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.CLIENT_AREA },
            Type = typeof(String),
            IsSystemColumn = false,
            Sort = SORT_CUSTOMER_AREA
        };



        public static readonly ColumnDefine COL_ASSIGNED_SERVICE_NAME = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.ASSIGNED_SERVICE_NAME },
            Type = typeof(String),
            IsSystemColumn = false,
            Sort = SORT_ASSIGNED_SERVICE_NAME,
            IsCopyAble = true
        };


        public static readonly ColumnDefine COL_ASSIGNED_SERVICE_ID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.ASSIGNED_SERVICE_ID },
            Type = typeof(String),
            IsSystemColumn = false,
            Sort = SORT_ASSIGNED_SERVICE_ID,
            IsCopyAble = true
        };


        public static readonly ColumnDefine COL_SYS_SERVICE = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_SERVICE },
            Type = typeof(String),
            IsSystemColumn = true,
            IsRequired = true,
            IsCopyAble = true
        };
        public static readonly ColumnDefine COL_SYS_ERROR = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_ERROR },
            Type = typeof(String),
            IsSystemColumn = true,
            IsRequired = true,
        };

        public static readonly ColumnDefine COL_SYS_SERVICE_ID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_SERVICE_ID },
            Type = typeof(String),
            IsSystemColumn = true,
            IsRequired = true,
            IsCopyAble = true
        };

        public static readonly ColumnDefine COL_CURRENT_SERVICE_ID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.CURRENT_SERVICE_ID },
            Type = typeof(String),
            IsRequired = true,
            ExportName = "工号",
            Sort = SORT_SERVICE_ID,
            IsCopyAble = true
        };
        public static readonly ColumnDefine COL_PAY_DATE = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.PAY_DATE, BillSheetColumns.PAY_DATE2 },
            Type = typeof(DateTime),
            ExportName = "应缴日期(YYYY-MM-DD)", 
        };

        public static readonly ColumnDefine COL_BILL_ID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.BILL_ID, BillSheetColumns.BILL_ID2 },
            Type = typeof(String),
            ExportName = BillSheetColumns.BILL_ID,
            Sort = SORT_BILL_ID,
            IsCopyAble = true
        };

        public static readonly ColumnDefine COL_GUID = new ColumnDefine()
        {
            Name = new string[] { BillSheetColumns.SYS_GUID },
            Type = typeof(Guid),
            IsSystemColumn = true,
            IsRequired = true 
        };


        public static List<ColumnDefine> Columns = new List<ColumnDefine>()
        {
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.CURRENT_SERVICE_NAME} , Type = typeof(String), Sort = SORT_SERVICE_NAME, IsCopyAble = true},
            COL_CURRENT_SERVICE_ID,
            COL_SERVICE_AREA,
            COL_BILL_ID,
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.CUSTOMER_NAME} , Type = typeof(String), Sort = SORT_CUSTOMER_NAME , IsCopyAble = true},
            COL_PAY_DATE,
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SELLER_NAME} , Type = typeof(String) , Sort =  SORT_SELLER_NAME, IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SELLER_ID} , Type = typeof(String), Sort = SORT_SELLER_ID, IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SELLER_STATE } , Type = typeof(String), Sort = SORT_SELLER_ZT},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SYS_SELLER_AREA } , Type = typeof(String), Sort = SORT_SELLER_AREA, IsRequired = true },


            new ColumnDefine(){ Name = new string[]{BillSheetColumns.SELL_MODEL} , Type = typeof(String)},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.CUSTOMER_ACCOUNT }, Type = typeof(String)},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.CUSTOMER_ADDRESS, BillSheetColumns.CUSTOMER_ADDRESS, BillSheetColumns.PAY_ADDRESS2, BillSheetColumns.PAY_ADDRESS3  }, Type = typeof(String), Sort = SORT_ADDRESS_A, IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.CUSTOMER_ADDRESS, BillSheetColumns.PAY_ADDRESS4, BillSheetColumns.PAY_ADDRESS5 }, Type = typeof(String), Sort = SORT_ADDRESS_B, IsCopyAble = true},

            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.CUSTOMER_PASSPORT_ID, BillSheetColumns.CUSTOMER_PASSPORT_ID2 } , Type = typeof(String) , Sort =  SORT_CUSTOMER_PASSPORT, IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.BILL_PRICE } , Type = typeof(double), Sort = SORT_BILL_PRICE},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.PRODUCT_NAME }, Type = typeof(String), Sort = SORT_BILL_NAME, IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.PER_PAY } , Type = typeof(String)},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.MOBILE_PHONE }, Type = typeof(String), IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.ASSIGNED_SERVICE_NAME} , Type = typeof(String), IsCopyAble = true},
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.ASSIGNED_SERVICE_ID} , Type = typeof(String),IsCopyAble = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SYS_FILTER} , Type = typeof(String), IsSystemColumn = true, IsRequired = true},
            new ColumnDefine(){ Name =new string[]{ BillSheetColumns.SYS_FINISHED} , Type = typeof(String), IsSystemColumn = true, IsRequired = true},
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.SYS_HISTORY} , Type = typeof(String), IsSystemColumn = true, IsRequired = true},
            COL_SERVICE_STATUS,
            COL_CUSTOMER_AREA,
            COL_ASSIGNED_SERVICE_NAME,
            COL_ASSIGNED_SERVICE_ID,
            COL_SYS_SERVICE,
            COL_SYS_SERVICE_ID,
            COL_SYS_ERROR,
            new ColumnDefine(){ Name =new string[]{BillSheetColumns.PREVIOUS_PAY_DATE} , Type = typeof(DateTime), Sort = SORT_PAY_DATE},
            new ColumnDefine(){ Name =new string[]{BillSheetColumns.PREVIOUS_SERVICE} , Type = typeof(String) , Sort = SORT_PREVIOUS_SERVICE, IsCopyAble = true},
            COL_PREVIOUS_SERVICE_ID,
            new ColumnDefine(){ Name =new string[]{BillSheetColumns.IS_OURS, BillSheetColumns.IS_OURS2, BillSheetColumns.IS_OURS3 }, Type = typeof(String)},
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.CREDITOR} , Type = typeof(String), IsCopyAble = true},
            new ColumnDefine(){ Name = new string[]{BillSheetColumns.PAY_NO } , Type = typeof(int) , Sort = SORT_PAY_NO },
            COL_GUID
        };
    }
}
