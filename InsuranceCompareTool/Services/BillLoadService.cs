using System;
using System.Collections.Generic;
using System.IO;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services {
    public class BillLoadService
    {
        private static BillLoadService Singleton = null;
        public static BillLoadService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillLoadService();
            }
            return Singleton;
        }

        public void Load(string file)
        {
            string tempFile = Path.GetTempFileName(); 
            File.Copy(file, tempFile, true);
            IWorkbook excel = new XSSFWorkbook(tempFile);

            try
            { 
                if(excel.NumberOfSheets <= 0)
                {
                    throw new Exception("保单数据表文件缺少数据");
                }

                ISheet sheet = excel.GetSheetAt(0);
                SheetReader sheetReader = new SheetReader(sheet);
                var columns = sheetReader.GetColumns();
                if(columns.Count <= 0)
                {
                    throw new Exception("保单数据表文件缺少表头数据");
                }

                InitColumnsData(columns);
                ValidateColumns();
                mBills = new List<Bill>();

                for(int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    var bill = GetBill(row);
                    if(bill != null)
                    {
                        mBills.Add(bill);
                    }
                }


            }
            catch(Exception ex)
            {
                throw;
            }
            finally
            {
                excel.Close();
            }
           
        }

        private void ValidateColumns()
        {
            List<string> mis = new List<string>();
            if(mColID.Index<0)
                mis.Add(BillSheetColumns.BILL_ID);
            //if(mColLastServiceID.Index <0)
            //    mis.Add(BillSheetColumns.LAST_SERVICE_ID);
            //if(mColLastServiceName.Index <0)
            //    mis.Add(BillSheetColumns.LAST_SERVICE_NAME);
            if (mColCurrentServiceID.Index < 0)
                mis.Add(BillSheetColumns.CURRENT_SERVICE_ID);
            if (mColCurrentServiceName.Index < 0)
                mis.Add(BillSheetColumns.CURRENT_SERVICE_NAME);
            if (mColCustomerName.Index <0)
                mis.Add(BillSheetColumns.CUSTOMER_NAME);
            if(mColPayDate.Index < 0 )
                mis.Add(BillSheetColumns.PAY_DATE);
            if(mColSellerID.Index < 0)
                mis.Add( BillSheetColumns.SELLER_ID);
            if(mColSellerName.Index <0)
                mis.Add(BillSheetColumns.SELLER_NAME);
            if(mColPayNo.Index <0)
                mis.Add(BillSheetColumns.PAY_NO);
            if(mColSellModel.Index <0)
                mis.Add(BillSheetColumns.SELL_MODEL);
            if(mColSellerState.Index <0)
                mis.Add(BillSheetColumns.SELLER_STATE);
            if(mColCustomerPassportID.Index < 0)
                mis.Add(BillSheetColumns.CUSTOMER_PASSPORT_ID);
            if(mis.Count > 0)
            {
                string v = "保单表缺少以下列： " + string.Join(", ", mis);
                throw new Exception(v);
            }
        }

        private List<Bill> mBills;
        public List<Bill> GetBills()
        {
            return mBills;
        }
        private void InitColumnsData(List<SheetColumn> columns)
        {
            foreach(var col in columns)
            {
                switch(col.Title)
                {
                    case BillSheetColumns.BILL_ID :
                    {
                        mColID.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CUSTOMER_PASSPORT_ID:
                    case BillSheetColumns.CUSTOMER_PASSPORT_ID2:
                    {
                        mColCustomerPassportID.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CUSTOMER_NAME:
                    {
                        mColCustomerName.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CUSTOMER_ADDRESS:
                    {
                        mColCustomerAddress.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.PAY_ADDRESS:
                    case BillSheetColumns.PAY_ADDRESS2:
                    {
                        mColPayAddress.Index = col.Index;
                        break;
                    }
                    
                    case BillSheetColumns.CURRENT_SERVICE_ID:
                    {
                        mColCurrentServiceID.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CURRENT_SERVICE_NAME:
                    {
                        mColCurrentServiceName.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.SELLER_ID:
                    {
                        mColSellerID.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.SELLER_NAME:
                    {
                        mColSellerName.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.PAY_NO:
                    {
                        mColPayNo.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.PAY_DATE:
                    {
                        mColPayDate.Index = col.Index;
                        break;
                    }
                    
                    case BillSheetColumns.SELL_MODEL:
                    {
                        mColSellModel.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.SELLER_STATE:
                    {
                        mColSellerState.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.ORGANIZATION4:
                    {
                        mColOrganization4.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.BILL_PRICE:
                    {
                        mColPrice.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.PRODUCT_NAME:
                    {
                        mColProductName.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CREDITOR:
                    {
                        mColCreditor.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.PER_PAY:
                    {
                        mColPerPay.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.CUSTOMER_ACCOUNT:
                    {
                        mColAccount.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.MOBILE_PHONE:
                    {
                        mColMobilePhone.Index = col.Index;
                        break;
                    }
                    case BillSheetColumns.IS_OURS:
                    {
                        mColIsOurs.Index = col.Index;
                        break;
                    }
                } 
            }
        }

        private Bill GetBill(IRow row)
        {
            var bill = new Bill();
              
            try
            {
                bill.PayDate = row.GetCell(mColPayDate.Index).DateCellValue ;
                bill.ID = row.GetCell(mColID.Index)?.StringCellValue.Trim(); 
                bill.CurrentServiceID = row.GetCell(mColCurrentServiceID.Index)?.StringCellValue.Trim();
                bill.CurrentServiceName = row.GetCell(mColCurrentServiceName.Index)?.StringCellValue.Trim(); 
                bill.CustomerPassportID = row.GetCell(mColCustomerPassportID.Index)?.StringCellValue.Trim(); 
                bill.CustomerName = row.GetCell(mColCustomerName.Index)?.StringCellValue.Trim(); 
                bill.SellerID = row.GetCell(mColSellerID.Index)?.StringCellValue.Trim(); 
                bill.SellerName = row.GetCell(mColSellerName.Index)?.StringCellValue.Trim();  
                bill.PayNo = Convert.ToInt32( row.GetCell(mColPayNo.Index).NumericCellValue); 
               
                bill.PayModel = row.GetCell(mColSellModel.Index)?.StringCellValue.Trim();
                bill.SellerState = row.GetCell(mColSellerState.Index)?.StringCellValue.Trim();
                bill.Organization4 = row.GetCell(mColOrganization4.Index)?.StringCellValue.Trim();
                bill.Price = row.GetCell(mColPrice.Index) == null ? 0 : row.GetCell(mColPrice.Index).NumericCellValue;
                bill.RowNum = row.RowNum;
                if (mColProductName.Index >= 0)
                    bill.ProductName = row.GetCell(mColProductName.Index)?.StringCellValue.Trim();
                if (mColPerPay.Index >= 0)
                    bill.PerPay = row.GetCell(mColPerPay.Index)?.StringCellValue.Trim();
                if (mColCreditor.Index >= 0)
                    bill.Creditor = row.GetCell(mColCreditor.Index)?.StringCellValue.Trim();
                if(mColAccount.Index >= 0)
                    bill.CustomerAccount = row.GetCell(mColAccount.Index)?.StringCellValue.Trim();
                if(mColMobilePhone.Index >= 0)
                    bill.MobilePhone = row.GetCell(mColMobilePhone.Index)?.StringCellValue.Trim();
                if(mColPayAddress.Index >= 0)
                    bill.PayAddress = row.GetCell(mColPayAddress.Index)?.StringCellValue.Trim();
                if(mColIsOurs.Index >= 0)
                {
                    bill.IsOurs = row.GetCell(mColIsOurs.Index)?.StringCellValue.Trim();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }

            return bill;
        }
         
        private readonly SheetColumn mColOrganization4 = new SheetColumn() { Title = BillSheetColumns.ORGANIZATION4 };  
        private readonly SheetColumn mColID = new SheetColumn(){ Title = BillSheetColumns.BILL_ID};
        private readonly SheetColumn mColCustomerPassportID = new SheetColumn(){ Title = BillSheetColumns.CUSTOMER_PASSPORT_ID };
        private readonly SheetColumn mColCustomerName = new SheetColumn(){ Title = BillSheetColumns.CUSTOMER_NAME  };
        private readonly SheetColumn mColCustomerAddress = new SheetColumn() { Title =  BillSheetColumns.CUSTOMER_ADDRESS };
        private readonly SheetColumn mColPayAddress = new SheetColumn() { Title = BillSheetColumns.PAY_ADDRESS };  
        private readonly SheetColumn mColCurrentServiceID = new SheetColumn() { Title = BillSheetColumns.CURRENT_SERVICE_ID };
        private readonly SheetColumn mColCurrentServiceName = new SheetColumn() { Title = BillSheetColumns.CURRENT_SERVICE_NAME };
        private readonly SheetColumn mColSellerID = new SheetColumn() { Title = BillSheetColumns.SELLER_ID };
        private readonly SheetColumn mColSellerName = new SheetColumn() { Title = BillSheetColumns.SELLER_NAME };
        private readonly SheetColumn mColPayNo = new SheetColumn() { Title = BillSheetColumns.PAY_NO };
        private readonly SheetColumn mColPayDate = new SheetColumn() { Title = BillSheetColumns.PAY_DATE , CellType = CellType.Numeric  }; 
        private readonly SheetColumn mColSellModel = new SheetColumn(){ Title = BillSheetColumns.SELL_MODEL};
        private readonly SheetColumn mColSellerState = new SheetColumn(){Title = BillSheetColumns.SELLER_STATE};
        private readonly SheetColumn mColPrice = new SheetColumn(){ Title = BillSheetColumns.BILL_PRICE, CellType = CellType.Numeric};
        private readonly SheetColumn mColProductName = new SheetColumn(){ Title =  BillSheetColumns.PRODUCT_NAME};
        private readonly SheetColumn mColPerPay = new SheetColumn(){ Title =  BillSheetColumns.PER_PAY};
        private readonly SheetColumn mColCreditor = new SheetColumn(){ Title =  BillSheetColumns.CREDITOR};
        private readonly SheetColumn mColAccount = new SheetColumn(){ Title =  BillSheetColumns.CUSTOMER_ACCOUNT};
        private readonly SheetColumn mColMobilePhone = new SheetColumn(){ Title =  BillSheetColumns.MOBILE_PHONE};
        private readonly SheetColumn mColIsOurs = new SheetColumn(){ Title = BillSheetColumns.IS_OURS};

    }
}