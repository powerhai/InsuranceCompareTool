using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Services;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Core {
    public class ServiceBillsTableWriterB
    {
        #region Fields
         
        private readonly short DataBackgroundColor = IndexedColors.White.Index;
        private readonly short DataBorderColor = IndexedColors.Grey40Percent.Index;
        private readonly short DataTextColor = IndexedColors.Black.Index;
        private readonly short
            DocumentTitleBackgroundColor = IndexedColors.SeaGreen.Index; // IndexedColors.RoyalBlue.Index;
        private readonly short DocumentTitleTextColor = IndexedColors.White.Index;
        private readonly short PageTitleBackgroundColor = IndexedColors.White.Index; // IndexedColors.RoyalBlue.Index;
        private readonly short PageTitleTextColor = IndexedColors.Black.Index;
        private readonly short ProjectBackgroundColor = IndexedColors.LightTurquoise.Index;
        private readonly short ProjectTextColor = IndexedColors.Black.Index;
        private readonly short TableTitleBackgroundColor = IndexedColors.Grey25Percent.Index;
        private readonly short TableTitleTextColor = IndexedColors.Black.Index;
        private readonly short ServiceTabColor = IndexedColors.Green.Index;
        private readonly short SellerTabColor = IndexedColors.Orange.Index;
        private readonly short ALLTabColor = IndexedColors.Blue.Index;

        private XSSFWorkbook mWorkbook;
        #endregion
        #region Public Methods
        public ServiceBillsTableWriterB()
        {
            mWorkbook = new XSSFWorkbook();
        }
        private List<string> GetSheets()
        {
            var items = new List<string>();
            for(var i = 0;i<  mWorkbook.NumberOfSheets; i ++)
            {
                var sheet = mWorkbook.GetSheetAt(i);
                items.Add(sheet.SheetName);
            }

            return items;
        }
        public void WriteBills(string name,string title,string area, List<Bill> bills ,  SheetTemplate template , string personID)
        {
            try
            {
                ISheet mSheet = null;

                if(mWorkbook.GetSheet(name) != null)
                {
                    name = name + "-" + personID;
                }
                mSheet =  mWorkbook.CreateSheet(name); 
                SetPrint(mSheet,template); 
                WritePageTitle(mSheet, $"{DateTime.Now.AddMonths(1).ToString("yyyy年M月")}收费清单 - {area}",  template);
                WriteHeaderRow(mSheet, template);
                var areaNameA = area + "市";
                var areaNameB = area + "县";
                foreach (var bill in bills)
                    WriteBillRow(bill, mSheet, template,areaNameA,areaNameB);
                WriteSumRow(title, mSheet, template);
            }
            catch (Exception e)
            { 
                throw new Exception($"写入表${name}时产生错误" + e.Message);
            } 
        }
        public void Save(string filename )
        {
            if(File.Exists(filename))
            {
                File.Delete(filename);
            }
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            mWorkbook.Write(file);
            file.Close();
            mWorkbook?.Close();  
        }
        #endregion
        #region Private or Protect Methods
        private void SetPrint(ISheet sheet, SheetTemplate template)
        {
            //sheet.Header.Center = $"{DateTime.Now.AddMonths(1).ToString("yyyy年M月")}收费清单";
            sheet.PrintSetup.Landscape = true;
            sheet.PrintSetup.PaperSize = 9;

            #region 设置打印 Fit to all columns on one page. 
            sheet.FitToPage = true;
            sheet.Autobreaks = true;
            sheet.PrintSetup.FitWidth = 1;
            sheet.PrintSetup.FitHeight = 0; 
            #endregion
          

            sheet.SetMargin(MarginType.RightMargin, (double)0.2);
            sheet.SetMargin(MarginType.TopMargin, (double)0.2);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.2);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.2);
            sheet.RepeatingRows = new CellRangeAddress(0, 1, 0, template.HeadRow.Cells.Count);

        }
         
 

        private void WritePageTitle(ISheet sheet,String title,  SheetTemplate template)
        { 
            var row = sheet.CreateRow(0);
            
            foreach (var col in template.HeadRow.Cells)
            {
                var cell = row.CreateCell(col.ColumnIndex, CellType.String);
                sheet.SetColumnWidth(col.ColumnIndex, template.Sheet.GetColumnWidth(col.ColumnIndex));
                cell.CellStyle = sheet.Workbook.CreateCellStyle();
                cell.CellStyle.CloneStyleFrom(template.TitleRow.GetCell(0).CellStyle);
            }

            row.Height = template.TitleRow.Height;
            row.GetCell(row.FirstCellNum).SetCellValue(title);
            sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, row.FirstCellNum, row.LastCellNum - 1));


        }
        private void WriteBillRow(Bill bill, ISheet sheet, SheetTemplate template,string areaNameA, string areaNameB)
        {

            try
            {
                
                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                row.RowStyle = sheet.Workbook.CreateCellStyle();
                row.RowStyle.CloneStyleFrom(template.DataRow.GetCell(0).CellStyle);

                row.Height = template.DataRow.Height;
                foreach (var column in template.HeadRow.Cells)
                { 
                    var cell =  row.CreateCell(column.ColumnIndex);
                    cell.CellStyle = row.RowStyle; 
                    string cellName = column.StringCellValue;
                    switch(cellName)
                    {
                        case BillSheetColumns.CURRENT_SERVICE_NAME:
                        { 
                            cell.SetCellValue(bill.CurrentServiceName??""); 
                            break;
                        }
                        case BillSheetColumns.BILL_ID:
                        { 
                            cell.SetCellValue(bill.ID??"");
                            break;
                        }
                        case BillSheetColumns.PAY_NO:
                        {
                            cell.SetCellValue(bill.PayNo);
                            break;
                        }
                        case BillSheetColumns.BILL_PRICE:
                        {
                            cell.SetCellValue(bill.Price);
                            break;
                        }
                        case BillSheetColumns.CREDITOR:
                        { 
                            cell.SetCellValue(bill.Creditor??"");
                            break;
                        }
                        case BillSheetColumns.CUSTOMER_NAME:
                        { 
                            cell.SetCellValue(bill.CustomerName??"");
                            break;
                        }
                        case BillSheetColumns.CUSTOMER_ACCOUNT:
                        {
                            cell.SetCellValue(bill.CustomerAccount??"");
                            break;
                        }
                        case BillSheetColumns.SELLER_NAME:
                        {
                            cell.SetCellValue(bill.SellerName ?? "");
                            break;
                        }
                        case BillSheetColumns.SELLER_STATE:
                        {
                            cell.SetCellValue(bill.SellerState ?? "");
                            break;
                        }
                        case BillSheetColumns.MOBILE_PHONE:
                        {
                            cell.SetCellValue(bill?.MobilePhone?.Replace("M:", "") ?? ""); 
                            break;
                        }
                        case BillSheetColumns.PER_PAY:
                        {
                            cell.SetCellValue(bill.PerPay ?? "");
                            break;
                        }
                        case BillSheetColumns.PAY_ADDRESS:
                        { 
                            cell.SetCellValue(bill.PayAddress?.Replace("浙江省", "").Replace("金华市", "").Replace(areaNameA,"").Replace(areaNameB,"")??""); 
                            break;
                        }
                        case BillSheetColumns.IS_OURS:
                        case BillSheetColumns.IS_OURS2:
                        { 
                            cell.SetCellValue(bill.IsOurs??""); 
                            break;
                        }
                        case BillSheetColumns.STATUS_OF_DJ:
                        {
                            if(!string.IsNullOrEmpty(bill.StatusOfDj))
                            {
                                bill.StatusOfDj = bill.StatusOfDj.Replace("督缴-业务员", "督缴");
                                cell.SetCellValue(bill.StatusOfDj);
                            }

                            break;
                        }
                        case BillSheetColumns.PAY_DATE:
                        { 
                            cell.SetCellValue(bill.PayDate.ToString("yyyy/MM/dd")??"");  
                            break;
                        }
                    } 
                } 
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }

        }
        private void WriteHeaderRow(ISheet sheet, SheetTemplate template)
        {
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            row.Height = template.HeadRow.Height;
 
            foreach (var col in template.HeadRow.Cells)
            {
                var cell = row.CreateCell(col.ColumnIndex, CellType.String);
                cell.CellStyle = sheet.Workbook.CreateCellStyle();
                cell.CellStyle.CloneStyleFrom(template.HeadRow.GetCell(0).CellStyle);
                
                cell.SetCellValue(col.StringCellValue);

            }
        }
        private void WriteSumRow(string title, ISheet sheet, SheetTemplate template)
        { 
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            var style = template.SumRow.GetCell( template.SumRow.FirstCellNum ).CellStyle;
          
            foreach (var col in template.HeadRow.Cells)
            {
                var cell = row.CreateCell(col.ColumnIndex, CellType.String);
                cell.CellStyle = sheet.Workbook.CreateCellStyle();
                cell.CellStyle.CloneStyleFrom(style);
            }
            row.GetCell(row.FirstCellNum).SetCellValue(title); 
            sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, row.FirstCellNum, row.LastCellNum - 1));
        }
        #endregion
    }
}