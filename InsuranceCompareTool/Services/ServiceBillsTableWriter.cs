using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsuranceCompareTool.Models;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services {
    public class ServiceBillsTableWriter
    {
        #region Fields
        public readonly SheetColumn[] Columns =
        {
            //new SheetColumn(){Title = BillSheetColumns.CURRENT_SERVICE_ID, Index = 0 , Width = 800*3}, 
            new SheetColumn
            {
                Title = BillSheetColumns.CURRENT_SERVICE_NAME,
                Index = 0,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.SELLER_NAME,
                Index = 1,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.SELLER_STATE,
                Index = 2,
                Width = 700 * 2
            },
            new SheetColumn
            {
                Title = BillSheetColumns.BILL_ID,
                Index = 3,
                Width = 700 * 6
            },

            //new SheetColumn(){Title = BillSheetColumns.PRODUCT_NAME, Index = 5, Width = 800*10},
            new SheetColumn
            {
                Title = BillSheetColumns.CUSTOMER_NAME,
                Index = 4,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_ADDRESS,
                Index = 5,
                Width = 700 * 9,
                
            },

            //new SheetColumn(){Title = BillSheetColumns.CUSTOMER_PASSPORT_ID, Index = 7, Width = 800*7, Alignment = HorizontalAlignment.Left},
            new SheetColumn
            {
                Title = BillSheetColumns.BILL_PRICE,
                Index = 6,
                Width = 700 * 3,
                Alignment = HorizontalAlignment.Right
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_DATE,
                Index = 7,
                Width =700 * 4,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PER_PAY,
                Index = 8,
                Width = 650 * 2,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_NO,
                Index = 9,
                Width = 650 * 2,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.MOBILE_PHONE,
                Index = 10,
                Width = 650 * 5
            },
            new SheetColumn
            {
                Title = BillSheetColumns.CREDITOR,
                Index = 11,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.CUSTOMER_ACCOUNT,
                Index = 12,
                Width = 700 * 7   + 300 ,
                Alignment = HorizontalAlignment.Left
            }
        };
        private readonly short DataBackgroundColor = IndexedColors.White.Index;
        private readonly short DataBorderColor = IndexedColors.Grey40Percent.Index;
        private readonly short DataTextColor = IndexedColors.Black.Index;
        private readonly short
            DocumentTitleBackgroundColor = IndexedColors.SeaGreen.Index; // IndexedColors.RoyalBlue.Index;
        private readonly short DocumentTitleTextColor = IndexedColors.White.Index;
 
        private readonly short ProjectBackgroundColor = IndexedColors.LightTurquoise.Index;
        private readonly short ProjectTextColor = IndexedColors.Black.Index;
        private readonly short TableTitleBackgroundColor = IndexedColors.Grey40Percent.Index;
        private readonly short TableTitleTextColor = IndexedColors.White.Index;
        #endregion
        #region Public Methods
        public void Save(string filename, string title, List<Bill> bills)
        {
            IWorkbook mWorkbook = new XSSFWorkbook();
            var mSheet = mWorkbook.CreateSheet("保单");
            SetPrint(mSheet);
           
            WriteTitleRow(title, mSheet);
            WriteHeaderRow(mSheet);
            SetColumnsStyle(mSheet);
            var style = GetRowStyle(mSheet.Workbook, RowType.DataRow);
            foreach(var bill in bills)
                WriteBillRow(bill, mSheet);
            var file = new FileStream(filename, FileMode.CreateNew, FileAccess.Write);
            mWorkbook.Write(file);
            file.Close();
            mWorkbook?.Close();
        }
        #endregion
        #region Private or Protect Methods
        private void SetPrint(ISheet sheet)
        {
            sheet.Header.Center = $"{DateTime.Now.ToString("yyyy年M月")}收费清单";
            sheet.PrintSetup.Landscape = true;
            sheet.PrintSetup.PaperSize = 9;
            sheet.SetMargin(MarginType.RightMargin, (double)0.2);
            sheet.SetMargin(MarginType.TopMargin, (double)0.7);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.2);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.2);
            sheet.RepeatingRows = new CellRangeAddress(1, 1, 0, Columns.Length);
        }

        private ICellStyle GetRowStyle(IWorkbook book, RowType rowType)
        {
            var rowstyle = book.CreateCellStyle();
            var font = book.CreateFont();
            rowstyle.FillPattern = FillPattern.SolidForeground;
            rowstyle.SetFont(font);
            switch(rowType)
            {
                case RowType.DocumentTitle:
                {
                    rowstyle.FillForegroundColor = DocumentTitleBackgroundColor;
                    font.FontHeight = 6 * 20 * 2.4;
                    font.Color = DocumentTitleTextColor;
                    font.IsBold = true;
                    break;
                }
                case RowType.TableTitle:
                {
                    font.FontHeight = 4.5 * 20 * 2.4;
                    font.Color = TableTitleTextColor;
                    rowstyle.FillForegroundColor = TableTitleBackgroundColor;
                    rowstyle.Alignment = HorizontalAlignment.Center;
                    rowstyle.VerticalAlignment = VerticalAlignment.Center;
                    rowstyle.WrapText = true;
                    break;
                }
                case RowType.ProjectTitle:
                {
                    font.FontHeight = 6 * 20 * 2.4;
                    font.Color = ProjectTextColor;
                    rowstyle.FillForegroundColor = ProjectBackgroundColor;
                    break;
                }
                case RowType.DataRow:
                {
                    font.FontHeight = 5 * 18 * 2.4;
                    font.Color = DataTextColor;
                    rowstyle.FillForegroundColor = DataBackgroundColor;
                    //rowstyle.BorderTop = BorderStyle.Thin;
                    rowstyle.BorderBottom = BorderStyle.Thin;
                    //rowstyle.BorderLeft = BorderStyle.Thin;
                    //rowstyle.BorderRight = BorderStyle.Thin;
                    rowstyle.LeftBorderColor = DataBorderColor;
                    rowstyle.BottomBorderColor = DataBorderColor;
                    rowstyle.RightBorderColor = DataBorderColor;
                    rowstyle.TopBorderColor = DataBorderColor;
                    rowstyle.WrapText = true;
                    rowstyle.VerticalAlignment = VerticalAlignment.Center;
                    
                    break;
                }
            }

            return rowstyle;
        }
        private void SetColumnsStyle(ISheet sheet)
        {
            foreach(var col in Columns)
            {
                var style = GetRowStyle(sheet.Workbook, RowType.DataRow);
                style.Alignment = col.Alignment; 
                sheet.SetDefaultColumnStyle(col.Index, style);
            }
        }
        private void WriteBillRow(Bill bill, ISheet sheet   )
        { 
            //todo writed by haiser
            if(bill.PayDate > new DateTime(2019, 9, 3))
                return;
            
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            foreach (var column in Columns)
            {
                row.CreateCell(column.Index).CellStyle =   sheet.GetColumnStyle(column.Index); 
            }

            row.Height = 200 * 3;
            if (!string.IsNullOrEmpty(bill.CurrentServiceName))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.CURRENT_SERVICE_NAME)).Index)
                    .SetCellValue(bill.CurrentServiceName);
            if(!string.IsNullOrEmpty(bill.ID))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.BILL_ID)).Index)
                    .SetCellValue(bill.ID);
            row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.PAY_NO)).Index).SetCellValue(bill.PayNo);
            row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.BILL_PRICE)).Index) .SetCellValue(bill.Price);
            if(!string.IsNullOrEmpty(bill.Creditor))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.CREDITOR)).Index) .SetCellValue(bill.Creditor); 
            if(!string.IsNullOrEmpty(bill.CustomerName))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.CUSTOMER_NAME)).Index) .SetCellValue(bill.CustomerName); 
            if(!string.IsNullOrEmpty(bill.CustomerAccount))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.CUSTOMER_ACCOUNT)).Index)
                    .SetCellValue(bill.CustomerAccount);
            if(!string.IsNullOrEmpty(bill.SellerName))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.SELLER_NAME)).Index)
                    .SetCellValue(bill.SellerName);
            if(!string.IsNullOrEmpty(bill.SellerState))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.SELLER_STATE)).Index)
                    .SetCellValue(bill.SellerState);
            if(!string.IsNullOrEmpty(bill.MobilePhone))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.MOBILE_PHONE)).Index) .SetCellValue(bill.MobilePhone.Replace("M:",""));
            if (!string.IsNullOrEmpty(bill.PerPay))
                row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.PER_PAY)).Index).SetCellValue(bill.PerPay);
            if(!string.IsNullOrEmpty(bill.PayAddress))
                row.GetCell(Columns.First(a=>a.Title.Equals(BillSheetColumns.PAY_ADDRESS)).Index)?.SetCellValue(bill.PayAddress);
            var payDateCol = row.GetCell(Columns.First(a => a.Title.Equals(BillSheetColumns.PAY_DATE)).Index);
            payDateCol.SetCellValue(bill.PayDate.ToString("yyyy/MM/dd")); 
           
        }
        private void WriteHeaderRow(ISheet sheet)
        {
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            var style = GetRowStyle(sheet.Workbook, RowType.TableTitle);
            foreach(var col in Columns)
            {
                var cell = row.CreateCell(col.Index, CellType.String);
                cell.CellStyle = style;
                cell.SetCellValue(col.Title);
            }
        }
        private void WriteTitleRow(string title, ISheet sheet)
        {
            var row = sheet.CreateRow(0);
            var style = GetRowStyle(sheet.Workbook, RowType.DocumentTitle);
            foreach(var col in Columns)
            {
                var cell = row.CreateCell(col.Index, CellType.String);
                sheet.SetColumnWidth(col.Index, col.Width);
                cell.CellStyle = style;
            }

            row.GetCell(row.FirstCellNum).SetCellValue(title);
            sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, row.FirstCellNum, row.LastCellNum - 1));
        }
        #endregion
    }
}