using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Services;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Core {
    public class ServiceBillsTableWriterB
    {
        #region Fields

        public readonly SheetColumn[] ColumnsA =
        {
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
        public readonly SheetColumn[] ColumnsB =
        { 
           
            new SheetColumn
            {
                Title = BillSheetColumns.SELLER_NAME,
                Index = 0,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.SELLER_STATE,
                Index = 1,
                Width = 900 * 2
            },
            new SheetColumn
            {
                Title = BillSheetColumns.BILL_ID,
                Index = 2,
                Width = 700 * 6
            },
             
            new SheetColumn
            {
                Title = BillSheetColumns.CUSTOMER_NAME,
                Index = 3,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_ADDRESS,
                Index = 4,
                Width = 1200 * 9,

            },
             
            new SheetColumn
            {
                Title = BillSheetColumns.BILL_PRICE,
                Index = 5,
                Width = 700 * 3,
                Alignment = HorizontalAlignment.Right
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_DATE,
                Index = 6,
                Width =700 * 4,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PER_PAY,
                Index = 7,
                Width = 650 * 2,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.PAY_NO,
                Index = 8,
                Width = 650 * 2,
                Alignment = HorizontalAlignment.Center
            },
            new SheetColumn
            {
                Title = BillSheetColumns.MOBILE_PHONE,
                Index = 9,
                Width = 650 * 5
            },
            new SheetColumn
            {
                Title = BillSheetColumns.CREDITOR,
                Index = 10,
                Width = 700 * 3
            },
            new SheetColumn
            {
                Title = BillSheetColumns.CUSTOMER_ACCOUNT,
                Index = 11,
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
        public void WriteBills(string name,string title,string area, List<Bill> bills , WriteType writeType  )
        {
            try
            {
                ISheet mSheet = null;
                var sheets = GetSheets();
                switch (writeType)
                {
                    case WriteType.All:
                    {
                        //mSheet.TabColorIndex = ALLTabColor;
                        break;
                    }
                    case WriteType.Virtual:
                    case WriteType.Seller:
                    {
                        //mSheet.TabColorIndex = SellerTabColor;
                        name = "@ " + name ;
                        break;
                    }
                    case WriteType.Service:
                    {
                        //mSheet.TabColorIndex = ServiceTabColor;
                        break;
                    }
                }
                mSheet =  mWorkbook.CreateSheet(name);
                
           
                SetPrint(mSheet,writeType);

                WritePageTitle(mSheet, $"{DateTime.Now.AddMonths(1).ToString("yyyy年M月")}收费清单 - {area}",  writeType);
                WriteHeaderRow(mSheet, writeType);
                SetColumnsStyle(mSheet, writeType);
                var style = GetRowStyle(mSheet.Workbook, RowType.DataRow);
                foreach (var bill in bills)
                    WriteBillRow(bill, mSheet, writeType);
                WriteTitleRow(title, mSheet, writeType);
            }
            catch (Exception e)
            { 
                throw;
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
        private void SetPrint(ISheet sheet, WriteType writeType)
        {
            //sheet.Header.Center = $"{DateTime.Now.AddMonths(1).ToString("yyyy年M月")}收费清单";
            sheet.PrintSetup.Landscape = true;
            sheet.PrintSetup.PaperSize = 9;
            sheet.SetMargin(MarginType.RightMargin, (double)0.2);
            sheet.SetMargin(MarginType.TopMargin, (double)0.2);
            sheet.SetMargin(MarginType.LeftMargin, (double)0.2);
            sheet.SetMargin(MarginType.BottomMargin, (double)0.2);
            sheet.RepeatingRows = new CellRangeAddress(0, 1, 0, writeType == WriteType.Virtual ?  ColumnsB.Length : ColumnsA.Length);
        }

        private ICellStyle GetRowStyle(IWorkbook book, RowType rowType)
        {
            var rowstyle = book.CreateCellStyle();
            var font = book.CreateFont();
            rowstyle.FillPattern = FillPattern.SolidForeground;
            rowstyle.SetFont(font);
            switch (rowType)
            {
                case RowType.PageTitle:
                {
                    rowstyle.FillForegroundColor = PageTitleBackgroundColor;
                    font.FontHeight = 7 * 20 * 2.4;
                    font.Color = PageTitleTextColor;
                    font.IsBold = true;
                    font.FontName = "宋体";
                    rowstyle.Alignment = HorizontalAlignment.Center;
                    rowstyle.VerticalAlignment = VerticalAlignment.Center;
                    break;
                }
                case RowType.DocumentTitle:
                {
                    rowstyle.FillForegroundColor = DocumentTitleBackgroundColor;
                    font.FontHeight = 5 * 20 * 2.4;
                    font.Color = DocumentTitleTextColor;
                    font.IsBold = false;
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
                    font.FontName = "宋体";
                    font.IsBold = true;
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
                    font.FontName = "宋体";
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
        private void SetColumnsStyle(ISheet sheet, WriteType writeType)
        {
            var columns = writeType == WriteType.Virtual ? ColumnsB : ColumnsA;
            foreach (var col in columns)
            {
                var style = GetRowStyle(sheet.Workbook, RowType.DataRow);
                style.Alignment = col.Alignment;
                sheet.SetDefaultColumnStyle(col.Index, style);
            }
        }

        private void WritePageTitle(ISheet sheet,String title, WriteType writeType)
        { 
            var row = sheet.CreateRow(0);
            var style = GetRowStyle(sheet.Workbook, RowType.PageTitle);
            var columns = writeType == WriteType.Virtual ? ColumnsB : ColumnsA;
            foreach (var col in columns)
            {
                var cell = row.CreateCell(col.Index, CellType.String);
                sheet.SetColumnWidth(col.Index, col.Width);
                cell.CellStyle = style;
            }

            row.Height = 200 * 6;
            row.GetCell(row.FirstCellNum).SetCellValue(title);
            sheet.AddMergedRegion(new CellRangeAddress(row.RowNum, row.RowNum, row.FirstCellNum, row.LastCellNum - 1));


        }
        private void WriteBillRow(Bill bill, ISheet sheet, WriteType writeType)
        {
            //todo writed by haiser
            if (bill.PayDate > new DateTime(2019, 9, 3))
                return;
            try
            {
                var columns = writeType == WriteType.Virtual ? ColumnsB : ColumnsA;
                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                foreach (var column in columns)
                { 
                   var cell =  row.CreateCell(column.Index);
                    var style = sheet.GetColumnStyle(column.Index);
                    if(style!= null)
                        cell.CellStyle = style;
                }

                row.Height = 200 * 3;
                if(writeType != WriteType.Virtual)
                {
                    if (!string.IsNullOrEmpty(bill.CurrentServiceName))
                        row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.CURRENT_SERVICE_NAME)).Index)
                            .SetCellValue(bill.CurrentServiceName);
                }

                if (!string.IsNullOrEmpty(bill.ID))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.BILL_ID)).Index)
                        .SetCellValue(bill.ID);
                row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.PAY_NO)).Index).SetCellValue(bill.PayNo);
                row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.BILL_PRICE)).Index).SetCellValue(bill.Price);
                if (!string.IsNullOrEmpty(bill.Creditor))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.CREDITOR)).Index).SetCellValue(bill.Creditor);
                if (!string.IsNullOrEmpty(bill.CustomerName))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.CUSTOMER_NAME)).Index).SetCellValue(bill.CustomerName);
                if (!string.IsNullOrEmpty(bill.CustomerAccount))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.CUSTOMER_ACCOUNT)).Index)
                        .SetCellValue(bill.CustomerAccount);
                if (!string.IsNullOrEmpty(bill.SellerName))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.SELLER_NAME)).Index)
                        .SetCellValue(bill.SellerName);
                if (!string.IsNullOrEmpty(bill.SellerState))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.SELLER_STATE)).Index)
                        .SetCellValue(bill.SellerState);
                if (!string.IsNullOrEmpty(bill.MobilePhone))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.MOBILE_PHONE)).Index).SetCellValue(bill.MobilePhone.Replace("M:", ""));
                if (!string.IsNullOrEmpty(bill.PerPay))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.PER_PAY)).Index).SetCellValue(bill.PerPay);
                if (!string.IsNullOrEmpty(bill.PayAddress))
                    row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.PAY_ADDRESS)).Index)?.SetCellValue(bill.PayAddress.Replace("浙江省","").Replace("金华市",""));
                var payDateCol = row.GetCell(columns.First(a => a.Title.Equals(BillSheetColumns.PAY_DATE)).Index);
                payDateCol.SetCellValue(bill.PayDate.ToString("yyyy/MM/dd"));
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }


        }
        private void WriteHeaderRow(ISheet sheet, WriteType writeType)
        {
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            var style = GetRowStyle(sheet.Workbook, RowType.TableTitle);
            var columns = writeType == WriteType.Virtual ? ColumnsB : ColumnsA;
            foreach (var col in columns)
            {
                var cell = row.CreateCell(col.Index, CellType.String);
                cell.CellStyle = style;
                cell.SetCellValue(col.Title);
            }
        }
        private void WriteTitleRow(string title, ISheet sheet, WriteType writeType)
        {
            var row = sheet.CreateRow(sheet.LastRowNum + 1);
            var style = GetRowStyle(sheet.Workbook, RowType.DocumentTitle);
            var columns = writeType == WriteType.Virtual ? ColumnsB : ColumnsA;
            foreach (var col in columns)
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