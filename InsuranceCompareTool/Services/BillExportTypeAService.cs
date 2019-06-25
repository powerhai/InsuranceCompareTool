using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.ShareCommon;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class BillExportTypeAService
    {
        private static BillExportTypeAService Singleton = null;
        public static BillExportTypeAService CreateInstance()
        {
            if (Singleton == null)
            {
                Singleton = new BillExportTypeAService();
            }
            return Singleton;
        }

        public string Export(string  sourceFile, string targetFile, List<Bill> bills)
        { 
            //
            var tempFile = "abc.xlsx";
            File.Copy(sourceFile,tempFile, true);

            IWorkbook tarExcel = new XSSFWorkbook(tempFile);
            ISheet tarSheet = tarExcel.GetSheetAt(0);
            SheetReader sheetReader = new SheetReader(tarSheet); 
             
            AddColumns(tarSheet);

            var columns = sheetReader.GetColumns();
            
            foreach (var bill in bills)
            {
                CopyRow(bill, tarSheet, columns);
            }
            File.Delete(targetFile);
            var file = new FileStream(targetFile, FileMode.CreateNew, FileAccess.Write);
            tarExcel.Write(file);
            file.Close();
             
            tarExcel.Close();
            
            return "";
        }
        private void CopyHeaderRow(IRow row, ISheet tarSheet, ISheet srcSheet)
        {
            var trow = tarSheet.CreateRow(0);
            foreach (ICell cell in row.Cells)
            {
                var ncell = trow.CreateCell(cell.ColumnIndex);
                ncell.CellStyle = tarSheet.Workbook.CreateCellStyle();
                ncell.CellStyle.CloneStyleFrom(cell.CellStyle);
                ncell.SetCellType(cell.CellType);

                switch (cell.CellType)
                {
                    case CellType.String:
                    {
                        ncell.SetCellValue(cell.StringCellValue);
                        break;
                    }
                    case CellType.Boolean:
                    {
                        ncell.SetCellValue(cell.BooleanCellValue);
                        break;
                    }
                    case CellType.Error:
                    {
                        ncell.SetCellValue(cell.ErrorCellValue);
                        break;
                    }
                    case CellType.Numeric:
                    {
                        ncell.SetCellValue(cell.NumericCellValue);
                        break;
                    }
                    default:
                    {
                        ncell.SetCellValue(cell.StringCellValue);
                        break;
                    }
                }
                tarSheet.SetColumnWidth(cell.ColumnIndex, row.Sheet.GetColumnWidth(cell.ColumnIndex)); 
            }
        }
        private void AddColumns(ISheet tarSheet)
        {
            var row = tarSheet.GetRow(tarSheet.FirstRowNum);
 
            //row.CreateCell(row.LastCellNum).SetCellValue( BillStatus.ServiceNotAvailable.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.Yiwu.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.NoService.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.NoLastService.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.ServiceSameToSeller.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.BindService.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.AcrossArea234.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.AcrossArea5.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.AreaSell.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.SellerNotAvailable.GetDescription());
            //row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.Error.GetDescription());
            row.CreateCell(row.LastCellNum).SetCellValue(BillStatus.DifferentService.GetDescription());
            row.CreateCell(row.LastCellNum).SetCellValue("原客服");
            row.CreateCell(row.LastCellNum).SetCellValue("现客服");
            row.CreateCell(row.LastCellNum).SetCellValue("身份证");

        }
        private void CopyRow(Bill bill, ISheet tarSheet, List<SheetColumn> columns)
        {
            try
            { 
                var trow = tarSheet.GetRow(bill.RowNum);


                var curSevID = columns.FirstOrDefault(a => a.Title.Equals(BillSheetColumns.CURRENT_SERVICE_ID));
                var curSevName = columns.FirstOrDefault(a => a.Title.Equals(BillSheetColumns.CURRENT_SERVICE_NAME));
                //if (bill.CurrentServiceObj != null && bill.LastServiceObj != null && !bill.CurrentServiceObj.ID.Equals(bill.LastServiceObj.ID))
                //{
                //    trow.GetCell(curSevID.Index)?.SetCellValue(bill.CurrentServiceID);
                //    trow.GetCell(curSevName.Index)?.SetCellValue(bill.CurrentServiceName);
                //}

                trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.DifferentService) ? "Y" : "");
               
                if(bill.Statuses.Contains(BillStatus.DifferentService))
                {
                        trow.GetCell(curSevID.Index)?.SetCellValue(bill.CurrentServiceID);
                        trow.GetCell(curSevName.Index)?.SetCellValue(bill.CurrentServiceName);

                    trow.CreateCell(trow.LastCellNum).SetCellValue(bill.SrcServiceName);
                    trow.CreateCell(trow.LastCellNum).SetCellValue(bill.CurrentServiceName);
                    trow.CreateCell(trow.LastCellNum).SetCellValue(bill.CustomerPassportID);
                }
                //var lastServiceIdCell = trow.CreateCell(trow.LastCellNum);
                //lastServiceIdCell.SetCellValue(bill.LastServiceID);
                //var lastServiceName = trow.CreateCell(trow.LastCellNum);
                //lastServiceName.SetCellValue(bill.LastServiceName);
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.ServiceNotAvailable) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.Yiwu) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.NoService) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.NoLastService) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.ServiceSameToSeller) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.BindService) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.AcrossArea234) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.AcrossArea5) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.AreaSell) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.SellerNotAvailable) ? "Y" : "");
                //trow.CreateCell(trow.LastCellNum).SetCellValue(bill.Statuses.Contains(BillStatus.Error) ? "Y" : "");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            } 
        } 
    }
}
